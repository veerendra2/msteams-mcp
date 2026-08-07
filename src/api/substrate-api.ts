/**
 * Substrate API client for search and people operations.
 * 
 * Handles all calls to substrate.office.com endpoints.
 */

import { httpRequest } from '../utils/http.js';
import { SUBSTRATE_API, getBearerHeaders } from '../utils/api-config.js';
import { type Result, ok } from '../types/result.js';
import { requireSubstrateTokenAsync, getTenantId, getTeamsBaseUrl, handleSubstrateError } from '../utils/auth-guards.js';
import { getValidSubstrateToken } from '../auth/token-extractor.js';
import {
  parseSearchResults,
  parseEmailSearchResults,
  parsePeopleResults,
  parseChannelResults,
  filterChannelsByName,
  type PersonSearchResult,
  type ChannelSearchResult,
  type LinkContext,
} from '../utils/parsers.js';
import { getMyTeamsAndChannels } from './csa-api.js';
import { getConversations } from './chatsvc-conversations.js';
import { getConversationProperties } from './chatsvc-messaging.js';
import type { TeamsSearchResult, SearchPaginationResult, EmailSearchResult } from '../types/teams.js';

/** Search result with pagination. */
export interface SearchResult {
  results: TeamsSearchResult[];
  pagination: SearchPaginationResult;
}

/** Email search result with pagination. */
export interface EmailSearchResultSet {
  results: EmailSearchResult[];
  pagination: SearchPaginationResult;
  /** Number of calendar responses filtered out (if any). */
  filteredCount?: number;
}

/** People search result. */
export interface PeopleSearchResult {
  results: PersonSearchResult[];
  returned: number;
}

/**
 * Computes Substrate search pagination hasMore from the page that was returned.
 * Uses resultsLength (not the requested size) so a short last page does not
 * falsely claim more results exist.
 */
export function computeSearchHasMore(
  from: number,
  resultsLength: number,
  size: number,
  total: number | undefined,
): boolean {
  return total !== undefined ? from + resultsLength < total : resultsLength >= size;
}

/**
 * Searches Teams messages using the Substrate v2 query API.
 */
export async function searchMessages(
  query: string,
  options: { from?: number; size?: number; maxResults?: number } = {}
): Promise<Result<SearchResult>> {
  const tokenResult = await requireSubstrateTokenAsync();
  if (!tokenResult.ok) {
    return tokenResult;
  }
  const token = tokenResult.value;

  const from = options.from ?? 0;
  const size = options.size ?? 25;

  // Generate unique IDs for this request
  const cvid = crypto.randomUUID();
  const logicalId = crypto.randomUUID();

  const body = {
    entityRequests: [{
      entityType: 'Message',
      contentSources: ['Teams'],
      propertySet: 'Optimized',
      fields: [
        'Extension_SkypeSpaces_ConversationPost_Extension_FromSkypeInternalId_String',
        'Extension_SkypeSpaces_ConversationPost_Extension_ThreadType_String',
        'Extension_SkypeSpaces_ConversationPost_Extension_SkypeGroupId_String',
      ],
      query: {
        queryString: `${query} AND NOT (isClientSoftDeleted:TRUE)`,
        displayQueryString: query,
      },
      from,
      size,
      topResultsCount: 5,
    }],
    QueryAlterationOptions: {
      EnableAlteration: true,
      EnableSuggestion: true,
      SupportedRecourseDisplayTypes: ['Suggestion'],
    },
    cvid,
    logicalId,
    scenario: {
      Dimensions: [
        { DimensionName: 'QueryType', DimensionValue: 'Messages' },
        { DimensionName: 'FormFactor', DimensionValue: 'general.web.reactSearch' },
      ],
      Name: 'powerbar',
    },
    timezone: Intl.DateTimeFormat().resolvedOptions().timeZone,
  };

  const response = await httpRequest<Record<string, unknown>>(
    SUBSTRATE_API.search,
    {
      method: 'POST',
      headers: getBearerHeaders(token),
      body: JSON.stringify(body),
    }
  );

  if (!response.ok) {
    return handleSubstrateError(response);
  }

  const data = response.value.data;
  const linkContext: LinkContext = {
    tenantId: getTenantId() ?? undefined,
    teamsBaseUrl: getTeamsBaseUrl(),
  };
  const { results, total } = parseSearchResults(
    data.EntitySets as unknown[] | undefined,
    linkContext
  );

  const maxResults = options.maxResults ?? size;
  const limitedResults = results.slice(0, maxResults);

  return ok({
    results: limitedResults,
    pagination: {
      from,
      size,
      returned: limitedResults.length,
      total,
      hasMore: computeSearchHasMore(from, results.length, size, total),
    },
  });
}

/**
 * Searches emails using the Substrate v2 query API.
 * 
 * Uses the same Substrate token as Teams message search — the
 * SubstrateSearch-Internal.ReadWrite scope covers both Teams and Exchange.
 * The key difference is entityType 'Message' with contentSources ['Exchange'].
 */
export async function searchEmails(
  query: string,
  options: { from?: number; size?: number; maxResults?: number } = {}
): Promise<Result<EmailSearchResultSet>> {
  const tokenResult = await requireSubstrateTokenAsync();
  if (!tokenResult.ok) {
    return tokenResult;
  }
  const token = tokenResult.value;

  const from = options.from ?? 0;
  const size = options.size ?? 25;

  const cvid = crypto.randomUUID();
  const logicalId = crypto.randomUUID();

  const body = {
    entityRequests: [{
      entityType: 'Message',
      contentSources: ['Exchange'],
      propertySet: 'Optimized',
      fields: [
        'Subject',
        'From',
        'ToRecipients',
        'CcRecipients',
        'DisplayTo',
        'DisplayCc',
        'DateTimeReceived',
        'DateTimeSent',
        'HasAttachments',
        'Importance',
        'IsRead',
        'Preview',
        'WebLink',
        'ConversationId',
        'FromAddress',
        'SenderSmtpAddress',
      ],
      query: {
        queryString: query,
        displayQueryString: query,
      },
      from,
      size,
      topResultsCount: 5,
    }],
    QueryAlterationOptions: {
      EnableAlteration: true,
      EnableSuggestion: true,
      SupportedRecourseDisplayTypes: ['Suggestion'],
    },
    cvid,
    logicalId,
    scenario: {
      Dimensions: [
        { DimensionName: 'QueryType', DimensionValue: 'Mail' },
        { DimensionName: 'FormFactor', DimensionValue: 'general.web.reactSearch' },
      ],
      Name: 'powerbar',
    },
    timezone: Intl.DateTimeFormat().resolvedOptions().timeZone,
  };

  const response = await httpRequest<Record<string, unknown>>(
    SUBSTRATE_API.search,
    {
      method: 'POST',
      headers: getBearerHeaders(token),
      body: JSON.stringify(body),
    }
  );

  if (!response.ok) {
    return handleSubstrateError(response);
  }

  const data = response.value.data;
  const entitySets = Array.isArray(data.EntitySets) ? data.EntitySets : undefined;
  const { results, total, filteredCount } = parseEmailSearchResults(entitySets);

  const maxResults = options.maxResults ?? size;
  const limitedResults = results.slice(0, maxResults);

  return ok({
    results: limitedResults,
    filteredCount,
    pagination: {
      from,
      size,
      returned: limitedResults.length,
      total,
      hasMore: computeSearchHasMore(from, results.length, size, total),
    },
  });
}

/**
 * Searches for people by name or email.
 */
export async function searchPeople(
  query: string,
  limit: number = 10
): Promise<Result<PeopleSearchResult>> {
  const tokenResult = await requireSubstrateTokenAsync();
  if (!tokenResult.ok) {
    return tokenResult;
  }
  const token = tokenResult.value;

  const cvid = crypto.randomUUID();
  const logicalId = crypto.randomUUID();

  const body = {
    EntityRequests: [{
      Query: {
        QueryString: query,
        DisplayQueryString: query,
      },
      EntityType: 'People',
      Size: limit,
      Fields: [
        'Id',
        'MRI',
        'DisplayName',
        'EmailAddresses',
        'GivenName',
        'Surname',
        'JobTitle',
        'Department',
        'CompanyName',
      ],
    }],
    cvid,
    logicalId,
  };

  const response = await httpRequest<Record<string, unknown>>(
    SUBSTRATE_API.peopleSearch,
    {
      method: 'POST',
      headers: getBearerHeaders(token),
      body: JSON.stringify(body),
    }
  );

  if (!response.ok) {
    return handleSubstrateError(response);
  }

  const results = parsePeopleResults(response.value.data.Groups as unknown[] | undefined);

  return ok({
    results,
    returned: results.length,
  });
}

/**
 * Gets the user's frequently contacted people.
 */
export async function getFrequentContacts(
  limit: number = 50
): Promise<Result<PeopleSearchResult>> {
  const tokenResult = await requireSubstrateTokenAsync();
  if (!tokenResult.ok) {
    return tokenResult;
  }
  const token = tokenResult.value;

  const cvid = crypto.randomUUID();
  const logicalId = crypto.randomUUID();

  const body = {
    EntityRequests: [{
      Query: {
        QueryString: '',
        DisplayQueryString: '',
      },
      EntityType: 'People',
      Size: limit,
      Fields: [
        'Id',
        'MRI',
        'DisplayName',
        'EmailAddresses',
        'GivenName',
        'Surname',
        'JobTitle',
        'Department',
        'CompanyName',
      ],
    }],
    cvid,
    logicalId,
  };

  const response = await httpRequest<Record<string, unknown>>(
    SUBSTRATE_API.frequentContacts,
    {
      method: 'POST',
      headers: getBearerHeaders(token),
      body: JSON.stringify(body),
    }
  );

  if (!response.ok) {
    return handleSubstrateError(response);
  }

  const contacts = parsePeopleResults(response.value.data.Groups as unknown[] | undefined);

  return ok({
    results: contacts,
    returned: contacts.length,
  });
}

/** Channel search result. */
export interface ChannelSearchResultSet {
  results: ChannelSearchResult[];
  returned: number;
}

/**
 * Searches for Teams channels by name using:
 * 1. User's own teams/channels (CSA Teams List) - reliable, shows membership
 * 2. Organisation-wide discovery (Substrate suggestions) when Substrate is available
 * 3. Messaging conversation list + properties when CSA/Substrate are unavailable
 *
 * Results are merged and deduplicated, with membership status indicated.
 *
 * @param query - Channel name to search for
 * @param limit - Maximum number of results (default: 10, max: 50)
 */
export async function searchChannels(
  query: string,
  limit: number = 10
): Promise<Result<ChannelSearchResultSet>> {
  // CSA membership list does not need Substrate — prefer it first on New Teams.
  const myTeamsResult = await getMyTeamsAndChannels();

  const memberChannelIds = new Set<string>();
  const myChannelsMatching: ChannelSearchResult[] = [];

  if (myTeamsResult.ok) {
    const matching = filterChannelsByName(myTeamsResult.value.teams, query);
    for (const channel of matching) {
      memberChannelIds.add(channel.channelId);
      myChannelsMatching.push(channel);
    }

    for (const team of myTeamsResult.value.teams) {
      for (const channel of team.channels) {
        memberChannelIds.add(channel.channelId);
      }
    }
  }

  // Optional org-wide Substrate suggestions (often InteractionRequired on New Teams).
  // Use cached token only — do not trigger browser refresh for channel name lookup.
  const orgChannels: ChannelSearchResult[] = [];
  const substrateToken = getValidSubstrateToken();
  if (substrateToken) {
    const orgSearchResult = await searchChannelsOrgWide(query, limit, substrateToken);
    if (orgSearchResult.ok) {
      for (const channel of orgSearchResult.value) {
        channel.isMember = memberChannelIds.has(channel.channelId);
        orgChannels.push(channel);
      }
    }
  }

  const seenIds = new Set<string>();
  const merged: ChannelSearchResult[] = [];

  for (const channel of myChannelsMatching) {
    if (!seenIds.has(channel.channelId)) {
      seenIds.add(channel.channelId);
      merged.push(channel);
    }
  }

  for (const channel of orgChannels) {
    if (!seenIds.has(channel.channelId)) {
      seenIds.add(channel.channelId);
      merged.push(channel);
    }
  }

  // Messaging fallback when CSA/Substrate yielded nothing (encrypted CSA missing, etc.).
  if (merged.length === 0) {
    const messagingMatches = await searchChannelsViaMessaging(query, limit);
    if (messagingMatches.ok) {
      for (const channel of messagingMatches.value) {
        if (!seenIds.has(channel.channelId)) {
          seenIds.add(channel.channelId);
          merged.push(channel);
        }
      }
    } else if (!myTeamsResult.ok) {
      return myTeamsResult;
    }
  }

  const limited = merged.slice(0, limit);

  return ok({
    results: limited,
    returned: limited.length,
  });
}

/**
 * Resolves channel names from recent chatsvc conversations when CSA/Substrate
 * channel discovery is unavailable.
 */
async function searchChannelsViaMessaging(
  query: string,
  limit: number
): Promise<Result<ChannelSearchResult[]>> {
  const list = await getConversations({ limit: 100 });
  if (!list.ok) {
    return list;
  }

  const lowerQuery = query.toLowerCase();
  const matches: ChannelSearchResult[] = [];

  for (const chat of list.value.conversations) {
    if (chat.chatType !== 'channel') continue;

    let name = chat.topic;
    if (!name) {
      const props = await getConversationProperties(chat.conversationId);
      if (props.ok) {
        name = props.value.displayName || '';
      }
    }
    if (!name || !name.toLowerCase().includes(lowerQuery)) continue;

    matches.push({
      channelId: chat.conversationId,
      channelName: name,
      teamName: '',
      teamId: '',
      channelType: 'Standard',
      isMember: true,
    });

    if (matches.length >= limit) break;
  }

  return ok(matches);
}

/**
 * Internal: Searches for channels org-wide using the Substrate suggestions API.
 * 
 * This is a typeahead/autocomplete API, so matching behaviour may be inconsistent
 * for multi-word queries. Used as a supplement to the user's own teams list.
 */
async function searchChannelsOrgWide(
  query: string,
  limit: number,
  token: string
): Promise<Result<ChannelSearchResult[]>> {
  const cvid = crypto.randomUUID();
  const logicalId = crypto.randomUUID();

  const body = {
    EntityRequests: [{
      Query: {
        QueryString: query,
        DisplayQueryString: query,
      },
      EntityType: 'TeamsChannel',
      Size: Math.min(limit, 50),
    }],
    cvid,
    logicalId,
  };

  const response = await httpRequest<Record<string, unknown>>(
    SUBSTRATE_API.channelSearch,
    {
      method: 'POST',
      headers: getBearerHeaders(token),
      body: JSON.stringify(body),
    }
  );

  if (!response.ok) {
    return handleSubstrateError(response);
  }

  const data = response.value.data as Record<string, unknown> | undefined;
  const results = parseChannelResults(data?.Groups as unknown[] | undefined);

  // Mark all org-wide results as isMember: false initially
  // (caller will update based on actual membership)
  for (const result of results) {
    result.isMember = false;
  }

  return ok(results);
}
