/**
 * CalendarBot for Google Apps Script
 *
 * Google-only port of the process-calendar capabilities that can run inside
 * Apps Script: incremental sync, travel buffers, flight buffers, and automatic
 * event coloring. Cross-calendar sync is intentionally not included.
 *
 * Setup:
 *   1. Enable the Advanced Calendar service (Calendar API v3).
 *   2. Edit Config.gs, then run saveDefaultConfiguration().
 *   3. Run install() once while signed in as the calendar owner.
 *
 * State is stored in User Properties. Install one copy per user: Apps Script
 * installable triggers execute as the account which created the trigger.
 */

const CONFIG_PROPERTY = 'calendarbot.config.v1';
const SYNC_PREFIX = 'calendarbot.sync.';
const BOT_EVENT_CACHE_PREFIX = 'calendarbot.bot-event.';
const BOT_MARKER = 'Scheduled by CalendarBot';
const LOOKAHEAD_WEEKS = 12;
const CALENDAR_COLOR_IDS = {
  lavender: '1', sage: '2', grape: '3', flamingo: '4', banana: '5',
  tangerine: '6', peacock: '7', graphite: '8', blueberry: '9', basil: '10', tomato: '11',
};

/**
 * Persists the configuration for the installing user. Trigger executions read
 * this stored value, not DEFAULT_CONFIG, so run this after config edits.
 */
function saveConfiguration(config) {
  validateConfig_(config);
  PropertiesService.getUserProperties().setProperty(CONFIG_PROPERTY, JSON.stringify(config));
  log_('configuration_saved', { calendars: config.calendars.length });
}

/** Persists the DEFAULT_CONFIG object defined in Config.gs for the current user. */
function saveDefaultConfiguration() {
  saveConfiguration(DEFAULT_CONFIG);
}

/**
 * Installs one Calendar-update trigger per configured calendar plus a 15-minute
 * reconciliation trigger. Reconciliation covers missed/coalesced notifications.
 */
function install() {
  const config = getConfig_();
  uninstall();
  config.calendars.forEach(({ calendarId, triggerEmail }) => {
    ScriptApp.newTrigger('onCalendarEventUpdated')
      .forUserCalendar(triggerEmail)
      .onEventUpdated()
      .create();
    log_('calendar_trigger_installed', { calendarId: calendarId, triggerEmail: triggerEmail });
  });
  ScriptApp.newTrigger('reconcileAllCalendars').timeBased().everyMinutes(15).create();
  log_('reconciliation_trigger_installed', { intervalMinutes: 15 });
}

/** Removes only triggers owned by this script. It does not delete Calendar events. */
function uninstall() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (['onCalendarEventUpdated', 'reconcileAllCalendars'].includes(trigger.getHandlerFunction())) {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });
  if (removed) log_('triggers_removed', { count: removed });
}

/**
 * Installable Calendar trigger entry point. Google supplies only the calendar
 * ID, so syncChanges_ determines the individual events that changed.
 */
function onCalendarEventUpdated(e) {
  if (!e || !e.calendarId) return;
  log_('calendar_trigger_received', { calendarId: e.calendarId, triggerUid: e.triggerUid || null });
  processCalendar_(e.calendarId);
}

/**
 * Periodic safety net for missed or coalesced Calendar notifications. This is
 * also safe to run manually when a normal trigger run needs to be retried.
 */
function reconcileAllCalendars() {
  log_('reconciliation_started', { calendars: getConfig_().calendars.length });
  getConfig_().calendars.forEach(({ calendarId }) => processCalendar_(calendarId));
  log_('reconciliation_completed', {});
}

/** Clears one calendar's token so its next run performs a full sync. */
function resetSync(calendarId) {
  PropertiesService.getUserProperties().deleteProperty(syncPropertyKey_(calendarId));
  log_('sync_reset', { calendarId: calendarId });
}

/**
 * Run a complete repair pass for one configured calendar. This resets the
 * token, processes a full Calendar API listing, then removes CalendarBot
 * artifacts whose source event is absent from the reconciliation window.
 */
function fullReconcileCalendar(calendarId) {
  const config = getConfig_();
  if (!config.calendars.some((entry) => entry.calendarId === calendarId)) {
    throw new Error(`Calendar is not configured: ${calendarId}`);
  }
  const startedAt = Date.now();
  log_('full_reconciliation_started', { calendarId: calendarId });
  resetSync(calendarId);
  processCalendar_(calendarId);
  const removed = removeOrphanedBotEvents_(calendarId);
  log_('full_reconciliation_completed', { calendarId: calendarId, orphanedBotEventsRemoved: removed, elapsedMs: Date.now() - startedAt });
}

/**
 * Reconciles all CalendarBot behavior for one calendar change batch. A user
 * lock serializes trigger and clock executions because they share sync tokens
 * and can otherwise make duplicate create decisions.
 */
function processCalendar_(calendarId) {
  const lock = LockService.getUserLock();
  if (!lock.tryLock(5000)) {
    log_('processing_skipped_locked', { calendarId: calendarId });
    return; // A change-trigger is already processing this user.
  }
  const startedAt = Date.now();
  try {
    const config = getConfig_();
    if (!config.calendars.some((entry) => entry.calendarId === calendarId)) {
      log_('processing_skipped_unconfigured', { calendarId: calendarId });
      return;
    }
    log_('processing_started', { calendarId: calendarId });
    const changed = syncChanges_(calendarId);
    changed.forEach((event) => log_('calendar_change_received', {
      calendarId: calendarId,
      status: event.status || null,
      recurringEventId: event.recurringEventId || null,
      isBotEvent: isBotEvent_(event),
      ...eventLogDetails_(event),
    }));
    if (!changed.length) {
      log_('processing_no_changes', { calendarId: calendarId });
      return;
    }

    const botCache = {};
    const active = changed.filter((event) => event.status !== 'cancelled');

    // First remove artifacts owned by cancelled or no-longer-qualifying events.
    changed.filter((event) => event.status === 'cancelled').forEach((event) => {
      removeFeatureSet_(calendarId, event, ['travel_time', 'travel_time_return', 'flight_travel', 'flight_boarding', 'flight_layover', 'flight_travel_from'], botCache);
    });

    if (config.colorRules.enabled) {
      active.filter((event) => !isBotEvent_(event)).forEach((event) => applyColorRules_(calendarId, event, config.colorRules));
    }

    // Include cancellations as context seeds: cancelling an event can change a
    // neighboring, otherwise unchanged buffer decision.
    const changedTimed = changed.filter((event) => event.start && event.start.dateTime && !isBotEvent_(event));
    if (config.travelTime.enabled && changedTimed.length) {
      processTravel_(calendarId, changedTimed, config.travelTime, botCache);
    }
    if (config.flightHandling.enabled) {
      processFlights_(calendarId, changed, config.flightHandling, config.colorRules, botCache);
    }
    log_('processing_completed', { calendarId: calendarId, changedEvents: changed.length, elapsedMs: Date.now() - startedAt });
  } catch (error) {
    log_('processing_failed', { calendarId: calendarId, elapsedMs: Date.now() - startedAt, error: String(error) });
    throw error;
  } finally {
    lock.releaseLock();
  }
}

/** Returns changed events and advances the incremental sync token only after paging succeeds. */
function syncChanges_(calendarId) {
  const props = PropertiesService.getUserProperties();
  const propertyKey = syncPropertyKey_(calendarId);
  const syncToken = props.getProperty(propertyKey);
  const syncMode = syncToken ? 'incremental' : 'full';
  const events = [];
  let pageToken;
  let response;
  try {
    do {
      const options = { showDeleted: true, maxResults: 2500 };
      if (pageToken) options.pageToken = pageToken;
      if (syncToken) {
        options.syncToken = syncToken;
      } else {
        options.timeMin = new Date().toISOString();
        options.timeMax = new Date(Date.now() + LOOKAHEAD_WEEKS * 7 * 86400000).toISOString();
      }
      response = Calendar.Events.list(calendarId, options);
      (response.items || []).forEach((event) => events.push(event));
      pageToken = response.nextPageToken;
    } while (pageToken);
  } catch (error) {
    // Calendar API invalidates sync tokens with HTTP 410. Retry once as full sync.
    if (syncToken && /410|sync token.*(invalid|valid)/i.test(String(error))) {
      props.deleteProperty(propertyKey);
      log_('sync_token_invalidated', { calendarId: calendarId });
      return syncChanges_(calendarId);
    }
    throw error;
  }
  if (response && response.nextSyncToken) props.setProperty(propertyKey, response.nextSyncToken);
  // Incremental sync may return old changes. Keep only the same bounded window
  // used by full reconciliation: one day back through the future lookahead.
  const retained = events.filter(isWithinReconciliationWindow_);
  log_('sync_completed', { calendarId: calendarId, mode: syncMode, receivedEvents: events.length, processedEvents: retained.length });
  return retained;
}

/**
 * Recomputes travel buffers around changed timed events. The surrounding window
 * is deliberately wider than a buffer so an unchanged neighboring event can be
 * updated when the gap between events changes.
 */
function processTravel_(calendarId, changedTimed, travel, botCache) {
  // A two-hour context makes incremental changes safe: a new Zoom meeting can
  // shrink the buffer belonging to an unchanged physical event beside it.
  const context = listWindow_(calendarId, changedTimed, 2).filter((event) => !isBotEvent_(event));
  const qualifying = context.filter(isTravelEligible_).sort(compareStart_);
  const allTimed = context.filter((event) => event.start && event.start.dateTime && isAcceptedOrOwned_(event)).sort(compareStart_);
  const decisions = computeTravelDecisions_(qualifying, travel.bufferMinutes);
  shrinkTravelAgainstConflicts_(decisions, qualifying, allTimed);
  mergeTravelDecisions_(decisions, qualifying);
  log_('travel_reconciliation_started', { calendarId: calendarId, changedSeeds: changedTimed.length, qualifyingEvents: qualifying.length });

  // Reconcile every nearby qualifying event, not just the changed one, because
  // either neighbor may now require an updated/deleted buffer.
  qualifying.forEach((event) => {
    const decision = decisions[event.id];
    if (decision.travelStart) {
      upsertBotEvent_(calendarId, event, 'travel_time', botEventBody_(travel.eventName, `Travel to: ${event.summary || 'event'}`, decision.travelStart, event.start.dateTime, event.start.timeZone, travel.colorId), botCache);
    } else {
      removeBotEvent_(calendarId, event, 'travel_time', botCache);
    }
    if (decision.returnEnd) {
      upsertBotEvent_(calendarId, event, 'travel_time_return', botEventBody_(travel.eventName, `Travel from: ${event.summary || 'event'}`, event.end.dateTime, decision.returnEnd, event.end.timeZone || event.start.timeZone, travel.colorId), botCache);
    } else {
      removeBotEvent_(calendarId, event, 'travel_time_return', botCache);
    }
  });

  // A changed event with an address removed needs cleanup but is absent above.
  changedTimed.filter((event) => !isTravelEligible_(event)).forEach((event) => {
    removeFeatureSet_(calendarId, event, ['travel_time', 'travel_time_return'], botCache);
  });
}

/**
 * Reconciles flight-related buffers. Adjacent flights in the local window are
 * used to determine connections and layovers rather than trusting event order
 * from Calendar's incremental-sync response.
 */
function processFlights_(calendarId, changed, flight, colorRules, botCache) {
  // Cancelled flights seed a neighboring-flight reconciliation as well.
  const changedFlights = changed.filter(isFlightEvent_);
  if (!changedFlights.length) return;
  // ±8h mirrors the existing service's layover lookup window.
  const flights = listWindow_(calendarId, changedFlights, 8).filter((event) => !isBotEvent_(event) && isFlightEvent_(event)).sort(compareStart_);
  log_('flight_reconciliation_started', { calendarId: calendarId, changedFlights: changedFlights.length, contextualFlights: flights.length });
  flights.forEach((event, index) => {
    const prior = flights[index - 1];
    const next = flights[index + 1];
    const priorIsConnection = prior && isLayover_(prior, event, flight.layoverBufferMinutes);
    const nextIsConnection = next && isLayover_(event, next, flight.layoverBufferMinutes);
    const boardingStart = addMinutes_(event.start.dateTime, -flight.boardingMinutes);
    const allTimed = listWindow_(calendarId, [event], 4).filter((other) => !isBotEvent_(other) && other.start && other.start.dateTime);

    if (flight.travelToEnabled && !priorIsConnection) {
      const travelStart = shrinkBefore_(addMinutes_(boardingStart, -flight.travelMinutes), boardingStart, event.id, allTimed);
      if (travelStart) upsertBotEvent_(calendarId, event, 'flight_travel', botEventBody_(flight.travelToName, `Travel to airport for: ${event.summary || 'flight'}`, travelStart, boardingStart, event.start.timeZone, flight.travelToColorId), botCache);
      else removeBotEvent_(calendarId, event, 'flight_travel', botCache);
    } else {
      removeBotEvent_(calendarId, event, 'flight_travel', botCache);
    }

    if (flight.boardingEnabled) upsertBotEvent_(calendarId, event, 'flight_boarding', botEventBody_(flight.boardingName, `Boarding time for: ${event.summary || 'flight'}`, boardingStart, event.start.dateTime, event.start.timeZone, flight.boardingColorId), botCache);
    else removeBotEvent_(calendarId, event, 'flight_boarding', botCache);

    if (flight.layoverEnabled && nextIsConnection) {
      const nextBoarding = addMinutes_(next.start.dateTime, -flight.boardingMinutes);
      upsertBotEvent_(calendarId, event, 'flight_layover', botEventBody_(flight.layoverName, `Layover between: ${event.summary || 'flight'} and ${next.summary || 'flight'}`, event.end.dateTime, nextBoarding, event.start.timeZone, flight.layoverColorId), botCache);
    } else {
      removeBotEvent_(calendarId, event, 'flight_layover', botCache);
    }

    if (flight.travelFromEnabled && !nextIsConnection && event.end && event.end.dateTime) {
      const travelEnd = shrinkAfter_(event.end.dateTime, addMinutes_(event.end.dateTime, flight.travelFromMinutes), event.id, allTimed);
      if (travelEnd) upsertBotEvent_(calendarId, event, 'flight_travel_from', botEventBody_(flight.travelFromName, `Travel from airport for: ${event.summary || 'flight'}`, event.end.dateTime, travelEnd, event.end.timeZone || event.start.timeZone, flight.travelFromColorId), botCache);
      else removeBotEvent_(calendarId, event, 'flight_travel_from', botCache);
    } else {
      removeBotEvent_(calendarId, event, 'flight_travel_from', botCache);
    }

    if (colorRules.enabled) applyColorRules_(calendarId, event, colorRules);
  });
}

/**
 * Calculates the desired before/after buffers before non-travel conflicts are
 * considered. Buffers between adjacent eligible events are merged downstream.
 */
function computeTravelDecisions_(events, bufferMinutes) {
  const decisions = {};
  events.forEach((event, index) => {
    let travelStart = addMinutes_(event.start.dateTime, -bufferMinutes);
    let returnEnd = event.end && event.end.dateTime ? addMinutes_(event.end.dateTime, bufferMinutes) : null;
    const previous = events[index - 1];
    const next = events[index + 1];
    if (previous && previous.end && previous.end.dateTime) {
      const gap = minutesBetween_(previous.end.dateTime, event.start.dateTime);
      if (gap <= 0) travelStart = null;
      else if (gap < 2 * bufferMinutes) travelStart = previous.end.dateTime;
    }
    if (next && event.end && event.end.dateTime) {
      const gap = minutesBetween_(event.end.dateTime, next.start.dateTime);
      if (gap < 2 * bufferMinutes) returnEnd = null;
    }
    decisions[event.id] = { travelStart: travelStart, returnEnd: returnEnd };
  });
  return decisions;
}

function shrinkTravelAgainstConflicts_(decisions, qualifying, allTimed) {
  const qualifyingIds = new Set(qualifying.map((event) => event.id));
  qualifying.forEach((event) => {
    const decision = decisions[event.id];
    const start = new Date(event.start.dateTime).getTime();
    const end = event.end && event.end.dateTime ? new Date(event.end.dateTime).getTime() : null;
    if (decision.travelStart) {
      let latest = new Date(decision.travelStart).getTime();
      allTimed.forEach((other) => {
        if (other.id === event.id || qualifyingIds.has(other.id) || !other.end || !other.end.dateTime) return;
        if (new Date(other.start.dateTime).getTime() < start && new Date(other.end.dateTime).getTime() > latest) latest = Math.max(latest, new Date(other.end.dateTime).getTime());
      });
      decision.travelStart = latest >= start ? null : new Date(latest).toISOString();
    }
    if (decision.returnEnd && end) {
      let earliest = new Date(decision.returnEnd).getTime();
      allTimed.forEach((other) => {
        if (other.id === event.id || qualifyingIds.has(other.id)) return;
        const otherEnd = other.end && other.end.dateTime ? new Date(other.end.dateTime).getTime() : new Date(other.start.dateTime).getTime();
        if (otherEnd > end && new Date(other.start.dateTime).getTime() < earliest) earliest = Math.min(earliest, new Date(other.start.dateTime).getTime());
      });
      decision.returnEnd = earliest <= end ? null : new Date(earliest).toISOString();
    }
  });
}

function mergeTravelDecisions_(decisions, events) {
  for (let i = 0; i < events.length - 1; i++) {
    const current = decisions[events[i].id];
    const next = decisions[events[i + 1].id];
    if (current.returnEnd && next.travelStart && new Date(current.returnEnd) > new Date(next.travelStart)) {
      next.travelStart = events[i].end.dateTime;
      current.returnEnd = null;
    }
  }
}

/** Applies the first matching configured color rule, respecting its priority. */
function applyColorRules_(calendarId, event, colorConfig) {
  const external = colorConfig.externalAttendee || {};
  const useExternalFirst = external.enabled && external.priority === 'before_rules';
  if (useExternalFirst && applyExternalColor_(calendarId, event, external)) return;
  for (const rule of colorConfig.rules || []) {
    const text = rule.matchField === 'description' ? (event.description || '') : (event.summary || '');
    const normalize = (value) => rule.caseSensitive ? value.replace(/\uFE0F/g, '') : value.replace(/\uFE0F/g, '').toLowerCase();
    const candidate = normalize(text);
    const pattern = normalize(rule.pattern || '');
    const matched = rule.matchType === 'exact' ? candidate === pattern : rule.matchType === 'starts_with' ? candidate.startsWith(pattern) : candidate.includes(pattern);
    if (matched) {
      const colorId = calendarColorId_(rule.colorId);
      if (colorId && event.colorId !== colorId) {
        Calendar.Events.patch({ colorId: colorId }, calendarId, event.id);
        logEventChange_('modified', calendarId, event.id, { reason: 'color_rule', colorId: colorId, ...eventLogDetails_(event) });
        log_('color_rule_applied', { calendarId: calendarId, colorId: colorId, matchField: rule.matchField, matchType: rule.matchType, ...eventLogDetails_(event) });
      }
      return;
    }
  }
  if (!useExternalFirst) applyExternalColor_(calendarId, event, external);
}

function applyExternalColor_(calendarId, event, config) {
  const internal = new Set((config.internalDomains || []).map((domain) => domain.toLowerCase()));
  const externalAttendee = (event.attendees || []).some((attendee) => {
    const domain = attendee.email && attendee.email.split('@')[1];
    return !attendee.self && domain && !internal.has(domain.toLowerCase());
  });
  const colorId = calendarColorId_(config.colorId);
  if (externalAttendee && colorId && event.colorId !== colorId) {
    Calendar.Events.patch({ colorId: colorId }, calendarId, event.id);
    logEventChange_('modified', calendarId, event.id, { reason: 'external_attendee_color', colorId: colorId, ...eventLogDetails_(event) });
    log_('external_attendee_color_applied', { calendarId: calendarId, colorId: colorId, ...eventLogDetails_(event) });
    return true;
  }
  return false;
}

/**
 * Creates or updates the one derived event identified by source ID and feature.
 * Private properties are the durable identity; the user cache bridges Calendar
 * API indexing delays immediately after a create.
 */
function upsertBotEvent_(calendarId, source, feature, body, cache) {
  const key = botKey_(source.id, feature);
  body.extendedProperties = { private: {
    calendarbotKey: key,
    calendarbotSourceId: source.id,
    calendarbotSourceTitle: source.summary || '(untitled)',
    calendarbotSourceStart: source.start && (source.start.dateTime || source.start.date) || '',
    calendarbotFeature: feature,
  } };
  const existing = findBotEvent_(calendarId, key, cache, source);
  if (existing) {
    if (botEventMatches_(existing, body)) {
      rememberBotEvent_(calendarId, key, existing.id);
      log_('bot_event_unchanged', { calendarId: calendarId, botEventId: existing.id, feature: feature, ...eventLogDetails_(source) });
      return;
    }
    try {
      Calendar.Events.patch(body, calendarId, existing.id);
      rememberBotEvent_(calendarId, key, existing.id);
      logEventChange_('modified', calendarId, existing.id, { reason: feature, sourceEventId: source.id, ...eventLogDetails_(source) });
      log_('bot_event_updated', { calendarId: calendarId, botEventId: existing.id, feature: feature, ...eventLogDetails_(source) });
      return;
    } catch (error) {
      if (!/404|410/.test(String(error))) throw error;
    }
  }
  const created = Calendar.Events.insert(body, calendarId, { sendUpdates: 'none' });
  cache[calendarId + ':' + key] = created;
  // Extended-property filtering can lag behind event writes. Retain the ID long
  // enough for the write-triggered reconciliation pass to retrieve it directly.
  rememberBotEvent_(calendarId, key, created.id);
  logEventChange_('added', calendarId, created.id, { reason: feature, sourceEventId: source.id, ...eventLogDetails_(source) });
  log_('bot_event_created', { calendarId: calendarId, botEventId: created.id, feature: feature, ...eventLogDetails_(source) });
}

function removeFeatureSet_(calendarId, source, features, cache) {
  features.forEach((feature) => removeBotEvent_(calendarId, source, feature, cache));
}

/** Removes the derived event for a source/feature pair when it is no longer needed. */
function removeBotEvent_(calendarId, source, feature, cache) {
  const sourceId = typeof source === 'string' ? source : source.id;
  const key = botKey_(sourceId, feature);
  const existing = findBotEvent_(calendarId, key, cache, typeof source === 'string' ? null : source);
  if (!existing) return;
  try { Calendar.Events.remove(calendarId, existing.id, { sendUpdates: 'none' }); } catch (error) {
    if (!/404|410/.test(String(error))) throw error;
  }
  delete cache[calendarId + ':' + key];
  forgetBotEvent_(calendarId, key);
  logEventChange_('removed', calendarId, existing.id, { reason: feature, sourceEventId: sourceId, ...(typeof source === 'string' ? {} : eventLogDetails_(source)) });
  log_('bot_event_removed', {
    calendarId: calendarId,
    botEventId: existing.id,
    feature: feature,
    ...(typeof source === 'string' ? { sourceEventId: sourceId } : eventLogDetails_(source)),
  });
}

/**
 * Locates a derived event without creating it. Lookup order favors the current
 * run cache, then the short-lived ID cache, then Calendar's property filter.
 * Invited-event copies that evade that filter fall back to a bounded local scan.
 */
function findBotEvent_(calendarId, key, cache, source) {
  const cacheKey = calendarId + ':' + key;
  if (Object.prototype.hasOwnProperty.call(cache, cacheKey)) return cache[cacheKey];
  const rememberedId = CacheService.getUserCache().get(botEventCacheKey_(calendarId, key));
  if (rememberedId) {
    try {
      const remembered = Calendar.Events.get(calendarId, rememberedId);
      const properties = remembered.extendedProperties && remembered.extendedProperties.private;
      if (properties && properties.calendarbotKey === key) {
        cache[cacheKey] = remembered;
        return remembered;
      }
    } catch (error) {
      if (!/404|410/.test(String(error))) throw error;
    }
    forgetBotEvent_(calendarId, key);
  }
  const result = Calendar.Events.list(calendarId, {
    privateExtendedProperty: 'calendarbotKey=' + key,
    showDeleted: false,
    maxResults: 1,
  });
  cache[cacheKey] = (result.items || [])[0] || null;
  // Some invited-event copies do not appear in Calendar's server-side
  // privateExtendedProperty filter. Search the source event's small local window
  // and match the property in the returned resources instead.
  if (!cache[cacheKey] && source && source.start && source.start.dateTime) {
    const start = new Date(source.start.dateTime).getTime() - 86400000;
    const end = new Date((source.end && source.end.dateTime) || source.start.dateTime).getTime() + 86400000;
    cache[cacheKey] = listRange_(calendarId, new Date(start).toISOString(), new Date(end).toISOString()).find((event) => {
      const properties = event.extendedProperties && event.extendedProperties.private;
      return properties && properties.calendarbotKey === key;
    }) || null;
  }
  if (cache[cacheKey]) rememberBotEvent_(calendarId, key, cache[cacheKey].id);
  return cache[cacheKey];
}

function botEventCacheKey_(calendarId, key) { return BOT_EVENT_CACHE_PREFIX + Utilities.base64EncodeWebSafe(calendarId + ':' + key); }
function rememberBotEvent_(calendarId, key, eventId) { CacheService.getUserCache().put(botEventCacheKey_(calendarId, key), eventId, 600); }
function forgetBotEvent_(calendarId, key) { CacheService.getUserCache().remove(botEventCacheKey_(calendarId, key)); }

/**
 * Compares only the fields CalendarBot owns. Calendar may serialize identical
 * instants with different ISO offsets, so start/end comparison uses timestamps.
 */
function botEventMatches_(existing, desired) {
  const existingPrivate = existing.extendedProperties && existing.extendedProperties.private || {};
  const desiredPrivate = desired.extendedProperties && desired.extendedProperties.private || {};
  return existing.summary === desired.summary &&
    existing.description === desired.description &&
    sameEventTime_(existing.start, desired.start) &&
    sameEventTime_(existing.end, desired.end) &&
    (!desired.colorId || existing.colorId === desired.colorId) &&
    Object.keys(desiredPrivate).every((key) => existingPrivate[key] === desiredPrivate[key]);
}

function sameEventTime_(left, right) {
  return left && right && left.dateTime && right.dateTime &&
    new Date(left.dateTime).getTime() === new Date(right.dateTime).getTime();
}

function botEventBody_(summary, description, start, end, timeZone, colorId) {
  const text = `${description}\n\n<i>${BOT_MARKER}</i>`;
  const body = { summary: summary, description: text, start: { dateTime: start, timeZone: timeZone }, end: { dateTime: end, timeZone: timeZone } };
  const resolvedColorId = calendarColorId_(colorId);
  if (resolvedColorId) body.colorId = resolvedColorId;
  return body;
}

/**
 * Removes tagged events whose source is absent from a full reconciliation
 * window. Marker-only legacy events are retained because ownership is unknown.
 */
function removeOrphanedBotEvents_(calendarId) {
  const now = Date.now();
  // Extra time on either side accounts for travel and boarding buffers.
  const events = listRange_(
    calendarId,
    new Date(now - 86400000).toISOString(),
    new Date(now + LOOKAHEAD_WEEKS * 7 * 86400000 + 86400000).toISOString(),
  );
  const liveSourceIds = new Set(events.filter((event) => !isBotEvent_(event)).map((event) => event.id));
  let removed = 0;
  events.filter(isBotEvent_).forEach((botEvent) => {
    const properties = botEvent.extendedProperties && botEvent.extendedProperties.private;
    const sourceId = properties && properties.calendarbotSourceId;
    // Marker-only events from an older script cannot be proven orphaned safely.
    if (!sourceId || liveSourceIds.has(sourceId)) return;
    try {
      Calendar.Events.remove(calendarId, botEvent.id, { sendUpdates: 'none' });
      removed++;
      logEventChange_('removed', calendarId, botEvent.id, { reason: 'orphaned_bot_event', sourceEventId: sourceId, feature: properties.calendarbotFeature || null });
      log_('orphaned_bot_event_removed', {
        calendarId: calendarId,
        botEventId: botEvent.id,
        feature: properties.calendarbotFeature || null,
        sourceEventId: sourceId,
        sourceEventTitle: properties.calendarbotSourceTitle || '(unknown; created by an older script version)',
        sourceEventStart: properties.calendarbotSourceStart || null,
      });
    } catch (error) {
      if (!/404|410/.test(String(error))) throw error;
    }
  });
  log_('orphan_scan_completed', { calendarId: calendarId, inspectedEvents: events.length, orphanedBotEventsRemoved: removed });
  return removed;
}

/** Lists single-event instances in the smallest window surrounding seed events. */
function listWindow_(calendarId, events, hours) {
  const times = events.flatMap((event) => [new Date(event.start.dateTime).getTime(), new Date((event.end && event.end.dateTime) || event.start.dateTime).getTime()]);
  if (!times.length) return [];
  return listRange_(calendarId, new Date(Math.min.apply(null, times) - hours * 3600000).toISOString(), new Date(Math.max.apply(null, times) + hours * 3600000).toISOString());
}

/** Lists all pages in an ISO-8601 range, expanding recurring events into instances. */
function listRange_(calendarId, timeMin, timeMax) {
  const items = [];
  let pageToken;
  do {
    const result = Calendar.Events.list(calendarId, {
      timeMin: timeMin,
      timeMax: timeMax,
      singleEvents: true,
      showDeleted: false,
      maxResults: 2500,
      pageToken: pageToken,
    });
    (result.items || []).forEach((event) => items.push(event));
    pageToken = result.nextPageToken;
  } while (pageToken);
  return items;
}

function isTravelEligible_(event) {
  return event.status !== 'cancelled' && event.start && event.start.dateTime && hasPhysicalAddress_(event.location) && isAcceptedOrOwned_(event);
}
function hasPhysicalAddress_(location) { return !!location && !/(zoom\.us|meet\.google\.com|teams\.microsoft\.com|webex)/i.test(location) && /\d+\s+\w/.test(location); }
function isAcceptedOrOwned_(event) { const self = (event.attendees || []).find((attendee) => attendee.self); return !self || self.responseStatus === 'accepted'; }
function isFlightEvent_(event) { const title = event.summary || ''; const description = event.description || ''; return !isBotEvent_(event) && (description.includes('#flight') || title.startsWith('Flight:') || /^[✈️✈]/.test(title)); }
function isBotEvent_(event) { return !!((event.extendedProperties && event.extendedProperties.private && event.extendedProperties.private.calendarbotKey) || (event.description || '').includes(BOT_MARKER)); }
function isLayover_(first, second, minBuffer) { if (!first || !second || !first.end || !first.end.dateTime) return false; const gap = minutesBetween_(first.end.dateTime, second.start.dateTime); return gap >= minBuffer && gap <= 480; }
function shrinkBefore_(start, end, sourceId, events) { let latest = new Date(start).getTime(); events.forEach((event) => { if (event.id !== sourceId && event.end && event.end.dateTime && new Date(event.start.dateTime) < new Date(end) && new Date(event.end.dateTime) > latest) latest = Math.max(latest, new Date(event.end.dateTime).getTime()); }); return latest >= new Date(end).getTime() ? null : new Date(latest).toISOString(); }
function shrinkAfter_(start, end, sourceId, events) { let earliest = new Date(end).getTime(); events.forEach((event) => { if (event.id !== sourceId && new Date((event.end && event.end.dateTime) || event.start.dateTime) > new Date(start) && new Date(event.start.dateTime) < earliest) earliest = Math.min(earliest, new Date(event.start.dateTime).getTime()); }); return earliest <= new Date(start).getTime() ? null : new Date(earliest).toISOString(); }
function compareStart_(a, b) { return new Date(a.start.dateTime).getTime() - new Date(b.start.dateTime).getTime(); }
function addMinutes_(iso, minutes) { return new Date(new Date(iso).getTime() + minutes * 60000).toISOString(); }
function minutesBetween_(start, end) { return (new Date(end).getTime() - new Date(start).getTime()) / 60000; }
/** Resolves a readable palette name or legacy numeric ID to Calendar API colorId. */
function calendarColorId_(value) {
  if (value === undefined || value === null || value === '') return null;
  const color = String(value).trim().toLowerCase();
  if (/^(?:[1-9]|1[01])$/.test(color)) return color;
  if (CALENDAR_COLOR_IDS[color]) return CALENDAR_COLOR_IDS[color];
  throw new Error(`Unknown calendar color: ${value}. Use a palette name such as "peacock" or an ID from 1 through 11.`);
}
function botKey_(sourceId, feature) { return sourceId + ':' + feature; }
function eventLogDetails_(event) { return { sourceEventId: event.id, sourceEventTitle: event.summary || '(untitled)', sourceEventStart: event.start && (event.start.dateTime || event.start.date) || null }; }
function logEventChange_(change, calendarId, eventId, details) { log_('event_change', { calendarId: calendarId, change: change, eventId: eventId, ...details }); }
function syncPropertyKey_(calendarId) { return SYNC_PREFIX + Utilities.base64EncodeWebSafe(calendarId); }
function isWithinReconciliationWindow_(event) {
  // Cancelled recurring instances can carry originalStartTime instead of start.
  const boundary = event.start || event.originalStartTime;
  const raw = boundary && (boundary.dateTime || boundary.date);
  if (!raw) return false;
  const time = new Date(raw).getTime();
  const now = Date.now();
  return time >= now - 86400000 && time <= now + LOOKAHEAD_WEEKS * 7 * 86400000;
}
function getConfig_() { const raw = PropertiesService.getUserProperties().getProperty(CONFIG_PROPERTY); if (!raw) throw new Error('No CalendarBot configuration. Edit DEFAULT_CONFIG and run saveDefaultConfiguration().'); const config = JSON.parse(raw); validateConfig_(config); return config; }
/** Write JSON to the Apps Script execution log without event contents or attendees. */
function log_(action, details) { console.log(JSON.stringify({ service: 'calendarbot', action: action, at: new Date().toISOString(), ...details })); }
function validateConfig_(config) {
  if (!config || !Array.isArray(config.calendars) || !config.calendars.length || config.calendars.some((entry) => !entry.calendarId || !entry.triggerEmail)) {
    throw new Error('config.calendars must contain { calendarId, triggerEmail } entries.');
  }
  const flight = config.flightHandling || {};
  const colorRules = config.colorRules || {};
  [
    config.travelTime && config.travelTime.colorId,
    flight.travelToColorId, flight.travelFromColorId, flight.boardingColorId, flight.layoverColorId,
    colorRules.externalAttendee && colorRules.externalAttendee.colorId,
    ...(colorRules.rules || []).map((rule) => rule.colorId),
  ].forEach(calendarColorId_);
}
