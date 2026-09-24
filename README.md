# CalendarBot Apps Script

This is a Google-only, standalone port of CalendarBot's processing engine. It includes:

- Calendar `EventUpdated` triggers and incremental Calendar API sync tokens
- Travel-to and travel-from buffers for accepted events with a physical address
- Flight travel, boarding, layover, and arrival buffers
- Rule-based and external-attendee event coloring
- A 15-minute reconciliation trigger for missed/coalesced Calendar changes

It intentionally excludes cross-calendar sync, iCloud/CalDAV, Supabase caches, Slack status, and the separate managed-event scheduler.

## How it works

1. An installable Calendar trigger or the 15-minute safety-net trigger starts a reconciliation run.
2. Calendar API incremental sync returns events changed since the previous successful run. A full bounded sync is used when no token exists or Calendar invalidates it.
3. CalendarBot gathers a small time window around changed events, then recalculates travel, flight, and color decisions using the current surrounding schedule.
4. Derived events are identified by private `calendarbotKey` properties and are created, updated, or removed to match those decisions.

The trigger payload identifies only the calendar, not the changed event. The `calendar_change_received` logs identify every event returned by the corresponding sync batch.

## Color rules

Set `colorRules` in `Config.gs` (or provide the same object to `saveConfiguration()`). Rules are evaluated in order and the first match wins.

```js
colorRules: {
  enabled: true,
  rules: [
    {
      matchField: 'title',
      matchType: 'starts_with',
      pattern: 'Client:',
      colorId: 'tomato',
      caseSensitive: false,
    },
    {
      matchField: 'description',
      matchType: 'contains',
      pattern: '#focus',
      colorId: 'blueberry',
      caseSensitive: false,
    },
    {
      matchField: 'title',
      matchType: 'exact',
      pattern: 'Lunch',
      colorId: 'banana',
      caseSensitive: false,
    },
  ],
  externalAttendee: {
    enabled: true,
    internalDomains: ['yourcompany.com', 'subsidiary.com'],
    colorId: 'tangerine',
    priority: 'after_rules', // use 'before_rules' to override matching rules
  },
},
```

`matchField` accepts `title` or `description`; `matchType` accepts `contains`, `exact`, or `starts_with`. Set any `colorId` field to one of: `lavender`, `sage`, `grape`, `flamingo`, `banana`, `tangerine`, `peacock`, `graphite`, `blueberry`, `basil`, or `tomato`. Numeric IDs `1` through `11` remain supported.

A minimal deep-work rule is:

```js
colorRules: {
  enabled: true,
  rules: [{
    matchField: 'description',
    matchType: 'contains',
    pattern: '#deepwork',
    colorId: 'blueberry',
    caseSensitive: false,
  }],
  externalAttendee: { enabled: false },
},
```

## Install

1. Create a standalone Apps Script project and add `Code.gs`, `Config.gs`, and `appsscript.json`.
2. In **Services**, enable **Google Calendar API** (the Advanced Calendar service). If prompted, enable its associated API in the Google Cloud project as well.
3. Edit `DEFAULT_CONFIG` in `Config.gs`, especially `calendars`. Use the calendar owner's email for both fields; Calendar update events identify the calendar by that email, so do not use the Calendar API alias `primary` here. Run `saveDefaultConfiguration()` once from the editor and approve scopes. Keep `Config.gs` when replacing `Code.gs` with a later engine update.
4. Run `install()` once. It creates a Calendar-change trigger for each configured calendar and a 15-minute reconciliation trigger.

Use `resetSync('you@example.com')` (with the configured calendar email) to force the next run to reconcile a calendar in full. For a complete repair that also removes tagged bot events whose deleted source event is absent from the full listing, run `fullReconcileCalendar('you@example.com')`. Use `uninstall()` to remove triggers without deleting any events.

After modifying `DEFAULT_CONFIG`, run `saveDefaultConfiguration()` once. Code-only changes do not require it.

Calendar processing is bounded to events starting from one day ago through the next 12 weeks. This keeps incremental-sync responses from causing historical calendar events to be reprocessed.

## State and safety

Bot-created events carry private Calendar extended properties (`calendarbotKey`) and a description marker. This makes create/update/delete idempotent without a database and prevents trigger loops. Sync tokens and configuration are stored in the installing user's User Properties.

This is appropriate for a personal or modest per-user deployment. Apps Script has a six-minute execution limit and finite trigger/runtime quotas, so retain the service backend for large, multi-user workloads.

## Logging

Every trigger receipt, synchronization result, reconciliation, Calendar mutation, color application, skipped lock, and failure emits one structured JSON line to the Apps Script execution log. Each trigger's synchronized change batch is logged as `calendar_change_received`, including event ID, title, start date/time, status, recurrence parent, and whether CalendarBot created it. Every CalendarBot mutation is also logged as `event_change` with an `added`, `modified`, or `removed` action, the affected event ID, and its reason. Unchanged derived events are logged as `bot_event_unchanged` without a Calendar write. Bot-event and color actions include the source event ID, title, and start date/time, along with the feature name. Logs omit event descriptions, locations, and attendee email addresses.
