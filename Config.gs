/**
 * CalendarBot user configuration.
 *
 * Keep this file when updating Code.gs. After changing a value, run
 * saveDefaultConfiguration() once in Apps Script to store it for your user.
 */
const DEFAULT_CONFIG = {
  // Legacy colorId values: lavender, sage, grape, flamingo, banana, tangerine,
  // peacock, graphite, blueberry, basil, tomato (or numeric IDs "1"–"11").
  // New Calendar event-label IDs are account/calendar-specific; use eventLabelId
  // in place of colorId when configuring one.
  // Calendar trigger events return the owner's email as `calendarId`, so use
  // that same email for both fields. Do not use the Calendar API alias
  // "primary" here.
  calendars: [{ calendarId: 'you@example.com', triggerEmail: 'you@example.com' }],
  travelTime: {
    enabled: true,
    bufferMinutes: 30,
    eventName: 'Travel buffer',
    // Use either colorId (legacy palette) or eventLabelId (new Calendar label).
    colorId: 'graphite',
    // eventLabelId: 'your-calendar-label-id',
  },
  flightHandling: {
    enabled: true,
    travelMinutes: 90,
    travelFromMinutes: 45,
    boardingMinutes: 30,
    layoverBufferMinutes: 45,
    travelToEnabled: true,
    travelFromEnabled: false,
    boardingEnabled: true,
    layoverEnabled: true,
    travelToName: 'Travel to airport',
    travelFromName: 'Travel from airport',
    boardingName: 'Boarding',
    layoverName: 'Layover',
    travelToColorId: 'graphite',
    travelFromColorId: 'graphite',
    boardingColorId: 'banana',
    layoverColorId: 'grape',
    // Optional: color the original flight event. Omit or set to null to leave it unchanged.
    // To use a new Calendar label instead, replace flightColorId with flightEventLabelId.
    flightColorId: null,
  },
  colorRules: {
    enabled: false,
    // First matching rule wins. matchField is "title" or "description";
    // matchType is "contains", "exact", "starts_with", or "regex".
    rules: [],
    // Each rule and externalAttendee can use eventLabelId instead of colorId.
    externalAttendee: {
      enabled: false,
      internalDomains: [],
      colorId: 'tomato',
      priority: 'after_rules', // or "before_rules"
    },
  },
};
