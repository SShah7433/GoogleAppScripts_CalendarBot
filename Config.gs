/**
 * CalendarBot user configuration.
 *
 * Keep this file when updating Code.gs. After changing a value, run
 * saveDefaultConfiguration() once in Apps Script to store it for your user.
 */
const DEFAULT_CONFIG = {
  // Calendar trigger events return the owner's email as `calendarId`, so use
  // that same email for both fields. Do not use the Calendar API alias
  // "primary" here.
  calendars: [{ calendarId: 'you@example.com', triggerEmail: 'you@example.com' }],
  travelTime: {
    enabled: true,
    bufferMinutes: 30,
    eventName: 'Travel buffer',
    colorId: 'graphite',
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
  },
  colorRules: {
    enabled: false,
    // First matching rule wins. matchField is "title" or "description";
    // matchType is "contains", "exact", or "starts_with".
    rules: [],
    externalAttendee: {
      enabled: false,
      internalDomains: [],
      colorId: 'tomato',
      priority: 'after_rules', // or "before_rules"
    },
  },
};
