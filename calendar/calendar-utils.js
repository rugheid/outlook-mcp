/**
 * Calendar utilities
 */
const { callGraphAPI } = require('../utils/graph-api');

/**
 * Cache of calendar information to reduce API calls
 * Format: { calendarName: { id, name } }
 */
const calendarCache = {};

/**
 * Resolve a calendar name or ID to a calendar endpoint path
 * @param {string} accessToken - Access token
 * @param {string} calendar - Calendar name or ID (optional, defaults to primary)
 * @returns {Promise<string>} - Resolved calendar endpoint path
 */
async function resolveCalendarPath(accessToken, calendar) {
  // Default to primary calendar if no calendar specified
  if (!calendar) {
    return 'me/calendar';
  }

  // Check if it looks like a calendar ID (long string)
  if (calendar.length > 50) {
    return `me/calendars/${calendar}`;
  }

  // Try to find calendar by name
  const calendarId = await getCalendarIdByName(accessToken, calendar);
  if (calendarId) {
    return `me/calendars/${calendarId}`;
  }

  // If not found, fall back to primary calendar
  console.error(`Couldn't find calendar "${calendar}", falling back to primary calendar`);
  return 'me/calendar';
}

/**
 * Get the ID of a calendar by its name
 * @param {string} accessToken - Access token
 * @param {string} calendarName - Name of the calendar to find
 * @returns {Promise<string|null>} - Calendar ID or null if not found
 */
async function getCalendarIdByName(accessToken, calendarName) {
  // Check cache first
  if (calendarCache[calendarName]) {
    console.error(`Using cached calendar ID for "${calendarName}"`);
    return calendarCache[calendarName].id;
  }

  try {
    console.error(`Looking for calendar with name "${calendarName}"`);
    const response = await callGraphAPI(
      accessToken,
      'GET',
      'me/calendars',
      null,
      { $select: 'id,name' }
    );

    if (response.value) {
      const lowerCalendarName = calendarName.toLowerCase();
      const matchingCalendar = response.value.find(
        cal => cal.name.toLowerCase() === lowerCalendarName
      );

      if (matchingCalendar) {
        console.error(`Found calendar "${calendarName}" with ID: ${matchingCalendar.id}`);
        // Cache the result
        calendarCache[calendarName] = {
          id: matchingCalendar.id,
          name: matchingCalendar.name
        };
        return matchingCalendar.id;
      }
    }

    console.error(`No calendar found matching "${calendarName}"`);
    return null;
  } catch (error) {
    console.error(`Error finding calendar "${calendarName}": ${error.message}`);
    return null;
  }
}

/**
 * Get all calendars
 * @param {string} accessToken - Access token
 * @returns {Promise<Array>} - Array of calendar objects
 */
async function getAllCalendars(accessToken) {
  try {
    const response = await callGraphAPI(
      accessToken,
      'GET',
      'me/calendars',
      null,
      {
        $select: 'id,name,canEdit,canShare,canViewPrivateItems,owner,isDefaultCalendar',
        $top: 50
      }
    );

    if (!response.value) {
      return [];
    }

    return response.value;
  } catch (error) {
    console.error(`Error getting all calendars: ${error.message}`);
    return [];
  }
}

module.exports = {
  resolveCalendarPath,
  getCalendarIdByName,
  getAllCalendars,
  calendarCache
};
