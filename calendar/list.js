/**
 * List events functionality
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { resolveCalendarPath } = require('./calendar-utils');

/**
 * List events handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListEvents(args) {
  const count = Math.min(args.count || 10, config.MAX_RESULT_COUNT);
  const calendar = args.calendar;

  // Parse date range - default to now through 30 days from now
  const now = new Date();
  const defaultEnd = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);

  let startDateTime, endDateTime;

  if (args.startDateTime) {
    startDateTime = new Date(args.startDateTime).toISOString();
  } else {
    startDateTime = now.toISOString();
  }

  if (args.endDateTime) {
    endDateTime = new Date(args.endDateTime).toISOString();
  } else {
    endDateTime = defaultEnd.toISOString();
  }

  try {
    // Get access token
    const accessToken = await ensureAuthenticated();

    // Resolve calendar path and build calendarView endpoint
    const calendarPath = await resolveCalendarPath(accessToken, calendar);
    const endpoint = `${calendarPath}/calendarView`;

    // Add query parameters
    const queryParams = {
      startDateTime,
      endDateTime,
      $top: count,
      $orderby: 'start/dateTime',
      $select: config.CALENDAR_SELECT_FIELDS
    };

    // Make API call
    const response = await callGraphAPI(accessToken, 'GET', endpoint, null, queryParams);

    if (!response.value || response.value.length === 0) {
      const startDisplay = new Date(startDateTime).toLocaleDateString();
      const endDisplay = new Date(endDateTime).toLocaleDateString();
      return {
        content: [{
          type: "text",
          text: `No calendar events found between ${startDisplay} and ${endDisplay}.`
        }]
      };
    }

    // Format results
    const eventList = response.value.map((event, index) => {
      const startDate = new Date(event.start.dateTime).toLocaleString();
      const endDate = new Date(event.end.dateTime).toLocaleString();
      const location = event.location?.displayName || 'No location';

      return `${index + 1}. ${event.subject}\n   Time: ${startDate} - ${endDate}\n   Location: ${location}\n   ID: ${event.id}`;
    }).join("\n\n");

    const startDisplay = new Date(startDateTime).toLocaleDateString();
    const endDisplay = new Date(endDateTime).toLocaleDateString();

    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} event(s) between ${startDisplay} and ${endDisplay}:\n\n${eventList}`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: "text",
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error listing events: ${error.message}`
      }]
    };
  }
}

module.exports = handleListEvents;
