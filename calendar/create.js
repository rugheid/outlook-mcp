/**
 * Create event functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { DEFAULT_TIMEZONE } = require('../config');
const { resolveCalendarPath } = require('./calendar-utils');

/**
 * Validate recurrence pattern
 * @param {object} recurrence - Recurrence object to validate
 * @returns {object|null} - Error object if invalid, null if valid
 */
function validateRecurrence(recurrence) {
  if (!recurrence) return null;

  const { pattern, range } = recurrence;

  // Validate pattern
  if (!pattern || !pattern.type || !pattern.interval) {
    return { error: "Recurrence pattern must include 'type' and 'interval'" };
  }

  if (pattern.interval < 1) {
    return { error: "Recurrence interval must be at least 1" };
  }

  // Validate pattern-specific fields
  if (pattern.type === 'weekly' && (!pattern.daysOfWeek || pattern.daysOfWeek.length === 0)) {
    return { error: "Weekly recurrence requires 'daysOfWeek' to be specified" };
  }

  if (pattern.type === 'absoluteMonthly' && !pattern.dayOfMonth) {
    return { error: "Absolute monthly recurrence requires 'dayOfMonth' to be specified" };
  }

  if (pattern.type === 'relativeMonthly' && (!pattern.daysOfWeek || !pattern.index)) {
    return { error: "Relative monthly recurrence requires 'daysOfWeek' and 'index' to be specified" };
  }

  if (pattern.type === 'absoluteYearly' && (!pattern.dayOfMonth || !pattern.month)) {
    return { error: "Absolute yearly recurrence requires 'dayOfMonth' and 'month' to be specified" };
  }

  if (pattern.type === 'relativeYearly' && (!pattern.daysOfWeek || !pattern.index || !pattern.month)) {
    return { error: "Relative yearly recurrence requires 'daysOfWeek', 'index', and 'month' to be specified" };
  }

  // Validate range
  if (!range || !range.type || !range.startDate) {
    return { error: "Recurrence range must include 'type' and 'startDate'" };
  }

  if (range.type === 'endDate' && !range.endDate) {
    return { error: "Recurrence range type 'endDate' requires 'endDate' to be specified" };
  }

  if (range.type === 'numbered') {
    if (!range.numberOfOccurrences && range.numberOfOccurrences !== 0) {
      return { error: "Recurrence range type 'numbered' requires 'numberOfOccurrences' to be specified" };
    }
    if (range.numberOfOccurrences < 1) {
      return { error: "Number of occurrences must be at least 1" };
    }
  }

  return null;
}

/**
 * Create event handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateEvent(args) {
  const { subject, start, end, calendar, attendees, body, recurrence } = args;

  if (!subject || !start || !end) {
    return {
      content: [{
        type: "text",
        text: "Subject, start, and end times are required to create an event."
      }]
    };
  }

  // Validate recurrence if provided
  if (recurrence) {
    const validationError = validateRecurrence(recurrence);
    if (validationError) {
      return {
        content: [{
          type: "text",
          text: `Invalid recurrence pattern: ${validationError.error}`
        }]
      };
    }
  }

  try {
    // Get access token
    const accessToken = await ensureAuthenticated();

    // Resolve calendar path and build events endpoint
    const calendarPath = await resolveCalendarPath(accessToken, calendar);
    const endpoint = `${calendarPath}/events`;

    // Request body
    const bodyContent = {
      subject,
      start: { dateTime: start.dateTime || start, timeZone: start.timeZone || DEFAULT_TIMEZONE },
      end: { dateTime: end.dateTime || end, timeZone: end.timeZone || DEFAULT_TIMEZONE },
      attendees: attendees?.map(email => ({ emailAddress: { address: email }, type: "required" })),
      body: { contentType: "HTML", content: body || "" }
    };

    // Add recurrence if provided
    if (recurrence) {
      bodyContent.recurrence = recurrence;
    }

    // Make API call
    const response = await callGraphAPI(accessToken, 'POST', endpoint, bodyContent);

    const eventType = recurrence ? 'recurring event' : 'event';
    return {
      content: [{
        type: "text",
        text: `${eventType.charAt(0).toUpperCase() + eventType.slice(1)} '${subject}' has been successfully created.`
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
        text: `Error creating event: ${error.message}`
      }]
    };
  }
}

module.exports = handleCreateEvent;