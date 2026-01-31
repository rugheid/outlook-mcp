/**
 * List calendars functionality
 */
const config = require('../config');
const { ensureAuthenticated } = require('../auth');
const { getAllCalendars } = require('./calendar-utils');

/**
 * List calendars handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListCalendars(args) {
  try {
    // Get access token
    const accessToken = await ensureAuthenticated();

    // Get all calendars using utility
    const calendars = await getAllCalendars(accessToken);

    if (!calendars || calendars.length === 0) {
      return {
        content: [{
          type: "text",
          text: "No calendars found."
        }]
      };
    }

    // Format results
    const calendarList = calendars.map((calendar, index) => {
      const isDefault = calendar.isDefaultCalendar ? ' [DEFAULT]' : '';
      const owner = calendar.owner ? ` (Owner: ${calendar.owner.name})` : '';
      const permissions = [];
      if (calendar.canEdit) permissions.push('edit');
      if (calendar.canShare) permissions.push('share');
      if (calendar.canViewPrivateItems) permissions.push('view-private');
      const perms = permissions.length > 0 ? ` - Permissions: ${permissions.join(', ')}` : '';

      return `${index + 1}. ${calendar.name}${isDefault}${owner}${perms}\nID: ${calendar.id}\n`;
    }).join("\n");

    return {
      content: [{
        type: "text",
        text: `Found ${calendars.length} calendars:\n\n${calendarList}`
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
        text: `Error listing calendars: ${error.message}`
      }]
    };
  }
}

module.exports = handleListCalendars;
