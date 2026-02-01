/**
 * Calendar module for Outlook MCP server
 */
const handleListEvents = require('./list');
const handleListCalendars = require('./list-calendars');
const handleRespondToEvent = require('./respond');
const handleCreateEvent = require('./create');
const handleCancelEvent = require('./cancel');
const handleDeleteEvent = require('./delete');

// Calendar tool definitions
const calendarTools = [
  {
    name: "list-calendars",
    description: "Lists all calendars in your Outlook account",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleListCalendars
  },
  {
    name: "list-events",
    description: "Lists events from your calendar within a date range. Useful for checking conflicts before accepting meeting invites.",
    inputSchema: {
      type: "object",
      properties: {
        calendar: {
          type: "string",
          description: "Calendar name or ID (default: primary calendar). Use 'list-calendars' to see available calendars."
        },
        startDateTime: {
          type: "string",
          description: "Start of date range in ISO 8601 format (default: now)"
        },
        endDateTime: {
          type: "string",
          description: "End of date range in ISO 8601 format (default: 30 days from now)"
        },
        count: {
          type: "number",
          description: "Maximum number of events to retrieve (default: 10, max: 50)"
        }
      },
      required: []
    },
    handler: handleListEvents
  },
  {
    name: "respond-to-event",
    description: "Responds to a calendar event invitation (accept, decline, or tentatively accept)",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "The ID of the event to respond to"
        },
        response: {
          type: "string",
          description: "Your response to the invitation",
          enum: ["accept", "decline", "tentative"]
        },
        comment: {
          type: "string",
          description: "Optional message to send to the organizer"
        },
        sendResponse: {
          type: "boolean",
          description: "Whether to notify the organizer of your response (default: true)"
        }
      },
      required: ["eventId", "response"]
    },
    handler: handleRespondToEvent
  },
  {
    name: "create-event",
    description: "Creates a new calendar event (one-time or recurring)",
    inputSchema: {
      type: "object",
      properties: {
        subject: {
          type: "string",
          description: "The subject of the event"
        },
        start: {
          type: "string",
          description: "The start time of the event in ISO 8601 format"
        },
        end: {
          type: "string",
          description: "The end time of the event in ISO 8601 format"
        },
        calendar: {
          type: "string",
          description: "Calendar name or ID (default: primary calendar). Use 'list-calendars' to see available calendars."
        },
        attendees: {
          type: "array",
          items: {
            type: "string"
          },
          description: "List of attendee email addresses"
        },
        body: {
          type: "string",
          description: "Optional body content for the event"
        },
        recurrence: {
          type: "object",
          description: "Recurrence pattern for the event (optional). If not provided, creates a one-time event.",
          properties: {
            pattern: {
              type: "object",
              description: "The recurrence pattern",
              properties: {
                type: {
                  type: "string",
                  enum: ["daily", "weekly", "absoluteMonthly", "relativeMonthly", "absoluteYearly", "relativeYearly"],
                  description: "The recurrence pattern type"
                },
                interval: {
                  type: "number",
                  description: "Number of units between occurrences (e.g., 1 for every day, 2 for every other week)"
                },
                daysOfWeek: {
                  type: "array",
                  items: {
                    type: "string",
                    enum: ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"]
                  },
                  description: "Days of the week for weekly recurrence"
                },
                dayOfMonth: {
                  type: "number",
                  description: "Day of the month (1-31) for monthly recurrence"
                },
                month: {
                  type: "number",
                  description: "Month (1-12) for yearly recurrence"
                },
                firstDayOfWeek: {
                  type: "string",
                  enum: ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"],
                  description: "First day of the week (default: sunday)"
                },
                index: {
                  type: "string",
                  enum: ["first", "second", "third", "fourth", "last"],
                  description: "Week index in the month for relative monthly/yearly recurrence"
                }
              },
              required: ["type", "interval"]
            },
            range: {
              type: "object",
              description: "The recurrence range",
              properties: {
                type: {
                  type: "string",
                  enum: ["endDate", "noEnd", "numbered"],
                  description: "The recurrence range type"
                },
                startDate: {
                  type: "string",
                  description: "Start date in YYYY-MM-DD format"
                },
                endDate: {
                  type: "string",
                  description: "End date in YYYY-MM-DD format (required if type is 'endDate')"
                },
                numberOfOccurrences: {
                  type: "number",
                  description: "Number of occurrences (required if type is 'numbered')"
                }
              },
              required: ["type", "startDate"]
            }
          },
          required: ["pattern", "range"]
        }
      },
      required: ["subject", "start", "end"]
    },
    handler: handleCreateEvent
  },
  {
    name: "cancel-event",
    description: "Cancels a calendar event",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "The ID of the event to cancel"
        },
        comment: {
          type: "string",
          description: "Optional comment for cancelling the event"
        }
      },
      required: ["eventId"]
    },
    handler: handleCancelEvent
  },
  {
    name: "delete-event",
    description: "Deletes a calendar event",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "The ID of the event to delete"
        }
      },
      required: ["eventId"]
    },
    handler: handleDeleteEvent
  }
];

module.exports = {
  calendarTools,
  handleListEvents,
  handleListCalendars,
  handleRespondToEvent,
  handleCreateEvent,
  handleCancelEvent,
  handleDeleteEvent
};
