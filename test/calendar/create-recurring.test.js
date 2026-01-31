const handleCreateEvent = require('../../calendar/create');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');
const { resolveCalendarPath } = require('../../calendar/calendar-utils');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');
jest.mock('../../calendar/calendar-utils');

describe('handleCreateEvent - Recurring Events', () => {
  const mockAccessToken = 'dummy_access_token';
  const mockCalendarPath = 'me/calendar';

  beforeEach(() => {
    callGraphAPI.mockClear();
    ensureAuthenticated.mockClear();
    resolveCalendarPath.mockClear();
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    resolveCalendarPath.mockResolvedValue(mockCalendarPath);
  });

  describe('daily recurrence', () => {
    test('should create daily recurring event', async () => {
      const args = {
        subject: 'Daily Standup',
        start: '2024-03-10T09:00:00',
        end: '2024-03-10T09:30:00',
        recurrence: {
          pattern: {
            type: 'daily',
            interval: 1
          },
          range: {
            type: 'endDate',
            startDate: '2024-03-10',
            endDate: '2024-12-31'
          }
        }
      };

      callGraphAPI.mockResolvedValue({ id: 'event-123' });

      const result = await handleCreateEvent(args);

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        `${mockCalendarPath}/events`,
        expect.objectContaining({
          subject: 'Daily Standup',
          recurrence: args.recurrence
        })
      );

      expect(result.content[0].text).toContain('Recurring event');
      expect(result.content[0].text).toContain('Daily Standup');
    });

    test('should create every-other-day recurring event', async () => {
      const args = {
        subject: 'Every Other Day',
        start: '2024-03-10T10:00:00',
        end: '2024-03-10T11:00:00',
        recurrence: {
          pattern: {
            type: 'daily',
            interval: 2
          },
          range: {
            type: 'numbered',
            startDate: '2024-03-10',
            numberOfOccurrences: 10
          }
        }
      };

      callGraphAPI.mockResolvedValue({ id: 'event-123' });

      const result = await handleCreateEvent(args);

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        `${mockCalendarPath}/events`,
        expect.objectContaining({
          recurrence: expect.objectContaining({
            pattern: expect.objectContaining({
              interval: 2
            }),
            range: expect.objectContaining({
              numberOfOccurrences: 10
            })
          })
        })
      );

      expect(result.content[0].text).toContain('successfully created');
    });
  });

  describe('weekly recurrence', () => {
    test('should create weekly recurring event', async () => {
      const args = {
        subject: 'Team Meeting',
        start: '2024-03-11T14:00:00',
        end: '2024-03-11T15:00:00',
        recurrence: {
          pattern: {
            type: 'weekly',
            interval: 1,
            daysOfWeek: ['monday', 'wednesday', 'friday']
          },
          range: {
            type: 'noEnd',
            startDate: '2024-03-11'
          }
        }
      };

      callGraphAPI.mockResolvedValue({ id: 'event-123' });

      const result = await handleCreateEvent(args);

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        `${mockCalendarPath}/events`,
        expect.objectContaining({
          recurrence: expect.objectContaining({
            pattern: expect.objectContaining({
              type: 'weekly',
              daysOfWeek: ['monday', 'wednesday', 'friday']
            })
          })
        })
      );

      expect(result.content[0].text).toContain('successfully created');
    });

    test('should create bi-weekly recurring event', async () => {
      const args = {
        subject: 'Sprint Planning',
        start: '2024-03-11T10:00:00',
        end: '2024-03-11T11:00:00',
        recurrence: {
          pattern: {
            type: 'weekly',
            interval: 2,
            daysOfWeek: ['monday']
          },
          range: {
            type: 'numbered',
            startDate: '2024-03-11',
            numberOfOccurrences: 26
          }
        }
      };

      callGraphAPI.mockResolvedValue({ id: 'event-123' });

      const result = await handleCreateEvent(args);

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        `${mockCalendarPath}/events`,
        expect.objectContaining({
          recurrence: expect.objectContaining({
            pattern: expect.objectContaining({
              interval: 2
            })
          })
        })
      );

      expect(result.content[0].text).toContain('successfully created');
    });

    test('should reject weekly recurrence without daysOfWeek', async () => {
      const args = {
        subject: 'Invalid Weekly',
        start: '2024-03-11T10:00:00',
        end: '2024-03-11T11:00:00',
        recurrence: {
          pattern: {
            type: 'weekly',
            interval: 1
          },
          range: {
            type: 'noEnd',
            startDate: '2024-03-11'
          }
        }
      };

      const result = await handleCreateEvent(args);

      expect(callGraphAPI).not.toHaveBeenCalled();
      expect(result.content[0].text).toContain('Invalid recurrence pattern');
      expect(result.content[0].text).toContain('daysOfWeek');
    });
  });

  describe('monthly recurrence', () => {
    test('should create absolute monthly recurring event', async () => {
      const args = {
        subject: 'Monthly Report',
        start: '2024-03-15T09:00:00',
        end: '2024-03-15T10:00:00',
        recurrence: {
          pattern: {
            type: 'absoluteMonthly',
            interval: 1,
            dayOfMonth: 15
          },
          range: {
            type: 'endDate',
            startDate: '2024-03-15',
            endDate: '2024-12-15'
          }
        }
      };

      callGraphAPI.mockResolvedValue({ id: 'event-123' });

      const result = await handleCreateEvent(args);

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        `${mockCalendarPath}/events`,
        expect.objectContaining({
          recurrence: expect.objectContaining({
            pattern: expect.objectContaining({
              type: 'absoluteMonthly',
              dayOfMonth: 15
            })
          })
        })
      );

      expect(result.content[0].text).toContain('successfully created');
    });

    test('should create relative monthly recurring event', async () => {
      const args = {
        subject: 'First Monday Meeting',
        start: '2024-03-04T14:00:00',
        end: '2024-03-04T15:00:00',
        recurrence: {
          pattern: {
            type: 'relativeMonthly',
            interval: 1,
            daysOfWeek: ['monday'],
            index: 'first'
          },
          range: {
            type: 'noEnd',
            startDate: '2024-03-04'
          }
        }
      };

      callGraphAPI.mockResolvedValue({ id: 'event-123' });

      const result = await handleCreateEvent(args);

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        `${mockCalendarPath}/events`,
        expect.objectContaining({
          recurrence: expect.objectContaining({
            pattern: expect.objectContaining({
              type: 'relativeMonthly',
              daysOfWeek: ['monday'],
              index: 'first'
            })
          })
        })
      );

      expect(result.content[0].text).toContain('successfully created');
    });

    test('should reject absolute monthly recurrence without dayOfMonth', async () => {
      const args = {
        subject: 'Invalid Monthly',
        start: '2024-03-15T09:00:00',
        end: '2024-03-15T10:00:00',
        recurrence: {
          pattern: {
            type: 'absoluteMonthly',
            interval: 1
          },
          range: {
            type: 'noEnd',
            startDate: '2024-03-15'
          }
        }
      };

      const result = await handleCreateEvent(args);

      expect(callGraphAPI).not.toHaveBeenCalled();
      expect(result.content[0].text).toContain('Invalid recurrence pattern');
      expect(result.content[0].text).toContain('dayOfMonth');
    });

    test('should reject relative monthly recurrence without index', async () => {
      const args = {
        subject: 'Invalid Relative Monthly',
        start: '2024-03-04T14:00:00',
        end: '2024-03-04T15:00:00',
        recurrence: {
          pattern: {
            type: 'relativeMonthly',
            interval: 1,
            daysOfWeek: ['monday']
          },
          range: {
            type: 'noEnd',
            startDate: '2024-03-04'
          }
        }
      };

      const result = await handleCreateEvent(args);

      expect(callGraphAPI).not.toHaveBeenCalled();
      expect(result.content[0].text).toContain('Invalid recurrence pattern');
      expect(result.content[0].text).toContain('index');
    });
  });

  describe('yearly recurrence', () => {
    test('should create absolute yearly recurring event', async () => {
      const args = {
        subject: 'Annual Review',
        start: '2024-03-15T09:00:00',
        end: '2024-03-15T10:00:00',
        recurrence: {
          pattern: {
            type: 'absoluteYearly',
            interval: 1,
            dayOfMonth: 15,
            month: 3
          },
          range: {
            type: 'noEnd',
            startDate: '2024-03-15'
          }
        }
      };

      callGraphAPI.mockResolvedValue({ id: 'event-123' });

      const result = await handleCreateEvent(args);

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        `${mockCalendarPath}/events`,
        expect.objectContaining({
          recurrence: expect.objectContaining({
            pattern: expect.objectContaining({
              type: 'absoluteYearly',
              dayOfMonth: 15,
              month: 3
            })
          })
        })
      );

      expect(result.content[0].text).toContain('successfully created');
    });

    test('should create relative yearly recurring event', async () => {
      const args = {
        subject: 'Thanksgiving',
        start: '2024-11-28T12:00:00',
        end: '2024-11-28T18:00:00',
        recurrence: {
          pattern: {
            type: 'relativeYearly',
            interval: 1,
            daysOfWeek: ['thursday'],
            index: 'fourth',
            month: 11
          },
          range: {
            type: 'noEnd',
            startDate: '2024-11-28'
          }
        }
      };

      callGraphAPI.mockResolvedValue({ id: 'event-123' });

      const result = await handleCreateEvent(args);

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        `${mockCalendarPath}/events`,
        expect.objectContaining({
          recurrence: expect.objectContaining({
            pattern: expect.objectContaining({
              type: 'relativeYearly',
              daysOfWeek: ['thursday'],
              index: 'fourth',
              month: 11
            })
          })
        })
      );

      expect(result.content[0].text).toContain('successfully created');
    });

    test('should reject absolute yearly recurrence without month', async () => {
      const args = {
        subject: 'Invalid Yearly',
        start: '2024-03-15T09:00:00',
        end: '2024-03-15T10:00:00',
        recurrence: {
          pattern: {
            type: 'absoluteYearly',
            interval: 1,
            dayOfMonth: 15
          },
          range: {
            type: 'noEnd',
            startDate: '2024-03-15'
          }
        }
      };

      const result = await handleCreateEvent(args);

      expect(callGraphAPI).not.toHaveBeenCalled();
      expect(result.content[0].text).toContain('Invalid recurrence pattern');
      expect(result.content[0].text).toContain('month');
    });
  });

  describe('recurrence range validation', () => {
    test('should reject recurrence without range type', async () => {
      const args = {
        subject: 'Invalid Range',
        start: '2024-03-10T09:00:00',
        end: '2024-03-10T10:00:00',
        recurrence: {
          pattern: {
            type: 'daily',
            interval: 1
          },
          range: {
            startDate: '2024-03-10'
          }
        }
      };

      const result = await handleCreateEvent(args);

      expect(callGraphAPI).not.toHaveBeenCalled();
      expect(result.content[0].text).toContain('Invalid recurrence pattern');
      expect(result.content[0].text).toContain("'type'");
    });

    test('should reject endDate range without endDate', async () => {
      const args = {
        subject: 'Invalid EndDate',
        start: '2024-03-10T09:00:00',
        end: '2024-03-10T10:00:00',
        recurrence: {
          pattern: {
            type: 'daily',
            interval: 1
          },
          range: {
            type: 'endDate',
            startDate: '2024-03-10'
          }
        }
      };

      const result = await handleCreateEvent(args);

      expect(callGraphAPI).not.toHaveBeenCalled();
      expect(result.content[0].text).toContain('Invalid recurrence pattern');
      expect(result.content[0].text).toContain('endDate');
    });

    test('should reject numbered range without numberOfOccurrences', async () => {
      const args = {
        subject: 'Invalid Numbered',
        start: '2024-03-10T09:00:00',
        end: '2024-03-10T10:00:00',
        recurrence: {
          pattern: {
            type: 'daily',
            interval: 1
          },
          range: {
            type: 'numbered',
            startDate: '2024-03-10'
          }
        }
      };

      const result = await handleCreateEvent(args);

      expect(callGraphAPI).not.toHaveBeenCalled();
      expect(result.content[0].text).toContain('Invalid recurrence pattern');
      expect(result.content[0].text).toContain('numberOfOccurrences');
    });

    test('should reject numbered range with invalid numberOfOccurrences', async () => {
      const args = {
        subject: 'Invalid Count',
        start: '2024-03-10T09:00:00',
        end: '2024-03-10T10:00:00',
        recurrence: {
          pattern: {
            type: 'daily',
            interval: 1
          },
          range: {
            type: 'numbered',
            startDate: '2024-03-10',
            numberOfOccurrences: 0
          }
        }
      };

      const result = await handleCreateEvent(args);

      expect(callGraphAPI).not.toHaveBeenCalled();
      expect(result.content[0].text).toContain('Invalid recurrence pattern');
      expect(result.content[0].text).toContain('at least 1');
    });
  });

  describe('calendar integration', () => {
    test('should create recurring event in specific calendar', async () => {
      const customCalendarPath = 'me/calendars/calendar-123';
      resolveCalendarPath.mockResolvedValue(customCalendarPath);

      const args = {
        subject: 'Team Sync',
        start: '2024-03-11T10:00:00',
        end: '2024-03-11T11:00:00',
        calendar: 'Work Calendar',
        recurrence: {
          pattern: {
            type: 'weekly',
            interval: 1,
            daysOfWeek: ['monday']
          },
          range: {
            type: 'noEnd',
            startDate: '2024-03-11'
          }
        }
      };

      callGraphAPI.mockResolvedValue({ id: 'event-123' });

      const result = await handleCreateEvent(args);

      expect(resolveCalendarPath).toHaveBeenCalledWith(mockAccessToken, 'Work Calendar');
      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        `${customCalendarPath}/events`,
        expect.any(Object)
      );

      expect(result.content[0].text).toContain('successfully created');
    });
  });

  describe('one-time event compatibility', () => {
    test('should still create one-time event when recurrence not provided', async () => {
      const args = {
        subject: 'One-Time Meeting',
        start: '2024-03-10T14:00:00',
        end: '2024-03-10T15:00:00'
      };

      callGraphAPI.mockResolvedValue({ id: 'event-123' });

      const result = await handleCreateEvent(args);

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        `${mockCalendarPath}/events`,
        expect.not.objectContaining({
          recurrence: expect.anything()
        })
      );

      expect(result.content[0].text).toContain('Event');
      expect(result.content[0].text).not.toContain('Recurring');
      expect(result.content[0].text).toContain('successfully created');
    });
  });

  describe('error handling', () => {
    test('should reject recurrence with invalid interval', async () => {
      const args = {
        subject: 'Invalid Interval',
        start: '2024-03-10T09:00:00',
        end: '2024-03-10T10:00:00',
        recurrence: {
          pattern: {
            type: 'daily',
            interval: 0
          },
          range: {
            type: 'noEnd',
            startDate: '2024-03-10'
          }
        }
      };

      const result = await handleCreateEvent(args);

      expect(callGraphAPI).not.toHaveBeenCalled();
      expect(result.content[0].text).toContain('Invalid recurrence pattern');
      expect(result.content[0].text).toContain('interval');
    });

    test('should handle API errors gracefully', async () => {
      const args = {
        subject: 'API Error Test',
        start: '2024-03-10T09:00:00',
        end: '2024-03-10T10:00:00',
        recurrence: {
          pattern: {
            type: 'daily',
            interval: 1
          },
          range: {
            type: 'noEnd',
            startDate: '2024-03-10'
          }
        }
      };

      callGraphAPI.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleCreateEvent(args);

      expect(result.content[0].text).toContain('Error creating event');
      expect(result.content[0].text).toContain('Graph API Error');
    });
  });
});
