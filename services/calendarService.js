const axios = require('axios');
const moment = require('moment-timezone');
const authService = require('./authService');
require('dotenv').config();

class CalendarService {
    constructor() {
        this.graphApiUrl = process.env.GRAPH_API_URL;
        this.userEmail = process.env.USER_EMAIL;
    }

   async makeGraphRequest(method, endpoint, data = null) {
    try {
        const accessToken = await authService.getAccessToken();
        console.log("Access Token:", accessToken);  // Fixed variable name here

        const config = {
            method: method,
            url: `${this.graphApiUrl}${endpoint}`,
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            }
        };

        if (data) {
            config.data = data;
        }

        const response = await axios(config);
        return response.data;
    } catch (error) {
        console.error('Graph API Error:', error.response?.data || error.message);
        throw error;
    }
}

    parseDateTime(dateTimeString) {
        // Parse various date formats like "August 9 at 6am", "Aug 9, 2024 6:00 AM", etc.
        const currentYear = new Date().getFullYear();
        let parsedDate;

        // Handle "August 9 at 6am" format
        if (dateTimeString.includes('at')) {
            const [datePart, timePart] = dateTimeString.split(' at ');
            const dateWithYear = `${datePart} ${currentYear}`;
            parsedDate = moment(`${dateWithYear} ${timePart}`, 'MMMM D YYYY hA');
        } else {
            // Try other common formats
            parsedDate = moment(dateTimeString, [
                'MMMM D YYYY h:mm A',
                'MMM D YYYY h:mm A',
                'YYYY-MM-DD HH:mm',
                'MM/DD/YYYY h:mm A'
            ]);
        }

        if (!parsedDate.isValid()) {
            throw new Error('Invalid date format. Please use format like "August 9 at 6am"');
        }

        return parsedDate;
    }

    async checkAvailability(dateTimeString, durationMinutes = 60) {
        try {
            const startTime = this.parseDateTime(dateTimeString);
            const endTime = startTime.clone().add(durationMinutes, 'minutes');

            // Get calendar view for the specified time range
            const startISO = startTime.toISOString();
            const endISO = endTime.toISOString();

            const calendarView = await this.makeGraphRequest(
                'GET',
                `/users/${this.userEmail}/calendarView?startDateTime=${startISO}&endDateTime=${endISO}`
            );

            // Check if there are any conflicting events
            const conflictingEvents = calendarView.value.filter(event => {
                const eventStart = moment(event.start.dateTime);
                const eventEnd = moment(event.end.dateTime);
                
                // Check for overlap
                return (startTime.isBefore(eventEnd) && endTime.isAfter(eventStart));
            });

            return {
                available: conflictingEvents.length === 0,
                conflictingEvents: conflictingEvents.map(event => ({
                    subject: event.subject,
                    start: event.start.dateTime,
                    end: event.end.dateTime
                })),
                requestedSlot: {
                    start: startISO,
                    end: endISO
                }
            };
        } catch (error) {
            console.error('Error checking availability:', error);
            throw error;
        }
    }

    async bookAppointment(dateTimeString, subject, durationMinutes = 60, attendeeEmail = null) {
        try {
            // First check availability
            const availability = await this.checkAvailability(dateTimeString, durationMinutes);
            
            if (!availability.available) {
                return {
                    success: false,
                    message: 'This time slot is not available. Please try a different time.',
                    conflictingEvents: availability.conflictingEvents
                };
            }

            const startTime = this.parseDateTime(dateTimeString);
            const endTime = startTime.clone().add(durationMinutes, 'minutes');

            // Create event object
            const event = {
                subject: subject,
                start: {
                    dateTime: startTime.toISOString(),
                    timeZone: 'UTC'
                },
                end: {
                    dateTime: endTime.toISOString(),
                    timeZone: 'UTC'
                },
                body: {
                    contentType: 'text',
                    content: `Appointment booked via API on ${moment().format('YYYY-MM-DD HH:mm:ss')}`
                }
            };

            // Add attendee if provided
            if (attendeeEmail) {
                event.attendees = [{
                    emailAddress: {
                        address: attendeeEmail,
                        name: attendeeEmail
                    }
                }];
            }

            // Create the event
            const createdEvent = await this.makeGraphRequest(
                'POST',
                `/users/${this.userEmail}/events`,
                event
            );

            return {
                success: true,
                message: 'Appointment booked successfully!',
                event: {
                    id: createdEvent.id,
                    subject: createdEvent.subject,
                    start: createdEvent.start.dateTime,
                    end: createdEvent.end.dateTime,
                    webLink: createdEvent.webLink
                }
            };
        } catch (error) {
            console.error('Error booking appointment:', error);
            throw error;
        }
    }

    async getUpcomingEvents(days = 7) {
        try {
            const startTime = moment().toISOString();
            const endTime = moment().add(days, 'days').toISOString();

            const calendarView = await this.makeGraphRequest(
                'GET',
                `/users/${this.userEmail}/calendarView?startDateTime=${startTime}&endDateTime=${endTime}&$orderby=start/dateTime`
            );

            return calendarView.value.map(event => ({
                id: event.id,
                subject: event.subject,
                start: event.start.dateTime,
                end: event.end.dateTime,
                location: event.location?.displayName || 'No location'
            }));
        } catch (error) {
            console.error('Error getting upcoming events:', error);
            throw error;
        }
    }
}

module.exports = new CalendarService();