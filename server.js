const express = require('express');
const calendarService = require('./services/calendarService');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(express.json());

// Routes

// Health check endpoint
app.get('/health', (req, res) => {
    res.json({ status: 'OK', message: 'Outlook Calendar API is running' });
});

// Check availability endpoint
app.post('/check-availability', async (req, res) => {
    try {
        const { dateTime, duration = 60 } = req.body;
        
        if (!dateTime) {
            return res.status(400).json({ 
                error: 'dateTime is required. Format: "August 9 at 6am"' 
            });
        }

        const availability = await calendarService.checkAvailability(dateTime, duration);
        
        res.json({
            available: availability.available,
            message: availability.available ? 
                'Time slot is available' : 
                'Time slot is not available',
            requestedSlot: availability.requestedSlot,
            conflictingEvents: availability.conflictingEvents
        });
    } catch (error) {
        res.status(500).json({ 
            error: error.message,
            details: 'Error checking availability'
        });
    }
});

// Book appointment endpoint
app.post('/book-appointment', async (req, res) => {
    try {
        const { dateTime, subject, duration = 60, attendeeEmail } = req.body;
        
        if (!dateTime || !subject) {
            return res.status(400).json({ 
                error: 'dateTime and subject are required' 
            });
        }

        const result = await calendarService.bookAppointment(
            dateTime, 
            subject, 
            duration, 
            attendeeEmail
        );
        
        if (result.success) {
            res.json({
                success: true,
                message: result.message,
                event: result.event
            });
        } else {
            res.status(409).json({
                success: false,
                message: result.message,
                conflictingEvents: result.conflictingEvents
            });
        }
    } catch (error) {
        res.status(500).json({ 
            error: error.message,
            details: 'Error booking appointment'
        });
    }
});

// Get upcoming events endpoint
app.get('/upcoming-events', async (req, res) => {
    try {
        const { days = 7 } = req.query;
        
        const events = await calendarService.getUpcomingEvents(parseInt(days));
        
        res.json({
            events: events,
            count: events.length
        });
    } catch (error) {
        res.status(500).json({ 
            error: error.message,
            details: 'Error getting upcoming events'
        });
    }
});

// Combined endpoint for your specific use case
app.post('/smart-book', async (req, res) => {
    try {
        const { dateTime, subject = 'Appointment', duration = 60, attendeeEmail } = req.body;
        
        if (!dateTime) {
            return res.status(400).json({ 
                error: 'dateTime is required. Format: "August 9 at 6am"' 
            });
        }

        // Check availability first
        const availability = await calendarService.checkAvailability(dateTime, duration);
        
        if (!availability.available) {
            return res.json({
                success: false,
                message: 'This time slot is booked. Please try a different appointment time.',
                conflictingEvents: availability.conflictingEvents,
                suggestion: 'Try booking 1 hour later or on a different day'
            });
        }

        // Book the appointment
        const result = await calendarService.bookAppointment(
            dateTime, 
            subject, 
            duration, 
            attendeeEmail
        );

        res.json({
            success: true,
            message: 'Appointment booked successfully!',
            event: result.event
        });
    } catch (error) {
        res.status(500).json({ 
            error: error.message,
            details: 'Error processing appointment request'
        });
    }
});

// Error handling middleware
app.use((error, req, res, next) => {
    console.error(error);
    res.status(500).json({ 
        error: 'Internal server error',
        details: error.message 
    });
});

// Start server
app.listen(PORT, () => {
    console.log(`🚀 Outlook Calendar API server running on port ${PORT}`);
    console.log(`📅 Health check: http://localhost:${PORT}/health`);
    console.log(`📖 API Endpoints:`);
    console.log(`   POST /smart-book - Your main endpoint`);
    console.log(`   POST /check-availability - Check time slot`);
    console.log(`   POST /book-appointment - Book appointment`);
    console.log(`   GET /upcoming-events - Get upcoming events`);
});

module.exports = app;