const express = require('express');
const axios = require('axios');
const moment = require('moment');
const { ConfidentialClientApplication } = require('@azure/msal-node');
require('dotenv').config();

const app = express();
app.use(express.json());
app.use(express.static('public'));

// MSAL Configuration for Personal Accounts
const msalConfig = {
    auth: {
        clientId: process.env.CLIENT_ID,
        clientSecret: process.env.CLIENT_SECRET,
        authority: 'https://login.microsoftonline.com/common' // This works for personal accounts
    },
    system: {
        loggerOptions: {
            loggerCallback(loglevel, message, containsPii) {
                console.log(message);
            },
            piiLoggingEnabled: false,
            logLevel: 3, // Info level
        }
    }
};

const pca = new ConfidentialClientApplication(msalConfig);

// Store user tokens
let userTokens = {};

// Microsoft Graph API endpoint
const GRAPH_API_ENDPOINT = 'https://graph.microsoft.com/v1.0';

// Route 1: Start OAuth flow
app.get('/auth', async (req, res) => {
    try {
        const authCodeUrlParameters = {
            scopes: [
                'https://graph.microsoft.com/Calendars.ReadWrite',
                'https://graph.microsoft.com/Mail.Send',
                'https://graph.microsoft.com/User.Read'
            ],
            redirectUri: process.env.REDIRECT_URI,
        };

        const authUrl = await pca.getAuthCodeUrl(authCodeUrlParameters);
        console.log('🔗 Redirecting to Microsoft Auth:', authUrl);
        res.redirect(authUrl);
    } catch (error) {
        console.error('❌ Auth URL generation error:', error);
        res.status(500).send('Failed to generate auth URL');
    }
});

// Route 2: Handle OAuth callback
app.get('/auth/callback', async (req, res) => {
    const { code, state } = req.query;
    
    if (!code) {
        return res.status(400).send('Authorization code not found');
    }

    try {
        const tokenRequest = {
            code: code,
            scopes: [
                'https://graph.microsoft.com/Calendars.ReadWrite',
                'https://graph.microsoft.com/Mail.Send',
                'https://graph.microsoft.com/User.Read'
            ],
            redirectUri: process.env.REDIRECT_URI,
        };

        const response = await pca.acquireTokenByCode(tokenRequest);
        const accessToken = response.accessToken;
        
        // Get user info
        const userResponse = await axios.get(`${GRAPH_API_ENDPOINT}/me`, {
            headers: {
                'Authorization': `Bearer ${accessToken}`
            }
        });

        const userId = userResponse.data.id;
        userTokens[userId] = {
            accessToken: accessToken,
            account: response.account,
            expiresOn: response.expiresOn
        };

        console.log('✅ Authentication successful for user:', userResponse.data.displayName);
        
        res.send(`
            <!DOCTYPE html>
            <html>
            <head>
                <title>Authentication Success</title>
                <style>
                    body { font-family: Arial, sans-serif; margin: 40px; background: #f5f5f5; }
                    .container { background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); max-width: 600px; margin: 0 auto; }
                    .success { color: #28a745; }
                    .user-id { background: #e9ecef; padding: 10px; border-radius: 4px; font-family: monospace; word-break: break-all; margin: 10px 0; }
                    .button { display: inline-block; background: #007bff; color: white; text-decoration: none; padding: 12px 24px; border-radius: 4px; margin-top: 15px; }
                    .button:hover { background: #0056b3; }
                </style>
            </head>
            <body>
                <div class="container">
                    <h2 class="success">✅ Authentication Successful!</h2>
                    <p><strong>Welcome, ${userResponse.data.displayName}!</strong></p>
                    <p>Email: ${userResponse.data.mail || userResponse.data.userPrincipalName}</p>
                    <p>You can now use the calendar integration.</p>
                    
                    <p><strong>Your User ID (copy this):</strong></p>
                    <div class="user-id">${userId}</div>
                    
                    <a href="/calendar-form" class="button">📅 Go to Calendar Booking Form</a>
                </div>
            </body>
            </html>
        `);
    } catch (error) {
        console.error('❌ Token acquisition error:', error);
        res.status(500).send(`
            <h2>❌ Authentication Failed</h2>
            <p>Error: ${error.message}</p>
            <p><a href="/auth">Try Again</a></p>
        `);
    }
});

// Get calendar ID by name
async function getCalendarId(accessToken, calendarName) {
    try {
        const response = await axios.get(`${GRAPH_API_ENDPOINT}/me/calendars`, {
            headers: {
                'Authorization': `Bearer ${accessToken}`
            }
        });

        console.log('📅 Available calendars:');
        response.data.value.forEach(cal => {
            console.log(`   - ${cal.name} (ID: ${cal.id})`);
        });

        const calendar = response.data.value.find(cal => 
            cal.name.toLowerCase() === calendarName.toLowerCase()
        );

        if (!calendar) {
            throw new Error(`Calendar "${calendarName}" not found. Available calendars: ${response.data.value.map(c => c.name).join(', ')}`);
        }

        console.log(`✅ Found calendar "${calendarName}" with ID: ${calendar.id}`);
        return calendar.id;
    } catch (error) {
        console.error('❌ Error getting calendar:', error.response?.data || error.message);
        throw error;
    }
}

// Check availability
async function checkAvailability(accessToken, calendarId, startTime, endTime) {
    try {
        const response = await axios.get(
            `${GRAPH_API_ENDPOINT}/me/calendars/${calendarId}/calendarView`,
            {
                headers: {
                    'Authorization': `Bearer ${accessToken}`
                },
                params: {
                    startDateTime: startTime,
                    endDateTime: endTime
                }
            }
        );

        const events = response.data.value;
        const isAvailable = events.length === 0;

        console.log(`🔍 Availability check: ${isAvailable ? 'Available' : 'Busy'} (${events.length} conflicts)`);

        return {
            isAvailable,
            conflictingEvents: events
        };
    } catch (error) {
        console.error('❌ Availability check error:', error.response?.data || error.message);
        throw error;
    }
}

// Get alternative time slots
async function getAlternativeSlots(accessToken, calendarId, requestedStartTime) {
    const alternatives = [];
    const requestedDate = moment(requestedStartTime);
    
    console.log('🔍 Looking for alternative slots...');
    
    // Check next 7 days
    for (let i = 0; i < 7; i++) {
        const checkDate = moment(requestedDate).add(i, 'days');
        
        // Check business hours (9 AM to 5 PM)
        for (let hour = 9; hour < 17; hour++) {
            const slotStart = checkDate.clone().hour(hour).minute(0).second(0);
            const slotEnd = slotStart.clone().add(1, 'hour');
            
            // Skip weekends
            if (slotStart.day() === 0 || slotStart.day() === 6) continue;
            
            try {
                const availability = await checkAvailability(
                    accessToken, 
                    calendarId, 
                    slotStart.toISOString(), 
                    slotEnd.toISOString()
                );
                
                if (availability.isAvailable) {
                    alternatives.push({
                        startTime: slotStart.toISOString(),
                        endTime: slotEnd.toISOString(),
                        displayTime: slotStart.format('MMMM Do YYYY, h:mm A')
                    });
                }
                
                if (alternatives.length >= 5) {
                    break;
                }
            } catch (error) {
                continue;
            }
        }
        
        if (alternatives.length >= 5) {
            break;
        }
    }
    
    console.log(`✅ Found ${alternatives.length} alternative slots`);
    return alternatives;
}

// Create calendar event
async function createEvent(accessToken, calendarId, eventData) {
    const event = {
        subject: eventData.subject,
        start: {
            dateTime: eventData.startTime,
            timeZone: 'UTC'
        },
        end: {
            dateTime: eventData.endTime,
            timeZone: 'UTC'
        },
        attendees: [
            {
                emailAddress: {
                    address: eventData.attendeeEmail,
                    name: eventData.attendeeName
                }
            }
        ],
        body: {
            contentType: 'html',
            content: `
                <h3>Meeting Details</h3>
                <p><strong>Subject:</strong> ${eventData.subject}</p>
                <p><strong>Attendee:</strong> ${eventData.attendeeName}</p>
                <p><strong>Email:</strong> ${eventData.attendeeEmail}</p>
                <p><strong>Time:</strong> ${moment(eventData.startTime).format('MMMM Do YYYY, h:mm A')} - ${moment(eventData.endTime).format('h:mm A')}</p>
            `
        }
    };

    try {
        const response = await axios.post(
            `${GRAPH_API_ENDPOINT}/me/calendars/${calendarId}/events`,
            event,
            {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                }
            }
        );

        console.log('✅ Event created successfully:', response.data.subject);
        return response.data;
    } catch (error) {
        console.error('❌ Event creation error:', error.response?.data || error.message);
        throw error;
    }
}

// Send confirmation email
async function sendConfirmationEmail(accessToken, toEmail, toName, eventDetails) {
    const message = {
        subject: `✅ Appointment Confirmed - ${eventDetails.subject}`,
        body: {
            contentType: 'html',
            content: `
                <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
                    <h2 style="color: #28a745;">✅ Appointment Confirmed!</h2>
                    
                    <p>Dear <strong>${toName}</strong>,</p>
                    
                    <p>Thank you for booking an appointment with us!</p>
                    
                    <div style="background: #f8f9fa; padding: 20px; border-radius: 8px; margin: 20px 0;">
                        <h3 style="color: #495057; margin-top: 0;">📅 Appointment Details</h3>
                        <p><strong>Subject:</strong> ${eventDetails.subject}</p>
                        <p><strong>Date & Time:</strong> ${moment(eventDetails.startTime).format('MMMM Do YYYY, h:mm A')} - ${moment(eventDetails.endTime).format('h:mm A')}</p>
                        <p><strong>Duration:</strong> ${moment(eventDetails.endTime).diff(moment(eventDetails.startTime), 'minutes')} minutes</p>
                    </div>
                    
                    <p>We look forward to meeting with you.</p>
                    
                    <p>Best regards,<br>
                    <strong>Your Calendar Integration Team</strong></p>
                </div>
            `
        },
        toRecipients: [
            {
                emailAddress: {
                    address: toEmail,
                    name: toName
                }
            }
        ]
    };

    try {
        await axios.post(
            `${GRAPH_API_ENDPOINT}/me/sendMail`,
            { message },
            {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                }
            }
        );

        console.log('✅ Confirmation email sent to:', toEmail);
    } catch (error) {
        console.error('❌ Email sending error:', error.response?.data || error.message);
    }
}

// Book appointment route
app.post('/book-appointment', async (req, res) => {
    const { userId, startTime, endTime, subject, attendeeEmail, attendeeName } = req.body;

    if (!userTokens[userId]) {
        return res.status(401).json({ error: 'User not authenticated. Please authenticate first.' });
    }

    const { accessToken } = userTokens[userId];

    try {
        console.log('📅 Starting booking process...');
        
        // Get calendar ID
        const calendarId = await getCalendarId(accessToken, process.env.CALENDAR_NAME);

        // Check availability
        const availability = await checkAvailability(accessToken, calendarId, startTime, endTime);
        
        if (!availability.isAvailable) {
            console.log('⚠️ Time slot not available, finding alternatives...');
            const alternatives = await getAlternativeSlots(accessToken, calendarId, startTime);
            
            return res.json({
                success: false,
                message: 'The requested time slot is not available.',
                alternatives: alternatives
            });
        }

        // Book the appointment
        console.log('📝 Creating calendar event...');
        const event = await createEvent(accessToken, calendarId, {
            subject,
            startTime,
            endTime,
            attendeeEmail,
            attendeeName
        });

        // Send confirmation email
        console.log('📧 Sending confirmation email...');
        await sendConfirmationEmail(accessToken, attendeeEmail, attendeeName, {
            subject,
            startTime,
            endTime
        });

        res.json({
            success: true,
            message: 'Appointment booked successfully!',
            eventId: event.id,
            eventUrl: event.webLink
        });

    } catch (error) {
        console.error('❌ Booking error:', error.response?.data || error.message);
        res.status(500).json({ 
            error: 'Failed to book appointment',
            details: error.message 
        });
    }
});

// Calendar form route
app.get('/calendar-form', (req, res) => {
    res.send(`
        <!DOCTYPE html>
        <html>
        <head>
            <title>📅 Calendar Booking System</title>
            <style>
                body { 
                    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
                    margin: 0; 
                    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                    min-height: 100vh;
                    padding: 20px;
                }
                .container {
                    max-width: 600px;
                    margin: 0 auto;
                    background: white;
                    padding: 40px;
                    border-radius: 12px;
                    box-shadow: 0 10px 30px rgba(0,0,0,0.2);
                }
                h1 { color: #333; text-align: center; margin-bottom: 30px; }
                .form-group { margin: 20px 0; }
                label { 
                    display: block; 
                    margin-bottom: 8px; 
                    font-weight: 600; 
                    color: #555;
                }
                input, textarea { 
                    width: 100%; 
                    padding: 12px; 
                    border: 2px solid #e1e5e9; 
                    border-radius: 6px; 
                    font-size: 16px;
                    box-sizing: border-box;
                    transition: border-color 0.3s;
                }
                input:focus { 
                    outline: none; 
                    border-color: #667eea; 
                }
                button { 
                    background: linear-gradient(45deg, #667eea, #764ba2); 
                    color: white; 
                    padding: 15px 30px; 
                    border: none; 
                    border-radius: 6px; 
                    cursor: pointer; 
                    font-size: 16px;
                    font-weight: 600;
                    width: 100%;
                    transition: transform 0.2s;
                }
                button:hover { 
                    transform: translateY(-2px); 
                }
                .result { 
                    margin-top: 30px; 
                    padding: 20px; 
                    border-radius: 8px; 
                    font-weight: 500;
                }
                .success { 
                    background: #d4edda; 
                    color: #155724; 
                    border: 1px solid #c3e6cb; 
                }
                .error { 
                    background: #f8d7da; 
                    color: #721c24; 
                    border: 1px solid #f5c6cb; 
                }
                .alternatives { 
                    background: #fff3cd; 
                    color: #856404; 
                    border: 1px solid #ffeaa7; 
                }
                .auth-notice {
                    background: #cce5ff;
                    padding: 15px;
                    border-radius: 6px;
                    margin-bottom: 20px;
                    border-left: 4px solid #007bff;
                }
                .user-id-input {
                    font-family: monospace;
                    background: #f8f9fa;
                }
            </style>
        </head>
        <body>
            <div class="container">
                <h1>📅 Book Your Appointment</h1>
                
                <div class="auth-notice">
                    <strong>📝 Note:</strong> You need to authenticate first to get your User ID. 
                    <a href="/auth" style="color: #007bff;">Click here to authenticate</a>
                </div>
                
                <form id="bookingForm">
                    <div class="form-group">
                        <label>🆔 User ID (from authentication page):</label>
                        <input type="text" id="userId" class="user-id-input" required placeholder="Paste your User ID here">
                    </div>
                    
                    <div class="form-group">
                        <label>📋 Meeting Subject:</label>
                        <input type="text" id="subject" required placeholder="e.g., Project Discussion">
                    </div>
                    
                    <div class="form-group">
                        <label>🕐 Start Date & Time:</label>
                        <input type="datetime-local" id="startTime" required>
                    </div>
                    
                    <div class="form-group">
                        <label>🕑 End Date & Time:</label>
                        <input type="datetime-local" id="endTime" required>
                    </div>
                    
                    <div class="form-group">
                        <label>👤 Your Name:</label>
                        <input type="text" id="attendeeName" required placeholder="John Doe">
                    </div>
                    
                    <div class="form-group">
                        <label>📧 Your Email:</label>
                        <input type="email" id="attendeeEmail" required placeholder="john@example.com">
                    </div>
                    
                    <button type="submit">📅 Book Appointment</button>
                </form>
                
                <div id="result"></div>
            </div>

            <script>
                // Set minimum datetime to current time
                const now = new Date();
                now.setMinutes(now.getMinutes() - now.getTimezoneOffset());
                const minDateTime = now.toISOString().slice(0, 16);
                document.getElementById('startTime').min = minDateTime;
                document.getElementById('endTime').min = minDateTime;

                // Auto-set end time when start time changes
                document.getElementById('startTime').addEventListener('change', function() {
                    const startTime = new Date(this.value);
                    const endTime = new Date(startTime.getTime() + 60 * 60 * 1000); // Add 1 hour
                    document.getElementById('endTime').value = endTime.toISOString().slice(0, 16);
                });

                document.getElementById('bookingForm').addEventListener('submit', async (e) => {
                    e.preventDefault();
                    
                    const formData = {
                        userId: document.getElementById('userId').value.trim(),
                        subject: document.getElementById('subject').value,
                        startTime: new Date(document.getElementById('startTime').value).toISOString(),
                        endTime: new Date(document.getElementById('endTime').value).toISOString(),
                        attendeeName: document.getElementById('attendeeName').value,
                        attendeeEmail: document.getElementById('attendeeEmail').value
                    };

                    const resultDiv = document.getElementById('result');
                    resultDiv.innerHTML = '<div class="result">⏳ Booking appointment, please wait...</div>';

                    try {
                        const response = await fetch('/book-appointment', {
                            method: 'POST',
                            headers: {
                                'Content-Type': 'application/json'
                            },
                            body: JSON.stringify(formData)
                        });

                        const result = await response.json();

                        if (result.success) {
                            resultDiv.innerHTML = \`
                                <div class="result success">
                                    <h3>✅ Appointment Booked Successfully!</h3>
                                    <p><strong>\${result.message}</strong></p>
                                    <p>📧 A confirmation email has been sent to your email address.</p>
                                    <p>🔗 Event ID: \${result.eventId}</p>
                                    \${result.eventUrl ? \`<p><a href="\${result.eventUrl}" target="_blank">📅 View in Outlook</a></p>\` : ''}
                                </div>
                            \`;
                            document.getElementById('bookingForm').reset();
                        } else {
                            let alternativesHtml = '';
                            if (result.alternatives && result.alternatives.length > 0) {
                                alternativesHtml = '<h4>📅 Available Alternative Times:</h4><ul>';
                                result.alternatives.forEach(alt => {
                                    alternativesHtml += \`<li><strong>\${alt.displayTime}</strong></li>\`;
                                });
                                alternativesHtml += '</ul>';
                            }

                            resultDiv.innerHTML = \`
                                <div class="result alternatives">
                                    <h3>⚠️ Time Slot Not Available</h3>
                                    <p>\${result.message}</p>
                                    \${alternativesHtml}
                                </div>
                            \`;
                        }
                    } catch (error) {
                        resultDiv.innerHTML = \`
                            <div class="result error">
                                <h3>❌ Booking Failed</h3>
                                <p>Error: \${error.message}</p>
                                <p>Please try again or contact support.</p>
                            </div>
                        \`;
                    }
                });
            </script>
        </body>
        </html>
    `);
});

// Home route
app.get('/', (req, res) => {
    res.send(`
        <!DOCTYPE html>
        <html>
        <head>
            <title>📅 Outlook Calendar Integration</title>
            <style>
                body { 
                    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
                    margin: 0; 
                    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                    min-height: 100vh;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                }
                .container {
                    text-align: center;
                    background: white;
                    padding: 50px;
                    border-radius: 12px;
                    box-shadow: 0 10px 30px rgba(0,0,0,0.2);
                    max-width: 500px;
                }
                h1 { color: #333; margin-bottom: 20px; }
                p { color: #666; font-size: 18px; margin-bottom: 30px; }
                .button { 
                    display: inline-block;
                    background: linear-gradient(45deg, #667eea, #764ba2); 
                    color: white; 
                    text-decoration: none;
                    padding: 15px 30px; 
                    border-radius: 6px; 
                    font-size: 16px;
                    font-weight: 600;
                    margin: 10px;
                    transition: transform 0.2s;
                }
                .button:hover { 
                    transform: translateY(-2px); 
                }
            </style>
        </head>
        <body>
            <div class="container">
                <h1>📅 Outlook Calendar Integration</h1>
                <p>Connect your Microsoft account to start booking appointments</p>
                <a href="/auth" class="button">🔐 Authenticate with Microsoft</a>
                <a href="/calendar-form" class="button">📝 Go to Booking Form</a>
            </div>
        </body>
        </html>
    `);
});

// Start server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 Server running on http://localhost:${PORT}`);
    console.log(`🏠 Home page: http://localhost:${PORT}`);
    console.log(`🔐 Authentication: http://localhost:${PORT}/auth`);
    console.log(`📝 Booking form: http://localhost:${PORT}/calendar-form`);
    console.log('');
    console.log('📋 Setup Instructions:');
    console.log('1. Make sure your .env file has the correct CLIENT_ID and CLIENT_SECRET');
    console.log('2. In Azure Portal, ensure your app supports "Personal Microsoft accounts"');
    console.log('3. Visit /auth to authenticate first, then use the booking form');
});