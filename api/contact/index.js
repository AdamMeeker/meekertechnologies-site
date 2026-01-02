// HTML encoding helper to prevent XSS in email bodies
function escapeHtml(text) {
    if (!text) return '';
    const htmlEntities = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#39;'
    };
    return String(text).replace(/[&<>"']/g, char => htmlEntities[char]);
}

// Microsoft Graph email sending using client credentials flow
async function getAccessToken() {
    const tenantId = process.env.AZURE_TENANT_ID;
    const clientId = process.env.AZURE_CLIENT_ID;
    const clientSecret = process.env.AZURE_CLIENT_SECRET;

    const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
    const params = new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials'
    });

    const response = await fetch(tokenUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: params.toString()
    });

    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Failed to get access token: ${response.status} - ${errorText}`);
    }

    const data = await response.json();
    return data.access_token;
}

async function sendEmail(accessToken, fromEmail, toEmail, subject, htmlBody) {
    const graphUrl = `https://graph.microsoft.com/v1.0/users/${fromEmail}/sendMail`;

    const emailData = {
        message: {
            subject: subject,
            body: {
                contentType: 'HTML',
                content: htmlBody
            },
            toRecipients: [
                {
                    emailAddress: {
                        address: toEmail
                    }
                }
            ]
        },
        saveToSentItems: false
    };

    const response = await fetch(graphUrl, {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(emailData)
    });

    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Failed to send email: ${response.status} - ${errorText}`);
    }

    return true;
}

module.exports = async function (context, req) {
    context.log('Contact form submission received');

    // CORS configuration
    const allowedOrigins = [
        'https://meekertechnologies.com',
        'https://www.meekertechnologies.com',
        'https://witty-hill-016919110.1.azurestaticapps.net'
    ];
    const origin = req.headers.origin || req.headers.Origin;
    const corsOrigin = allowedOrigins.includes(origin) ? origin : allowedOrigins[0];

    const corsHeaders = {
        'Access-Control-Allow-Origin': corsOrigin,
        'Access-Control-Allow-Methods': 'POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type'
    };

    // Handle CORS preflight
    if (req.method === 'OPTIONS') {
        context.res = {
            status: 204,
            headers: corsHeaders
        };
        return;
    }

    const headers = {
        'Content-Type': 'application/json',
        ...corsHeaders
    };

    try {
        const { name, email, company, interests, message, honeypot } = req.body || {};

        // Check honeypot - if filled, it's likely a bot
        if (honeypot) {
            context.log('Honeypot triggered - likely spam');
            context.res = {
                status: 200,
                headers,
                body: {
                    success: true,
                    message: 'Thank you for your inquiry. We will be in touch within 24 hours.'
                }
            };
            return;
        }

        // Validate required fields
        if (!name || !email || !message) {
            context.res = {
                status: 400,
                headers,
                body: { success: false, error: 'Name, email, and message are required' }
            };
            return;
        }

        // Validate email format
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (!emailRegex.test(email)) {
            context.res = {
                status: 400,
                headers,
                body: { success: false, error: 'Invalid email format' }
            };
            return;
        }

        // Format interests array
        const interestsText = Array.isArray(interests) && interests.length > 0
            ? interests.join(', ')
            : 'Not specified';

        // Log the submission
        context.log('Contact form submission:');
        context.log(`  Name: ${name}`);
        context.log(`  Email: ${email}`);
        context.log(`  Company: ${company || 'Not provided'}`);
        context.log(`  Interests: ${interestsText}`);
        context.log(`  Message: ${message}`);

        // Send email via Microsoft Graph
        const senderEmail = process.env.SENDER_EMAIL || 'adam@adammeeker.com';
        const ownerEmail = process.env.OWNER_EMAIL || 'adam@adammeeker.com';

        try {
            const accessToken = await getAccessToken();

            // HTML-encode all user input to prevent XSS attacks in email
            const safeName = escapeHtml(name);
            const safeEmail = escapeHtml(email);
            const safeCompany = escapeHtml(company || 'Not provided');
            const safeInterests = escapeHtml(interestsText);
            const safeMessage = escapeHtml(message).replace(/\n/g, '<br>');

            const subject = `[Meeker Technologies] ${safeInterests} - from ${safeName}`;
            const htmlBody = `
                <h2>New Contact Form Submission</h2>
                <p><strong>From:</strong> ${safeName} (<a href="mailto:${safeEmail}">${safeEmail}</a>)</p>
                <p><strong>Company:</strong> ${safeCompany}</p>
                <p><strong>Interested in:</strong> ${safeInterests}</p>
                <hr>
                <p><strong>Message:</strong></p>
                <p>${safeMessage}</p>
                <hr>
                <p><em>Sent from meekertechnologies.com contact form</em></p>
            `;

            await sendEmail(accessToken, senderEmail, ownerEmail, subject, htmlBody);
            context.log('Email sent successfully');

            context.res = {
                status: 200,
                headers,
                body: {
                    success: true,
                    message: 'Thank you for your inquiry. We will be in touch within 24 hours.'
                }
            };

        } catch (emailError) {
            context.log('Email sending failed:', emailError.message);
            // Return success anyway - we logged it, and we don't want to expose internal errors
            context.res = {
                status: 200,
                headers,
                body: {
                    success: true,
                    message: 'Thank you for your inquiry. We will be in touch within 24 hours.',
                    note: 'Email delivery may be delayed.'
                }
            };
        }

    } catch (error) {
        context.log('Error processing contact form:', error.message || error);
        context.res = {
            status: 500,
            headers,
            body: {
                success: false,
                error: 'Unable to process your request. Please try again later.'
            }
        };
    }
};
