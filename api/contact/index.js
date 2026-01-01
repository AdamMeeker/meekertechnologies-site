module.exports = async function (context, req) {
    context.log('Contact form submission received');

    const { name, email, company, interests, message } = req.body || {};

    // Validate required fields
    if (!name || !email || !message) {
        context.res = {
            status: 400,
            body: { error: 'Name, email, and message are required' }
        };
        return;
    }

    // Validate email format
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email)) {
        context.res = {
            status: 400,
            body: { error: 'Invalid email format' }
        };
        return;
    }

    // Log the submission (in production, this would send to email service or CRM)
    context.log('Contact form data:', {
        name,
        email,
        company: company || 'Not provided',
        interests: interests || [],
        message,
        timestamp: new Date().toISOString()
    });

    // In production, integrate with:
    // - SendGrid, Mailgun, or Azure Communication Services for email
    // - Microsoft 365 / Graph API for Outlook integration
    // - CRM like Salesforce or HubSpot

    // For now, return success (email fallback in frontend handles actual delivery)
    context.res = {
        status: 200,
        headers: {
            'Content-Type': 'application/json'
        },
        body: {
            success: true,
            message: 'Thank you for your inquiry. We will be in touch within 24 hours.'
        }
    };
};
