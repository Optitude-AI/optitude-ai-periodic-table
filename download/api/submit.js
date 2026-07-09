// Vercel serverless function — handles Periodic Table download form submissions
// Receives form POST, validates, sends email to Simon via Resend, redirects to /thank-you
//
// Required environment variable (set in Vercel project settings → Environment Variables):
//   RESEND_API_KEY = "re_xxxxxxxxxxxx"  (get one free at https://resend.com/api-keys)
//
// Once the env var is set, redeploy (push any commit or click Redeploy in Vercel dashboard).

const RECIPIENT_EMAIL = "simon.childs@optitude360.com";
// Using Resend's onboarding address — works without domain verification, but only sends to the
// email you signed up with (Simon's). All lead notifications go to Simon, so this is sufficient.
// To use a custom from-address (e.g. downloads@optitude360.com), verify the domain at
// https://resend.com/domains and update this constant.
const FROM_EMAIL = "Optitude AI <onboarding@resend.dev>";
const THANK_YOU_URL = "https://periodic-table.optitudeai.com/thank-you";
const ERROR_URL = "https://periodic-table.optitudeai.com/?error=1";

// Fields we accept — anything else is ignored (prevents header injection / spam)
const ALLOWED_FIELDS = new Set([
  "firstName", "lastName", "email", "company", "jobTitle",
  "companySize", "aiChallenge", "opportunityReview", "marketingConsent"
]);

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function buildEmailHtml(data) {
  const rows = [
    ["First name", data.firstName],
    ["Last name", data.lastName || "—"],
    ["Email", `<a href="mailto:${escapeHtml(data.email)}">${escapeHtml(data.email)}</a>`],
    ["Company", data.company],
    ["Job title", data.jobTitle || "—"],
    ["Company size", data.companySize || "—"],
    ["Main AI challenge", data.aiChallengeText || data.aiChallenge],
    ["AI Opportunity Review requested", data.opportunityReview === "yes" ? "Yes — please contact them to book a call" : "No"],
    ["Marketing consent", data.marketingConsent === "yes" ? "Yes" : "No"],
    ["Consent timestamp", new Date().toISOString()],
    ["Source", "periodic-table-landing-page"],
    ["Page URL", "https://periodic-table.optitudeai.com/"]
  ];

  const rowsHtml = rows.map(([label, value]) => `
    <tr>
      <td style="padding:10px 16px;background:#F7F4EC;border-bottom:1px solid #E5E0D2;font-weight:600;width:35%;vertical-align:top;color:#1A1A1A;">${escapeHtml(label)}</td>
      <td style="padding:10px 16px;border-bottom:1px solid #E5E0D2;vertical-align:top;color:#1F1F1F;">${value}</td>
    </tr>
  `).join("");

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>New Periodic Table download</title>
</head>
<body style="margin:0;padding:0;background:#F7F4EC;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;">
  <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background:#F7F4EC;padding:24px;">
    <tr>
      <td align="center">
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="600" style="max-width:600px;background:#FFFFFF;border-radius:8px;overflow:hidden;box-shadow:0 4px 16px rgba(0,0,0,0.06);">
          <tr>
            <td style="background:#0E1A2B;padding:24px 32px;">
              <div style="font-family:Georgia,'Times New Roman',serif;font-size:22px;font-weight:700;color:#FFFFFF;">Optitude<span style="color:#C9A84C;">AI</span></div>
              <div style="font-size:13px;color:#C9A84C;margin-top:4px;letter-spacing:0.08em;text-transform:uppercase;">New download — Periodic Table of AI Elements</div>
            </td>
          </tr>
          <tr>
            <td style="padding:24px 32px 8px 32px;">
              <h1 style="margin:0 0 12px 0;font-family:Georgia,'Times New Roman',serif;font-size:24px;color:#1A1A1A;">New lead from the Periodic Table landing page</h1>
              <p style="margin:0 0 20px 0;font-size:14px;color:#5C5C5C;line-height:1.6;">A visitor just downloaded the Periodic Table of AI Elements. Their details are below. Reply directly to this email to follow up — the reply-to address is set to their submitted email.</p>
            </td>
          </tr>
          <tr>
            <td style="padding:0 32px 24px 32px;">
              <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;border:1px solid #E5E0D2;border-radius:4px;overflow:hidden;">
                ${rowsHtml}
              </table>
            </td>
          </tr>
          <tr>
            <td style="padding:16px 32px 24px 32px;border-top:1px solid #E5E0D2;background:#FBF9F3;">
              <p style="margin:0;font-size:12px;color:#6B6B6B;line-height:1.5;">This submission was sent automatically from <a href="https://periodic-table.optitudeai.com/" style="color:#8A6E2F;">periodic-table.optitudeai.com</a>. If the visitor ticked the AI Opportunity Review opt-in, you should reach out within 1 business day.</p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>`;
}

function buildEmailText(data) {
  return `New download — Periodic Table of AI Elements

First name: ${data.firstName}
Last name: ${data.lastName || "—"}
Email: ${data.email}
Company: ${data.company}
Job title: ${data.jobTitle || "—"}
Company size: ${data.companySize || "—"}
Main AI challenge: ${data.aiChallengeText || data.aiChallenge}
AI Opportunity Review requested: ${data.opportunityReview === "yes" ? "Yes — please contact them" : "No"}
Marketing consent: ${data.marketingConsent === "yes" ? "Yes" : "No"}
Consent timestamp: ${new Date().toISOString()}
Source: periodic-table-landing-page
Page URL: https://periodic-table.optitudeai.com/

—
Sent automatically from https://periodic-table.optitudeai.com/`;
}

module.exports = async (req, res) => {
  // Only accept POST
  if (req.method !== "POST") {
    res.statusCode = 405;
    res.setHeader("Allow", "POST");
    return res.end("Method Not Allowed");
  }

  // Parse form-encoded body (Vercel parses urlencoded automatically when Content-Type is set)
  let body = req.body || {};
  if (typeof body === "string") {
    // Fallback: parse manually if needed
    try { body = JSON.parse(body); } catch (e) {
      try { body = Object.fromEntries(new URLSearchParams(body)); } catch (e2) { body = {}; }
    }
  }

  // Honeypot check — if _honey field has a value, it's a bot. Pretend success.
  if (body._honey && body._honey.trim()) {
    res.statusCode = 303;
    res.setHeader("Location", THANK_YOU_URL);
    return res.end();
  }

  // Validate required fields
  const required = ["firstName", "email", "company", "aiChallenge"];
  for (const field of required) {
    if (!body[field] || !String(body[field]).trim()) {
      console.error("Validation failed: missing", field);
      res.statusCode = 303;
      res.setHeader("Location", ERROR_URL + "&missing=" + field);
      return res.end();
    }
  }

  // Validate email format
  const emailRe = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailRe.test(String(body.email).trim())) {
    console.error("Validation failed: invalid email", body.email);
    res.statusCode = 303;
    res.setHeader("Location", ERROR_URL + "&invalid=email");
    return res.end();
  }

  // Build clean data object
  const data = {};
  for (const key of Object.keys(body)) {
    if (ALLOWED_FIELDS.has(key)) {
      data[key] = String(body[key]).trim();
    }
  }

  // Convert aiChallenge dropdown value to human-readable text
  const challengeMap = {
    "no-idea-where-to-start": "I do not know where to start",
    "using-ai-without-structure": "We are using AI, but without structure",
    "improving-productivity": "Improving productivity",
    "automating-admin": "Automating admin or operational tasks",
    "recruitment-hr-people": "Recruitment, HR, or people processes",
    "sales-marketing-bd": "Sales, marketing, or business development",
    "ai-policy-risk-governance": "AI policy, risk, or governance",
    "training-managers-teams": "Training managers or teams",
    "other": "Other"
  };
  data.aiChallengeText = challengeMap[data.aiChallenge] || data.aiChallenge;

  // Send email via Resend
  const apiKey = process.env.RESEND_API_KEY;
  if (!apiKey) {
    console.error("RESEND_API_KEY environment variable is not set");
    res.statusCode = 303;
    res.setHeader("Location", ERROR_URL + "&err=config");
    return res.end();
  }

  try {
    const emailResponse = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${apiKey}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        from: FROM_EMAIL,
        to: [RECIPIENT_EMAIL],
        reply_to: data.email, // Simon can hit Reply to email the lead directly
        subject: `New download — Periodic Table of AI Elements (${data.firstName}, ${data.company})`,
        html: buildEmailHtml(data),
        text: buildEmailText(data)
      })
    });

    if (!emailResponse.ok) {
      const errText = await emailResponse.text();
      console.error("Resend API error:", emailResponse.status, errText);
      res.statusCode = 303;
      res.setHeader("Location", ERROR_URL + "&err=send");
      return res.end();
    }

    console.log("Email sent successfully to", RECIPIENT_EMAIL, "for lead:", data.email);

    // Redirect to thank-you page
    res.statusCode = 303;
    res.setHeader("Location", THANK_YOU_URL);
    return res.end();

  } catch (err) {
    console.error("Email send failed:", err.message);
    res.statusCode = 303;
    res.setHeader("Location", ERROR_URL + "&err=exception");
    return res.end();
  }
};
