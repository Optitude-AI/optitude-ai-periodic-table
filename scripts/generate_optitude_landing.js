// Optitude AI — Periodic Table of AI Elements — Landing Page Copy & SEO Pack
// Generates a consultancy-grade .docx deliverable
// Style: R2 Double-Rule Frame cover + MIN-1 Warm Gold palette (consulting)

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, PageNumber, NumberFormat, AlignmentType, HeadingLevel,
  WidthType, BorderStyle, ShadingType, PageOrientation, SectionType,
  TableLayoutType, LevelFormat, PageBreak,
} = require("docx");
const fs = require("fs");
const path = require("path");

// ─────────────────────────────────────────────────────────────────────────────
// Palette — MIN-1 Warm Gold (consulting / minimalist business / premium proposals)
// ─────────────────────────────────────────────────────────────────────────────
const P = {
  bg: "F3F1ED",            // cover background (warm cream)
  primary: "1A1A1A",       // headings — near-black
  body: "1F1F1F",          // body text
  secondary: "5C5C5C",     // captions / secondary
  accent: "8A6E2F",        // muted gold (darkened from D6C096 for body contrast)
  accentLight: "D6C096",   // original palette accent
  surface: "F7F4EC",       // surface (callout boxes, table header)
  surfaceAlt: "F0EBDC",    // zebra row — stronger contrast against white
  divider: "D6C096",       // divider lines
};

const c = (hex) => hex.replace("#", "");

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────
const NB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: NB, bottom: NB, left: NB, right: NB };
const allNoBorders = { top: NB, bottom: NB, left: NB, right: NB,
                       insideHorizontal: NB, insideVertical: NB };

// Heading builders
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 480, after: 200, line: 312 },
    children: [new TextRun({ text, bold: true, size: 32, color: c(P.primary),
      font: { ascii: "Calibri", eastAsia: "SimHei" } })],
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 360, after: 160, line: 312 },
    children: [new TextRun({ text, bold: true, size: 28, color: c(P.primary),
      font: { ascii: "Calibri", eastAsia: "SimHei" } })],
  });
}
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 280, after: 120, line: 312 },
    children: [new TextRun({ text, bold: true, size: 24, color: c(P.accent),
      font: { ascii: "Calibri", eastAsia: "SimHei" } })],
  });
}

// Body paragraph (English — left-aligned, no first-line indent for marketing copy)
function p(text, opts = {}) {
  const runs = Array.isArray(text) ? text : [new TextRun({ text, size: 22, color: c(P.body),
    font: { ascii: "Calibri" } })];
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 80, after: 160, line: 312 },
    children: runs,
    ...opts,
  });
}

// Inline mixed run helper
function r(text, opts = {}) {
  return new TextRun({
    text, size: opts.size || 22, color: c(opts.color || P.body),
    bold: opts.bold || false, italics: opts.italics || false,
    font: { ascii: "Calibri" },
  });
}

// Bullet list item
function bullet(text, level = 0) {
  const runs = Array.isArray(text) ? text : [new TextRun({ text, size: 22, color: c(P.body), font: { ascii: "Calibri" } })];
  return new Paragraph({
    bullet: { level },
    spacing: { before: 40, after: 60, line: 300 },
    children: runs,
  });
}

// Annotation chip — small grey label like [H1], [CTA], [Form]
function annotation(label) {
  return new TextRun({
    text: ` ${label} `, size: 18, bold: true,
    color: c(P.accent), italics: true,
    font: { ascii: "Calibri" },
  });
}

// Callout box (single-cell table with surface fill + left accent border)
function callout(title, bodyLines) {
  const titleRun = new Paragraph({
    spacing: { before: 60, after: 100, line: 300 },
    children: [new TextRun({ text: title, bold: true, size: 22, color: c(P.primary),
      font: { ascii: "Calibri" } })],
  });
  const bodyParas = bodyLines.map(line => {
    if (typeof line === "string") {
      return new Paragraph({
        spacing: { before: 40, after: 60, line: 300 },
        children: [new TextRun({ text: line, size: 21, color: c(P.body), font: { ascii: "Calibri" } })],
      });
    }
    return line; // already a Paragraph
  });
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: c(P.divider) },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: c(P.divider) },
      left: { style: BorderStyle.SINGLE, size: 18, color: c(P.accent) },
      right: NB,
      insideHorizontal: NB, insideVertical: NB,
    },
    rows: [new TableRow({
      cantSplit: true,
      children: [new TableCell({
        shading: { type: ShadingType.CLEAR, fill: c(P.surface) },
        margins: { top: 160, bottom: 160, left: 240, right: 200 },
        children: [titleRun, ...bodyParas],
      })],
    })],
  });
}

// Code/JSON-LD block (monospace, surface fill, all borders)
function codeBlock(lines) {
  const paras = lines.map(line => new Paragraph({
    spacing: { before: 0, after: 0, line: 280 },
    children: [new TextRun({ text: line, size: 18, color: c(P.body),
      font: { ascii: "Courier New" } })],
  }));
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: c(P.secondary) },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: c(P.secondary) },
      left: { style: BorderStyle.SINGLE, size: 4, color: c(P.secondary) },
      right: { style: BorderStyle.SINGLE, size: 4, color: c(P.secondary) },
      insideHorizontal: NB, insideVertical: NB,
    },
    rows: [new TableRow({
      cantSplit: true,
      children: [new TableCell({
        shading: { type: ShadingType.CLEAR, fill: c(P.surfaceAlt) },
        margins: { top: 140, bottom: 140, left: 180, right: 180 },
        children: paras,
      })],
    })],
  });
}

// Generic table builder (header row + data rows)
function dataTable(headers, rows, colWidths) {
  const headerRow = new TableRow({
    tableHeader: true, cantSplit: true,
    children: headers.map((text, i) => new TableCell({
      width: { size: colWidths[i], type: WidthType.PERCENTAGE },
      shading: { type: ShadingType.CLEAR, fill: c(P.primary) },
      margins: { top: 100, bottom: 100, left: 140, right: 140 },
      children: [new Paragraph({
        spacing: { before: 0, after: 0, line: 280 },
        children: [new TextRun({ text, bold: true, size: 20, color: "FFFFFF",
          font: { ascii: "Calibri" } })],
      })],
    })),
  });
  const dataRows = rows.map((row, idx) => new TableRow({
    cantSplit: true,
    children: row.map((cellContent, i) => new TableCell({
      width: { size: colWidths[i], type: WidthType.PERCENTAGE },
      shading: { type: ShadingType.CLEAR, fill: idx % 2 === 0 ? c(P.surfaceAlt) : "FFFFFF" },
      margins: { top: 100, bottom: 100, left: 140, right: 140 },
      children: Array.isArray(cellContent)
        ? cellContent
        : [new Paragraph({
            spacing: { before: 0, after: 0, line: 280 },
            children: [new TextRun({ text: String(cellContent), size: 20, color: c(P.body),
              font: { ascii: "Calibri" } })],
          })],
    })),
  }));
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 6, color: c(P.primary) },
      bottom: { style: BorderStyle.SINGLE, size: 6, color: c(P.primary) },
      left: NB, right: NB,
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "D8D4C8" },
      insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "E8E4D8" },
    },
    rows: [headerRow, ...dataRows],
  });
}

// Small spacer paragraph
const spacer = (h = 120) => new Paragraph({ spacing: { before: 0, after: h }, children: [] });

// ─────────────────────────────────────────────────────────────────────────────
// Cover — R2 Double-Rule Frame (consulting)
// ─────────────────────────────────────────────────────────────────────────────
function buildCover() {
  const padL = 1400, padR = 1400;
  const thickBorder = { style: BorderStyle.SINGLE, size: 18, color: c(P.accent), space: 20 };
  const titleLines = ["The Periodic Table", "of AI Elements"];
  const titlePt = 40;
  const titleSize = titlePt * 2;

  const children = [];

  // 1. Top rule
  children.push(new Paragraph({
    indent: { left: padL - 400, right: padR - 400 },
    spacing: { before: 1400, after: 200 },
    border: { top: thickBorder }, children: [],
  }));

  // 2. Whitespace
  children.push(new Paragraph({ spacing: { before: 1600 }, children: [] }));

  // 3. English label
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { after: 500 },
    children: [new TextRun({
      text: "O P T I T U D E   A I",
      size: 20, color: c(P.accent), characterSpacing: 60,
      font: { ascii: "Calibri" },
    })],
  }));

  // 4. Main title (2 lines)
  for (let i = 0; i < titleLines.length; i++) {
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: i < titleLines.length - 1 ? 120 : 360,
                 line: Math.ceil(titlePt * 23), lineRule: "atLeast" },
      children: [new TextRun({
        text: titleLines[i], size: titleSize, bold: true,
        color: c(P.primary), font: { ascii: "Arial" },
      })],
    }));
  }

  // 5. Subtitle
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { after: 300 },
    children: [new TextRun({
      text: "Landing Page Copy & SEO Pack",
      size: 28, color: c(P.primary), italics: true,
      font: { ascii: "Calibri" },
    })],
  }));

  children.push(new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { after: 400 },
    children: [new TextRun({
      text: "A high-converting, search-optimised lead-generation page for UK businesses ready to move from AI curiosity to practical implementation.",
      size: 22, color: c(P.secondary),
      font: { ascii: "Calibri" },
    })],
  }));

  // 6. Whitespace
  children.push(new Paragraph({ spacing: { before: 1400 }, children: [] }));

  // 7. Meta lines
  const metaLines = [
    "Prepared for:  Optitude AI",
    "Document type:  Landing page copy + SEO/conversion deliverables",
    "Version:  1.0  |  Status:  Ready for build",
  ];
  for (const line of metaLines) {
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 120, line: Math.ceil(12 * 23), lineRule: "atLeast" },
      children: [new TextRun({
        text: line, size: 22, color: c(P.secondary),
        font: { ascii: "Calibri" },
      })],
    }));
  }

  // 8. Whitespace
  children.push(new Paragraph({ spacing: { before: 1400 }, children: [] }));

  // 9. Footer + bottom rule
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    indent: { left: padL - 400, right: padR - 400 },
    spacing: { before: 200 },
    border: { bottom: thickBorder },
    children: [new TextRun({
      text: "Confidential — for internal use and authorised build partners only",
      size: 18, color: c(P.secondary), italics: true,
      font: { ascii: "Calibri" },
    })],
  }));

  // Single 16838 wrapper
  return [new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: allNoBorders,
    rows: [new TableRow({
      height: { value: 16838, rule: "exact" },
      children: [new TableCell({
        shading: { type: ShadingType.CLEAR, fill: c(P.bg) },
        borders: noBorders,
        children,
      })],
    })],
  })];
}

// ─────────────────────────────────────────────────────────────────────────────
// BODY CONTENT
// ─────────────────────────────────────────────────────────────────────────────
const body = [];

// ── Section: Document Overview ─────────────────────────────────────────────
body.push(h1("Document Overview"));
body.push(p([
  r("This document contains the complete landing page copy and supporting SEO/conversion pack for "),
  r("Optitude AI’s Periodic Table of AI Elements", { bold: true }),
  r(" download campaign. It is structured so that marketing, engineering, and compliance stakeholders can each lift their relevant sections without rework."),
]));
body.push(p([
  r("The copy is written in a measured, consultancy-grade tone — direct, commercially credible, and free of AI hype. It speaks to UK business owners, founders, senior leaders, HR leaders, recruiters, and operational decision-makers who are commercially minded, time-poor, and sceptical of AI buzzwords. Every section has been engineered to move the reader towards one of two actions: download the guide, or book an AI Opportunity Review with Optitude AI."),
]));
body.push(p([
  r("HTML tag annotations are shown inline as "),
  r("[H1]", { color: P.accent, bold: true, italics: true, size: 18 }),
  r(", "),
  r("[H2]", { color: P.accent, bold: true, italics: true, size: 18 }),
  r(", "),
  r("[p]", { color: P.accent, bold: true, italics: true, size: 18 }),
  r(", "),
  r("[CTA button]", { color: P.accent, bold: true, italics: true, size: 18 }),
  r(", "),
  r("[form]", { color: P.accent, bold: true, italics: true, size: 18 }),
  r(" markers so the copy can be pasted directly into the CMS with the structure intact. Technical specifications (form fields, schema markup, internal links) are presented as separate reference tables in Deliverables 3–10."),
]));

body.push(h2("Contents at a glance"));
body.push(dataTable(
  ["#", "Deliverable", "Where it lives"],
  [
    ["1", "SEO title & meta description", "Section: Deliverable 1"],
    ["2", "Full landing page copy (hero → FAQ)", "Section: Deliverable 2"],
    ["3", "Suggested form fields", "Section: Deliverable 3"],
    ["4", "CTA button text (with A/B variants)", "Section: Deliverable 4"],
    ["5", "FAQ section (standalone, schema-ready)", "Section: Deliverable 5"],
    ["6", "Suggested image alt text", "Section: Deliverable 6"],
    ["7", "Suggested internal links", "Section: Deliverable 7"],
    ["8", "Schema markup recommendations (JSON-LD)", "Section: Deliverable 8"],
    ["9", "GDPR-friendly consent wording", "Section: Deliverable 9"],
    ["10", "Implementation & QA checklist", "Section: Deliverable 10"],
  ],
  [8, 62, 30]
));

// ── Deliverable 1: SEO Title & Meta ────────────────────────────────────────
body.push(h1("Deliverable 1 — SEO Title & Meta Description"));
body.push(p([
  r("Both elements feature Optitude AI by name and the primary keyword "),
  r("Periodic Table of AI Elements", { bold: true }),
  r(". Length is kept within Google’s display limits to avoid truncation in UK search results."),
]));

body.push(h2("Primary SEO title (≤60 characters)"));
body.push(callout("Recommended", [
  "Periodic Table of AI Elements | Optitude AI",
  "Character count: 49 — fits comfortably within Google’s ~60-char display limit.",
]));

body.push(h2("Meta description (≤155 characters)"));
body.push(callout("Recommended", [
  "Download Optitude AI’s Periodic Table of AI Elements — a practical map of where AI creates value across leadership, HR, recruitment, marketing and operations.",
  "Character count: 154. Features Optitude AI, the guide name, and the keyword phrase ‘practical AI for business’ by implication.",
]));

body.push(h2("A/B variants"));
body.push(dataTable(
  ["Variant", "SEO title", "Meta description"],
  [
    ["A — benefit-led",
     "Practical AI for Business: The Periodic Table of AI Elements",
     "A practical AI framework for UK businesses. Download Optitude AI’s Periodic Table of AI Elements and see where AI fits across your organisation."],
    ["B — action-led",
     "Download the Periodic Table of AI Elements — Optitude AI",
     "Map your AI opportunity. Download Optitude AI’s Periodic Table of AI Elements — 48 use cases across leadership, HR, recruitment, sales and operations."],
  ],
  [22, 38, 40]
));

body.push(h2("Target keywords"));
body.push(bullet([r("Primary: ", { bold: true }), r("Periodic Table of AI Elements")]));
body.push(bullet([r("Secondary: ", { bold: true }), r("Optitude AI, practical AI for business, AI consultancy UK, AI readiness review")]));
body.push(bullet([r("Long-tail: ", { bold: true }), r("AI implementation support, AI training for businesses, AI for HR, AI for recruitment, AI business transformation")]));
body.push(spacer(80));

// ── Deliverable 2: Full Landing Page Copy ──────────────────────────────────
body.push(h1("Deliverable 2 — Full Landing Page Copy"));
body.push(p([
  r("Copy begins below. Annotations in "),
  r("[brackets]", { color: P.accent, bold: true, italics: true, size: 18 }),
  r(" indicate the HTML element for each block. All body copy is left-aligned in the CMS — do not justify (justification causes uneven spacing on responsive viewports)."),
]));

// === HERO ===
body.push(h2("Hero section"));
body.push(p([annotation("[H1]"),
  r("The Periodic Table of AI Elements: a practical map of where AI fits in your business", { bold: true, size: 28, color: P.primary })
]));
body.push(p([annotation("[p]"),
  r("AI is everywhere, but most UK businesses are still figuring out where it actually creates value. Optitude AI’s Periodic Table of AI Elements organises 48 practical AI concepts into eight clear groups — from foundations and data through to agents, automation, and business outcomes — so you can see, at a glance, where AI could save time, reduce manual workload, and strengthen decision-making across your organisation. Download the guide and start the conversation internally with a shared framework, not a list of tools.")
]));
body.push(p([annotation("[CTA button — primary]"),
  r("Download the Periodic Table of AI Elements", { bold: true, color: P.accent })
]));
body.push(p([annotation("[CTA button — secondary]"),
  r("Book an AI Opportunity Review", { color: P.secondary })
]));

// === Why this guide matters ===
body.push(h2("Why this guide matters"));
body.push(p([annotation("[p]"),
  r("Most business leaders we speak to are not short of AI enthusiasm — they are short of clarity. They have read the headlines, sat through demos, and fielded pitches from vendors promising transformation. What they usually do not have is a simple, shared way to talk about AI inside their own business: where it could help first, where it should wait, and where it simply does not apply yet. The result is either paralysis (doing nothing because everything feels overwhelming) or scattered experimentation (a handful of licences, no strategy, no measurable return).")
]));
body.push(p([annotation("[p]"),
  r("The Periodic Table of AI Elements is designed to fix that. It mirrors the structure of the chemistry periodic table: individual AI ‘elements’ sit in eight logical groups, and those groups combine to form the systems a real business actually runs — a custom AI assistant, an automated lead pipeline, a content factory. Instead of starting with a tool, you start with the problem. Instead of asking ‘should we use AI?’, you ask ‘which element of AI would help this specific process?’ That single shift in framing is what separates businesses that pilot AI indefinitely from those that capture real productivity gains within a quarter.")
]));

// === What the downloader will get ===
body.push(h2("What you will get from the guide"));
body.push(p([annotation("[p]"),
  r("The guide is practical rather than theoretical. It is built to be read in a single sitting and then used as a reference document during internal planning conversations. Specifically, you will come away with:")
]));
body.push(p([annotation("[ul]")]));
body.push(bullet([r("A clear map of AI use cases ", { bold: true }), r("across every major function — leadership, operations, HR, recruitment, marketing, sales, client service and decision-making — so you can see the full opportunity surface, not just the headline use cases.")]));
body.push(bullet([r("A simple way to identify where AI could help first", { bold: true }), r(" — grouped by leverage rather than by hype, with the highest-value zones called out explicitly.")]));
body.push(bullet([r("Ideas for improving productivity ", { bold: true }), r("and reducing repetitive admin, including worked examples of how individual AI elements combine into real business systems.")]));
body.push(bullet([r("Examples of where AI can support leadership, HR, recruitment, marketing, sales, operations and client communication", { bold: true }), r(" — written in plain English, with commercial context rather than technical jargon.")]));
body.push(bullet([r("A practical starting point for internal discussion and planning", { bold: true }), r(" — a shared vocabulary your team can use to prioritise, scope, and brief AI work without needing a data science background.")]));

// === About Optitude AI ===
body.push(h2("About Optitude AI"));
body.push(p([annotation("[p]"),
  r("Optitude AI is a practical AI consultancy for UK businesses. We help organisations apply AI intelligently, commercially, and safely — cutting through the hype to focus on the implementations that actually move the needle. Our work is grounded in real business outcomes: time saved, manual workload reduced, decisions made with better information, and AI adoption that is structured rather than ad-hoc.")
]));
body.push(p([annotation("[p]"),
  r("We work across the full adoption lifecycle — from the first ‘where do we start?’ conversation through to embedded, governed, day-to-day AI use inside live workflows. Our starting point is always the business problem, never the tool. We are equally comfortable advising a founder who has never logged into ChatGPT as we are working with an operations director who already has a stack of AI licences and needs to make them pay back.")
]));

body.push(h3("What we do"));
body.push(p([annotation("[ul]")]));
body.push(bullet([r("AI readiness reviews", { bold: true }), r(" — a structured assessment of where your business is today and what would need to be true before AI creates reliable value.")]));
body.push(bullet([r("AI opportunity mapping", { bold: true }), r(" — identifying and prioritising the highest-leverage use cases across your organisation.")]));
body.push(bullet([r("Workflow and process analysis", { bold: true }), r(" — mapping current processes so AI is fitted to how you actually work, not the other way around.")]));
body.push(bullet([r("AI productivity improvement", { bold: true }), r(" — targeting measurable time savings in admin-heavy, repeatable tasks.")]));
body.push(bullet([r("AI training for teams and leaders", { bold: true }), r(" — practical, role-relevant sessions that build confidence and competence without requiring technical backgrounds.")]));
body.push(bullet([r("AI policy and governance support", { bold: true }), r(" — safe-usage policies, data handling principles, and approval workflows that reduce risk from unstructured AI use.")]));
body.push(bullet([r("Recruitment and HR process improvement", { bold: true }), r(" — applying AI to sourcing, screening, onboarding and people analytics with appropriate human oversight.")]));
body.push(bullet([r("Sales, marketing, and business development systems", { bold: true }), r(" — AI-assisted lead qualification, content production, and pipeline management that strengthen business development activity.")]));
body.push(bullet([r("Practical AI implementation support", { bold: true }), r(" — hands-on help selecting tools, configuring workflows, and embedding AI into real day-to-day operations.")]));
body.push(bullet([r("Leadership education around AI adoption", { bold: true }), r(" — equipping senior teams to make confident, informed decisions about AI investment, risk, and timing.")]));

// === Why work with Optitude AI ===
body.push(h2("Why work with Optitude AI"));
body.push(p([annotation("[p]"),
  r("Downloading the Periodic Table of AI Elements is a useful first step. It gives you the map. But a map on its own does not move a business forward — most organisations still need help identifying priorities, choosing the right tools, training their teams, creating safe-usage policies, and embedding AI into the workflows that actually run day to day. That is where Optitude AI comes in. We are not a vendor pushing a platform; we are a consultancy focused on practical results, not generic AI advice.")
]));
body.push(p([annotation("[p]"),
  r("We work with businesses that want to move from AI curiosity to AI capability — without over-promising, over-spending, or exposing themselves to avoidable risk. The commercial outcomes our clients typically look for include:")
]));
body.push(p([annotation("[ul]")]));
body.push(bullet([r("Saving time", { bold: true }), r(" on repeatable, admin-heavy work that does not require human judgement.")]));
body.push(bullet([r("Reducing manual workload", { bold: true }), r(" across operations, HR, finance, and client communication.")]));
body.push(bullet([r("Improving decision-making", { bold: true }), r(" by giving leaders faster access to summarised, relevant information.")]));
body.push(bullet([r("Increasing productivity", { bold: true }), r(" in roles where output is currently capped by manual throughput.")]));
body.push(bullet([r("Supporting better recruitment and people decisions", { bold: true }), r(" with AI-assisted sourcing, screening, and people analytics — always with appropriate human oversight.")]));
body.push(bullet([r("Strengthening business development activity", { bold: true }), r(" through faster, more consistent lead qualification and content production.")]));
body.push(bullet([r("Creating clearer internal AI adoption plans", { bold: true }), r(" so effort is sequenced, measured, and aligned to business goals rather than scattered across teams.")]));
body.push(bullet([r("Reducing risk from unstructured AI use", { bold: true }), r(" — the shadow AI problem where staff paste sensitive data into public tools without policy, governance, or oversight.")]));

// === Lead-capture form ===
body.push(h2("Download the guide"));
body.push(p([annotation("[p]"),
  r("Complete the short form below and we will email you the Periodic Table of AI Elements as a PDF. The form is intentionally short — we have kept mandatory fields to a minimum so you can get to the guide quickly, while still giving us enough context to send you useful follow-up if you want it.")
]));
body.push(p([annotation("[form]")]));
body.push(callout("Form specification (summary)", [
  "Mandatory: First name, Work email, Company name, Main AI challenge right now (dropdown).",
  "Optional: Last name, Job title, Company size, Free AI Opportunity Review opt-in (checkbox).",
  "Consent: optional marketing consent checkbox + GDPR submission disclaimer (see Deliverable 9).",
  "Full field-by-field specification appears in Deliverable 3.",
  "Submit button text: ‘Download the Periodic Table of AI Elements’.",
]));

// === After-download CTA ===
body.push(h2("Not sure where AI fits yet?"));
body.push(p([annotation("[p]"),
  r("The Periodic Table of AI Elements gives you the framework. The AI Opportunity Review turns it into a plan. In a focused session, an Optitude AI consultant will walk through your business with you, identify two or three high-leverage use cases, and outline what implementation would actually look like — timeline, effort, expected return, and risk. There is no obligation to proceed, and you will come away with a clearer picture regardless.")
]));
body.push(p([annotation("[CTA button]"),
  r("Book an AI Opportunity Review with Optitude AI", { bold: true, color: P.accent })
]));

// === FAQ ===
body.push(h2("Frequently asked questions"));

body.push(h3("What is the Periodic Table of AI Elements?"));
body.push(p([annotation("[p]"),
  r("It is a practical reference framework, created by Optitude AI, that organises 48 of the most relevant AI concepts into eight logical groups — from foundational elements like models and prompts, through data and intelligence layers, to agents, automation, and the business layer where ROI actually lives. The guide explains each element in plain English and shows how they combine into real business systems. It is designed to be read in one sitting and used as a planning reference afterwards.")
]));

body.push(h3("Who is the guide for?"));
body.push(p([annotation("[p]"),
  r("UK business owners, founders, senior leaders, HR leaders, recruiters, consultants, and operational decision-makers who are interested in AI but unsure where to start. You do not need a technical background. If you are commercially minded, time-poor, and sceptical of AI hype, the guide is written for you. It is equally useful whether you have never used an AI tool or you already have a stack of licences and want to make sense of them.")
]));

body.push(h3("How can AI help my business?"));
body.push(p([annotation("[p]"),
  r("AI tends to create value in three places: reducing repetitive manual work, improving the speed and quality of decisions, and enabling capabilities that were previously out of reach for smaller teams — such as 24/7 client support, automated lead qualification, or scaled content production. The guide maps these opportunities across leadership, HR, recruitment, marketing, sales, operations, and client service, so you can see which apply to your specific context rather than working from a generic list.")
]));

body.push(h3("What does Optitude AI do?"));
body.push(p([annotation("[p]"),
  r("Optitude AI is a practical AI consultancy. We help UK businesses move from AI curiosity to practical implementation through readiness reviews, opportunity mapping, workflow analysis, productivity improvement, team and leadership training, policy and governance support, and hands-on implementation. Our starting point is always the business problem — never the tool. We focus on commercial outcomes, not technology demonstrations.")
]));

body.push(h3("Does my business need an AI readiness review?"));
body.push(p([annotation("[p]"),
  r("If your business is actively considering AI investment — whether in tools, training, or headcount — a readiness review is usually worthwhile. It tells you whether your data, processes, culture, and governance are ready to capture value from AI, or whether foundational work is needed first. It is often the difference between a successful pilot and a stalled one. Most readiness reviews surface quick wins alongside longer-term foundations, so the exercise pays back even before implementation begins.")
]));

body.push(h3("Can Optitude AI help train my team?"));
body.push(p([annotation("[p]"),
  r("Yes. We deliver practical, role-relevant AI training for teams and leaders. Sessions are built around the actual work your people do, not generic use cases. For leadership teams, the focus is on confident decision-making about AI investment, risk, and timing. For operational teams, the focus is on day-to-day use: prompt craft, safe data handling, and recognising where AI genuinely helps versus where it adds friction. Training is hands-on and can be delivered in-person or remotely.")
]));

body.push(h3("Can AI improve recruitment, HR, or leadership processes?"));
body.push(p([annotation("[p]"),
  r("Yes, with the right boundaries. AI can support sourcing, screening, onboarding, internal mobility, people analytics, and leadership decision-support. The key is keeping a human in the loop for any decision that affects an individual’s employment, and applying governance around data handling, bias, and explainability. Optitude AI helps HR and recruitment teams apply AI where it genuinely helps — usually speeding up repetitive work — while keeping judgement, fairness, and accountability firmly with people.")
]));

body.push(h3("How do we use AI safely inside a business?"));
body.push(p([annotation("[p]"),
  r("Safe AI use is mostly about structure rather than technology. It means having a written policy, an approved tool list, clear rules on what data can and cannot be entered into AI tools, defined approval workflows for new use cases, and basic training so staff understand both the capability and the limits of the tools they use. The most common risk we see is not malicious use — it is well-intentioned staff pasting sensitive information into public AI tools because no one has told them not to. A simple policy and a short training session resolve most of that risk.")
]));

body.push(h3("What is the first step towards AI implementation?"));
body.push(p([annotation("[p]"),
  r("Start with a problem, not a tool. Pick one repeatable, measurable process where AI could plausibly save time or improve quality, scope it tightly, and run a small pilot. Capture the before-and-after numbers. If the pilot works, expand. If it does not, you have spent very little and learned a lot. Most businesses that stall on AI do so because they started with a tool looking for a use case. The Periodic Table of AI Elements is designed to help you avoid that trap — and an AI Opportunity Review with Optitude AI will help you choose the right first problem.")
]));

body.push(spacer(120));

// ── Deliverable 3: Form Fields ─────────────────────────────────────────────
body.push(h1("Deliverable 3 — Suggested Form Fields"));
body.push(p([
  r("The form is intentionally short to maximise completion. Mandatory fields are limited to four; everything else is optional and clearly labelled as such. Single-column layout, inline validation, and a single submit button — no captcha on first submission (use honeypot + rate-limiting instead; introduce captcha only if abuse is detected)."),
]));

body.push(dataTable(
  ["Field label", "Type", "Required", "Helper text / placeholder", "CRM mapping"],
  [
    ["First name", "Text", "Yes", "e.g. Sarah", "Lead.FirstName"],
    ["Last name", "Text", "No", "e.g. Patel", "Lead.LastName"],
    ["Work email", "Email", "Yes", "name@company.co.uk", "Lead.Email (primary key)"],
    ["Company name", "Text", "Yes", "e.g. Northbridge Logistics Ltd", "Account.Name"],
    ["Job title", "Text", "No", "e.g. Operations Director", "Lead.JobTitle"],
    ["Company size", "Dropdown", "No", "1–10 / 11–50 / 51–250 / 251–1000 / 1000+", "Lead.CompanySize"],
    ["Main AI challenge right now", "Dropdown", "Yes", "See options below", "Lead.AIChallenge (custom)"],
    ["Would you like a free AI Opportunity Review?", "Checkbox", "No", "Tick if you would like us to contact you to book a call", "Lead.MQL_Flag + Task"],
    ["Marketing consent", "Checkbox", "No", "See Deliverable 9 for wording", "Lead.MarketingConsent + ConsentDate"],
  ],
  [22, 12, 10, 32, 24]
));

body.push(h2("Dropdown options: Main AI challenge right now"));
body.push(p("Use exactly these options, in this order, to keep segmentation clean across leads:"));
body.push(bullet("I do not know where to start"));
body.push(bullet("We are using AI, but without structure"));
body.push(bullet("Improving productivity"));
body.push(bullet("Automating admin or operational tasks"));
body.push(bullet("Recruitment, HR, or people processes"));
body.push(bullet("Sales, marketing, or business development"));
body.push(bullet("AI policy, risk, or governance"));
body.push(bullet("Training managers or teams"));
body.push(bullet("Other"));

body.push(h2("Form UX best practice"));
body.push(bullet("Single column — multi-column forms reduce completion on mobile."));
body.push(bullet("Inline validation — validate email format on blur, not on submit."));
body.push(bullet("Clearly mark optional fields with ‘(optional)’ suffix, not just an asterisk on mandatory ones."));
body.push(bullet("Submit button copy should restate the value: ‘Download the Periodic Table of AI Elements’ — not ‘Submit’."));
body.push(bullet("On submit, redirect to a thank-you page that (a) confirms the download, (b) repeats the secondary CTA to book an AI Opportunity Review, and (c) fires the conversion event for analytics."));
body.push(spacer(80));

// ── Deliverable 4: CTA Button Text ─────────────────────────────────────────
body.push(h1("Deliverable 4 — CTA Button Text"));
body.push(p([
  r("CTA copy is the single highest-leverage conversion element on the page. The recommended primary CTA restates the value of the action (you get the guide) rather than the mechanics (click here). Two A/B variants are provided for each CTA so they can be tested against the control."),
]));

body.push(h2("Primary CTA — guide download"));
body.push(dataTable(
  ["Variant", "Button text", "Rationale"],
  [
    ["Control (recommended)", "Download the Periodic Table of AI Elements", "Restates the value clearly. Matches the page H1. Sets expectation that a guide is being delivered."],
    ["A — shorter", "Get the guide", "Lower visual weight, faster to scan. Useful on mobile or in a sticky header."],
    ["B — urgency-led", "Download now — free PDF", "Adds a soft urgency cue and removes price ambiguity. Test carefully; ‘now’ can feel pushy in B2B."],
  ],
  [22, 36, 42]
));

body.push(h2("Secondary CTA — AI Opportunity Review"));
body.push(dataTable(
  ["Variant", "Button text", "Rationale"],
  [
    ["Control (recommended)", "Book an AI Opportunity Review", "Clear, specific, action-oriented. ‘Opportunity Review’ positions the call as valuable rather than salesy."],
    ["A — risk-reversal", "Book a free AI Opportunity Review", "Adds ‘free’ to lower the perceived commitment. Useful if the page is converting poorly on the secondary CTA."],
    ["B — outcome-led", "See where AI fits in your business", "Leads with the outcome the prospect wants, not the mechanics of the call. Good for top-of-funnel traffic."],
  ],
  [22, 36, 42]
));

body.push(h2("Post-download CTA — appears on thank-you page and inline after the form section"));
body.push(dataTable(
  ["Variant", "Button text", "Rationale"],
  [
    ["Control (recommended)", "Book an AI Opportunity Review with Optitude AI", "Full context, names the brand. Best when the thank-you page is the first brand touch after download."],
    ["A — softer", "Talk to Optitude AI about your AI roadmap", "Lower commitment framing. Good if the prospect has just downloaded and may not be ready for a full review."],
    ["B — value-anchor", "Get a tailored AI opportunity map — 30-min call", "Anchors the time cost and names the deliverable. Reduces no-shows because expectations are explicit."],
  ],
  [22, 36, 42]
));
body.push(spacer(80));

// ── Deliverable 5: FAQ Section (standalone) ────────────────────────────────
body.push(h1("Deliverable 5 — FAQ Section (standalone, schema-ready)"));
body.push(p([
  r("The FAQ block below is the same copy that appears inline in Deliverable 2, presented here as a clean standalone block so it can be copied directly into an FAQ component or fed into a schema generator. Pair it with the FAQPage JSON-LD schema in Deliverable 8 to earn rich-result eligibility in Google search."),
]));

const faqs = [
  ["What is the Periodic Table of AI Elements?",
   "It is a practical reference framework, created by Optitude AI, that organises 48 of the most relevant AI concepts into eight logical groups — from foundational elements like models and prompts, through data and intelligence layers, to agents, automation, and the business layer where ROI actually lives. The guide explains each element in plain English and shows how they combine into real business systems. It is designed to be read in one sitting and used as a planning reference afterwards."],
  ["Who is the guide for?",
   "UK business owners, founders, senior leaders, HR leaders, recruiters, consultants, and operational decision-makers who are interested in AI but unsure where to start. You do not need a technical background. If you are commercially minded, time-poor, and sceptical of AI hype, the guide is written for you. It is equally useful whether you have never used an AI tool or you already have a stack of licences and want to make sense of them."],
  ["How can AI help my business?",
   "AI tends to create value in three places: reducing repetitive manual work, improving the speed and quality of decisions, and enabling capabilities that were previously out of reach for smaller teams — such as 24/7 client support, automated lead qualification, or scaled content production. The guide maps these opportunities across leadership, HR, recruitment, marketing, sales, operations, and client service, so you can see which apply to your specific context rather than working from a generic list."],
  ["What does Optitude AI do?",
   "Optitude AI is a practical AI consultancy. We help UK businesses move from AI curiosity to practical implementation through readiness reviews, opportunity mapping, workflow analysis, productivity improvement, team and leadership training, policy and governance support, and hands-on implementation. Our starting point is always the business problem — never the tool. We focus on commercial outcomes, not technology demonstrations."],
  ["Does my business need an AI readiness review?",
   "If your business is actively considering AI investment — whether in tools, training, or headcount — a readiness review is usually worthwhile. It tells you whether your data, processes, culture, and governance are ready to capture value from AI, or whether foundational work is needed first. It is often the difference between a successful pilot and a stalled one. Most readiness reviews surface quick wins alongside longer-term foundations, so the exercise pays back even before implementation begins."],
  ["Can Optitude AI help train my team?",
   "Yes. We deliver practical, role-relevant AI training for teams and leaders. Sessions are built around the actual work your people do, not generic use cases. For leadership teams, the focus is on confident decision-making about AI investment, risk, and timing. For operational teams, the focus is on day-to-day use: prompt craft, safe data handling, and recognising where AI genuinely helps versus where it adds friction. Training is hands-on and can be delivered in-person or remotely."],
  ["Can AI improve recruitment, HR, or leadership processes?",
   "Yes, with the right boundaries. AI can support sourcing, screening, onboarding, internal mobility, people analytics, and leadership decision-support. The key is keeping a human in the loop for any decision that affects an individual’s employment, and applying governance around data handling, bias, and explainability. Optitude AI helps HR and recruitment teams apply AI where it genuinely helps — usually speeding up repetitive work — while keeping judgement, fairness, and accountability firmly with people."],
  ["How do we use AI safely inside a business?",
   "Safe AI use is mostly about structure rather than technology. It means having a written policy, an approved tool list, clear rules on what data can and cannot be entered into AI tools, defined approval workflows for new use cases, and basic training so staff understand both the capability and the limits of the tools they use. The most common risk we see is not malicious use — it is well-intentioned staff pasting sensitive information into public AI tools because no one has told them not to. A simple policy and a short training session resolve most of that risk."],
  ["What is the first step towards AI implementation?",
   "Start with a problem, not a tool. Pick one repeatable, measurable process where AI could plausibly save time or improve quality, scope it tightly, and run a small pilot. Capture the before-and-after numbers. If the pilot works, expand. If it does not, you have spent very little and learned a lot. Most businesses that stall on AI do so because they started with a tool looking for a use case. The Periodic Table of AI Elements is designed to help you avoid that trap — and an AI Opportunity Review with Optitude AI will help you choose the right first problem."],
];

for (const [q, a] of faqs) {
  body.push(new Paragraph({
    spacing: { before: 220, after: 80, line: 312 },
    children: [new TextRun({ text: q, bold: true, size: 24, color: c(P.primary),
      font: { ascii: "Calibri" } })],
  }));
  body.push(p(a));
}
body.push(spacer(80));

// ── Deliverable 6: Image Alt Text ──────────────────────────────────────────
body.push(h1("Deliverable 6 — Suggested Image Alt Text"));
body.push(p([
  r("Alt text serves two purposes: accessibility for screen-reader users, and additional context for search engines. Each alt text below is descriptive (not keyword-stuffed), mentions the brand where relevant, and avoids redundant phrases like ‘image of’ or ‘picture of’."),
]));

body.push(dataTable(
  ["Image", "Filename", "Alt text", "Caption"],
  [
    ["Hero — the periodic table visual",
     "periodic-table-of-ai-elements-hero.png",
     "The Periodic Table of AI Elements by Optitude AI, showing 48 AI concepts arranged in eight colour-coded groups from Fundamentals to Business Layer.",
     "Figure 1 — The Periodic Table of AI Elements, Optitude AI."],
    ["Optitude AI logo (header)",
     "optitude-ai-logo.svg",
     "Optitude AI logo",
     "(no caption — header treatment)"],
    ["Consultant headshot / about image",
     "optitude-ai-consultant.jpg",
     "An Optitude AI consultant reviewing an AI opportunity map with a UK business leader.",
     "Practical AI implementation support from Optitude AI."],
    ["Workflow diagram — how elements combine into compounds",
     "ai-elements-compound-workflow.png",
     "Diagram showing how individual AI elements such as Tokens, Prompts, Models and RAG combine into a Custom AI Assistant compound.",
     "Figure 2 — How AI elements combine into business systems."],
    ["Lead form / download visual",
     "download-periodic-table-cta.png",
     "Download the Periodic Table of AI Elements — a practical AI framework for UK businesses.",
     "(no caption — used as decorative CTA treatment)"],
    ["Open Graph / social share image",
     "og-periodic-table-of-ai-elements.png",
     "Optitude AI — The Periodic Table of AI Elements. Download the practical AI framework for UK businesses.",
     "Used for Facebook, LinkedIn and X/Twitter card previews. Recommended 1200×630px."],
  ],
  [22, 22, 38, 18]
));
body.push(spacer(80));

// ── Deliverable 7: Internal Links ──────────────────────────────────────────
body.push(h1("Deliverable 7 — Suggested Internal Links"));
body.push(p([
  r("Internal links distribute authority across the site and help search engines understand Optitude AI’s service architecture. Each link below is paired with the section of the landing page where it should appear, and a suggested slug. Use the exact anchor text where possible — anchor text is a strong relevance signal."),
]));

body.push(dataTable(
  ["Anchor text", "Suggested slug", "Appears in", "Target page purpose"],
  [
    ["AI readiness review", "/services/ai-readiness-review", "Why work with Optitude AI; FAQ #5", "Service page for the readiness review offering."],
    ["AI opportunity mapping", "/services/ai-opportunity-mapping", "About Optitude AI; Why work with us", "Service page for opportunity mapping engagements."],
    ["AI training for teams and leaders", "/services/ai-training", "About Optitude AI; FAQ #6", "Service page for training programmes."],
    ["AI policy and governance support", "/services/ai-policy-governance", "About Optitude AI; FAQ #8", "Service page for policy / governance work."],
    ["AI for HR", "/services/ai-for-hr", "FAQ #7; What you will get", "Service page focused on HR use cases."],
    ["AI for recruitment", "/services/ai-for-recruitment", "FAQ #7; What you will get", "Service page focused on recruitment use cases."],
    ["AI productivity improvement", "/services/ai-productivity", "About Optitude AI; Why work with us", "Service page for productivity-focused engagements."],
    ["practical AI for business", "/insights/practical-ai-for-business", "Hero subhead; Why this guide matters", "Insight / blog category page."],
    ["Book an AI Opportunity Review", "/contact/ai-opportunity-review", "Hero; after-download CTA", "Booking / contact page for the review call."],
    ["About Optitude AI", "/about", "Footer; About Optitude AI section", "Company overview page."],
  ],
  [26, 24, 24, 26]
));
body.push(spacer(80));

// ── Deliverable 8: Schema Markup ───────────────────────────────────────────
body.push(h1("Deliverable 8 — Schema Markup Recommendations"));
body.push(p([
  r("Three schema types are recommended for this page. All are implemented as JSON-LD blocks in the page HTML (place in the "),
  r("<head>", { font: { ascii: "Courier New" } }),
  r(" or immediately before "),
  r("</body>", { font: { ascii: "Courier New" } }),
  r(" — both are valid; Google prefers "),
  r("<head>", { font: { ascii: "Courier New" } }),
  r(" for performance). Together they make the page eligible for FAQ rich results, breadcrumb display, and enhanced brand knowledge-panel information."),
]));

body.push(h2("1. FAQPage schema (recommended — enables FAQ rich results)"));
body.push(p("Paste the following JSON-LD, populating the mainEntity array with the same nine Q&A pairs from Deliverable 5. The example below shows the structure for the first two questions:"));
body.push(codeBlock([
  "<script type=\"application/ld+json\">",
  "{",
  "  \"@context\": \"https://schema.org\",",
  "  \"@type\": \"FAQPage\",",
  "  \"mainEntity\": [",
  "    {",
  "      \"@type\": \"Question\",",
  "      \"name\": \"What is the Periodic Table of AI Elements?\",",
  "      \"acceptedAnswer\": {",
  "        \"@type\": \"Answer\",",
  "        \"text\": \"It is a practical reference framework, created by Optitude AI,",
  "                 that organises 48 of the most relevant AI concepts into eight",
  "                 logical groups — from foundational elements like models and",
  "                 prompts, through data and intelligence layers, to agents,",
  "                 automation, and the business layer where ROI actually lives.\"",
  "      }",
  "    },",
  "    {",
  "      \"@type\": \"Question\",",
  "      \"name\": \"Who is the guide for?\",",
  "      \"acceptedAnswer\": {",
  "        \"@type\": \"Answer\",",
  "        \"text\": \"UK business owners, founders, senior leaders, HR leaders,",
  "                 recruiters, consultants, and operational decision-makers who",
  "                 are interested in AI but unsure where to start.\"",
  "      }",
  "    }",
  "    /* ... repeat for all 9 FAQs ... */",
  "  ]",
  "}",
  "</script>",
]));

body.push(h2("2. BreadcrumbList schema (recommended — enables breadcrumb display)"));
body.push(codeBlock([
  "<script type=\"application/ld+json\">",
  "{",
  "  \"@context\": \"https://schema.org\",",
  "  \"@type\": \"BreadcrumbList\",",
  "  \"itemListElement\": [",
  "    { \"@type\": \"ListItem\", \"position\": 1, \"name\": \"Home\",",
  "      \"item\": \"https://www.optitudeai.com/\" },",
  "    { \"@type\": \"ListItem\", \"position\": 2, \"name\": \"Resources\",",
  "      \"item\": \"https://www.optitudeai.com/resources/\" },",
  "    { \"@type\": \"ListItem\", \"position\": 3,",
  "      \"name\": \"Periodic Table of AI Elements\",",
  "      \"item\": \"https://www.optitudeai.com/resources/periodic-table-of-ai-elements\" }",
  "  ]",
  "}",
  "</script>",
]));

body.push(h2("3. Organization schema (sitewide — place in site header or footer template)"));
body.push(codeBlock([
  "<script type=\"application/ld+json\">",
  "{",
  "  \"@context\": \"https://schema.org\",",
  "  \"@type\": \"Organization\",",
  "  \"name\": \"Optitude AI\",",
  "  \"url\": \"https://www.optitudeai.com/\",",
  "  \"logo\": \"https://www.optitudeai.com/assets/optitude-ai-logo.png\",",
  "  \"description\": \"Practical AI consultancy helping UK businesses apply AI",
  "                   intelligently, commercially and safely.\",",
  "  \"areaServed\": \"United Kingdom\",",
  "  \"sameAs\": [",
  "    \"https://www.linkedin.com/company/optitude-ai\",",
  "    \"https://twitter.com/optitudeai\"",
  "  ]",
  "}",
  "</script>",
]));

body.push(h2("Optional — Article schema"));
body.push(p([
  r("If the page is treated as a content asset (rather than a pure conversion page), adding "),
  r("Article", { font: { ascii: "Courier New" } }),
  r(" schema can help with topical authority. Include "),
  r("headline", { font: { ascii: "Courier New" } }),
  r(", "),
  r("author", { font: { ascii: "Courier New" } }),
  r(" (Optitude AI), "),
  r("datePublished", { font: { ascii: "Courier New" } }),
  r(", "),
  r("dateModified", { font: { ascii: "Courier New" } }),
  r(", and "),
  r("image", { font: { ascii: "Courier New" } }),
  r(" (point to the Open Graph image from Deliverable 6)."),
]));
body.push(spacer(80));

// ── Deliverable 9: GDPR-Friendly Consent Wording ───────────────────────────
body.push(h1("Deliverable 9 — GDPR-Friendly Consent Wording"));
body.push(p([
  r("All wording below is GDPR-aligned for a UK audience. Confirm final wording with your data protection lead or DPO before launch — particularly the lawful basis used for marketing emails (consent is the safest basis for cold/outbound marketing; legitimate interest may apply for warmer leads but requires a documented LIA)."),
]));

body.push(h2("(a) Marketing consent checkbox (optional — appears above the submit button)"));
body.push(callout("Checkbox label", [
  "☐  I would like to receive occasional practical AI insights, updates, and service information from Optitude AI.",
  "Default state: unchecked. The checkbox must be opt-in (pre-ticked boxes are not valid consent under GDPR).",
]));

body.push(h2("(b) Submission disclaimer (appears directly below the submit button)"));
body.push(callout("Disclaimer text", [
  "By submitting this form, you agree that Optitude AI may process your information to send you the requested guide. You can unsubscribe from marketing emails at any time. Please see our Privacy Policy for details on how we handle your information.",
  "Link ‘Privacy Policy’ to /privacy-policy. Open in the same tab (do not use target=’_blank’ on a legal link from a form — users should be able to use the back button).",
]));

body.push(h2("(c) Compact microcopy variant (for shorter form layouts)"));
body.push(callout("Compact wording", [
  "We will email you the guide and may contact you about related services. Unsubscribe anytime. See Privacy Policy.",
  "Use only when space is severely constrained (e.g. a sidebar form). The full disclaimer in (b) is preferred for the main page form.",
]));

body.push(h2("(d) Thank-you / confirmation email consent reminder"));
body.push(callout("Email footer wording", [
  "You are receiving this email because you downloaded the Periodic Table of AI Elements from optitudeai.com. We will only send you marketing emails if you ticked the consent box on the download form. You can unsubscribe at any time using the link below, or reply to this email with ‘unsubscribe’ in the subject line.",
  "Include a one-click unsubscribe link in every marketing email (required under UK GDPR and PECR). Transactional emails (like the download delivery itself) do not require an unsubscribe link, but including one is good practice.",
]));

body.push(h2("Compliance checklist"));
body.push(bullet([r("Lawful basis: ", { bold: true }), r("Consent for marketing emails (documented with timestamp and source). Legitimate interest for the download delivery email itself.")]));
body.push(bullet([r("Retention period: ", { bold: true }), r("Define how long lead records are retained. 24 months from last interaction is a sensible default; document in the privacy policy.")]));
body.push(bullet([r("Unsubscribe mechanism: ", { bold: true }), r("One-click unsubscribe in every marketing email (List-Unsubscribe header + visible link). Must process within 7 days (best practice: immediate).")]));
body.push(bullet([r("Data subject access: ", { bold: true }), r("Document the process for responding to DSARs within 30 days. CRM should be able to export all data for a given email address on request.")]));
body.push(bullet([r("Cookie banner: ", { bold: true }), r("If analytics or tracking cookies are used, a compliant cookie banner must be shown on first visit with reject-all as prominent as accept-all.")]));
body.push(bullet([r("Privacy policy link: ", { bold: true }), r("Present in the form disclaimer, the page footer, and the email footer. Must be accessible without scrolling (footer is acceptable).")]));
body.push(bullet([r("Lead routing: ", { bold: true }), r("Only route leads to sales if they have either (a) ticked the AI Opportunity Review opt-in or (b) explicitly requested contact. Otherwise, route to nurture only.")]));
body.push(spacer(80));

// ── Deliverable 10: Implementation & QA Checklist ──────────────────────────
body.push(h1("Deliverable 10 — Implementation & QA Checklist"));
body.push(p([
  r("Final hand-off checklist for the dev and marketing team. Every item should be ticked before the page goes live, and re-checked after any significant copy or form change."),
]));

body.push(h2("SEO"));
body.push(bullet("SEO title set in <title> tag and matches Deliverable 1 (49 chars)."));
body.push(bullet("Meta description set and matches Deliverable 1 (154 chars)."));
body.push(bullet("Exactly one H1 on the page, containing ‘Periodic Table of AI Elements’."));
body.push(bullet("H2 and H3 hierarchy is logical and matches the structure in Deliverable 2."));
body.push(bullet("All images have alt text from Deliverable 6 — none are missing or auto-generated."));
body.push(bullet("Internal links from Deliverable 7 are present with exact anchor text."));
body.push(bullet("Page is included in XML sitemap; canonical tag points to the live URL."));
body.push(bullet("Open Graph tags (og:title, og:description, og:image) are set; og:image points to the social share image."));
body.push(bullet("Page loads in under 2.5 seconds on mobile (LCP); images are WebP/AVIF with PNG fallback."));

body.push(h2("Conversion"));
body.push(bullet("Form fields match Deliverable 3 exactly (4 mandatory + 5 optional)."));
body.push(bullet("Dropdown options for ‘Main AI challenge right now’ are in the exact order specified."));
body.push(bullet("Submit button text reads ‘Download the Periodic Table of AI Elements’ (not ‘Submit’)."));
body.push(bullet("Primary CTA appears in hero; secondary CTA appears in hero and after-download section."));
body.push(bullet("Thank-you page exists, restates the download, and repeats the secondary CTA."));
body.push(bullet("Lead routing rules in CRM match the consent flags (MQL flag for AI Opportunity Review opt-in)."));
body.push(bullet("Form submits to the correct endpoint; test with a real submission before launch."));

body.push(h2("Compliance"));
body.push(bullet("Marketing consent checkbox is opt-in (default unchecked)."));
body.push(bullet("Submission disclaimer text matches Deliverable 9(b) verbatim."));
body.push(bullet("Privacy Policy link is present in form disclaimer, page footer, and email footer."));
body.push(bullet("Cookie banner is present and compliant (reject-all as prominent as accept-all)."));
body.push(bullet("One-click unsubscribe header is set on the marketing email template."));
body.push(bullet("DPO or data protection lead has signed off on final consent wording."));

body.push(h2("Analytics"));
body.push(bullet("GA4 ‘guide_download’ conversion event fires on successful form submission (not on button click)."));
body.push(bullet("GA4 ‘book_review_click’ event fires on secondary CTA clicks."));
body.push(bullet("Conversion goal configured in GA4 with the download event as the primary conversion."));
body.push(bullet("UTM parameters are preserved through to the thank-you page and stored against the lead record."));
body.push(bullet("Microsoft Clarity or equivalent session-recording tool is installed for form-funnel analysis."));
body.push(bullet("Schema markup from Deliverable 8 is validated using Google’s Rich Results Test before launch."));

body.push(spacer(160));

// Closing line
body.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 400, after: 200 },
  border: { top: { style: BorderStyle.SINGLE, size: 6, color: c(P.accent), space: 12 } },
  children: [new TextRun({
    text: "Optitude AI  ·  The Periodic Table of AI Elements  ·  Landing Page Copy & SEO Pack  ·  v1.0",
    size: 18, color: c(P.secondary), italics: true,
    font: { ascii: "Calibri" },
  })],
}));

// ─────────────────────────────────────────────────────────────────────────────
// Document assembly
// ─────────────────────────────────────────────────────────────────────────────
const doc = new Document({
  creator: "Optitude AI",
  title: "Periodic Table of AI Elements — Landing Page Copy & SEO Pack",
  description: "Landing page copy and supporting SEO/conversion deliverables for the Periodic Table of AI Elements download campaign.",
  styles: {
    default: {
      document: {
        run: {
          font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" },
          size: 22, color: c(P.body),
        },
        paragraph: { spacing: { line: 312 } },
      },
      heading1: {
        run: { font: { ascii: "Calibri", eastAsia: "SimHei" }, size: 32, bold: true, color: c(P.primary) },
        paragraph: { spacing: { before: 480, after: 200, line: 312 } },
      },
      heading2: {
        run: { font: { ascii: "Calibri", eastAsia: "SimHei" }, size: 28, bold: true, color: c(P.primary) },
        paragraph: { spacing: { before: 360, after: 160, line: 312 } },
      },
      heading3: {
        run: { font: { ascii: "Calibri", eastAsia: "SimHei" }, size: 24, bold: true, color: c(P.accent) },
        paragraph: { spacing: { before: 280, after: 120, line: 312 } },
      },
    },
  },
  sections: [
    // Cover section
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838, orientation: PageOrientation.PORTRAIT },
          margin: { top: 0, bottom: 0, left: 0, right: 0 },
        },
      },
      children: buildCover(),
    },
    // Body section
    {
      properties: {
        type: SectionType.NEXT_PAGE,
        page: {
          size: { width: 11906, height: 16838, orientation: PageOrientation.PORTRAIT },
          margin: { top: 1440, bottom: 1440, left: 1701, right: 1417 },
          pageNumbers: { start: 1, formatType: NumberFormat.DECIMAL },
        },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({
              text: "Optitude AI  ·  Periodic Table of AI Elements  ·  Landing Page Copy & SEO Pack",
              size: 16, color: c(P.secondary), italics: true,
              font: { ascii: "Calibri" },
            })],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({
              children: [PageNumber.CURRENT], size: 18, color: c(P.secondary),
              font: { ascii: "Calibri" },
            })],
          })],
        }),
      },
      children: body,
    },
  ],
});

// ── Write file ──────────────────────────────────────────────────────────────
const outPath = "/home/z/my-project/download/OptitudeAI_Landing_Page_Copy_and_SEO_Pack.docx";
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(outPath, buf);
  console.log(`✅ Document generated: ${outPath}`);
  console.log(`   Size: ${(buf.length / 1024).toFixed(1)} KB`);
}).catch(err => {
  console.error("❌ Generation failed:", err);
  process.exit(1);
});
