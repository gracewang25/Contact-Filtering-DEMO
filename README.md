# Contact Filtering

A small, focused web app meant to take a messy uploaded spreadsheet of contacts and quickly figure out what's probably a duplicate vs. what should be added as new. After matching, users are able to add contacts into the CRM with bulk workflows and inline create/update actions.

Originally built around QuickBase workflows, then cleaned up for sharing and demo-safe config. No real IDs or tokens, and a built-in sample dataset so the full UI still works.

---

## Files

```
contact-filtering/
├── contact-filtering.html       # Page layout and all UI elements
├── README.md
├── css/
│   └── contact-filtering.css    # Styling and responsive layout
└── js/
    ├── contact-filtering.js     # App orchestration, state, rendering, QB request flow
    ├── contact-filtering.utils.js   # Shared utilities (normalization, dates, error helpers)
    ├── contact-filtering.data.js    # File parsing and duplicate matching logic
    └── contact-filter.demo.js   # Demo mode — sample data and QB mock (removable)
```

---

## What it does

- Parses an uploaded spreadsheet of contacts
- Matches rows against an existing contact set loaded from QuickBase
- Flags likely duplicates and new contacts
- Row-level and bulk actions: update existing, create new, add interactions
- Interaction management linked to contacts (with a modal form)
- Exports results to Excel

---

## Matching strategy

Duplicate detection is intentionally hybrid so it's both precise and forgiving:

1. Exact email match
2. Exact name + affiliation match
3. Fuzzy fallback (Fuse.js) with a configurable confidence threshold

Exact rules catch the clean cases and fuzzy catches the messy ones (typos, name variations, etc.).

---

## Demo mode

The app ships with demo mode on so you can open it in a browser and try the full UI immediately. Keep in mind some features were not safely recreated for demo (view and edit existing interactions, etc)

**To try it:**
1. Open `contact-filtering.html` in a browser
2. The demo banner will show at the top — click **Load Demo Upload** to populate sample contacts
3. Review the duplicate/new split, try editing cards, add interactions, export results

## How to run

This is a plain static front-end app. Just open `contact-filtering.html` in a browser.

---

## What I'd improve next

- Automated tests for the parsing and matching logic
- Move write operations behind a thin backend service (idempotency, retry safety)
- Make fuzzy match field weighting configurable
- Richer conflict-resolution UX for ambiguous matches

---
