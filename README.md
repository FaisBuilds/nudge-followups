# Nudge

**Auto follow-ups for Gmail + Google Sheets.** Written in Google Apps Script.

Nudge sends follow-ups as replies inside your original email thread. It stops on its own if the person replies or the email bounces, and it spaces sends with random delays so it never looks like a bulk blast.


## Why I built it

I do cold outreach and kept forgetting follow-ups. Most mail-merge tools either send a brand new email (which looks spammy) or keep nudging people who already replied. Nudge only sends when it's safe to, and it sends inside the same thread.

## How it works

1. Log each email you send in a Google Sheet: a date and an email address.
2. Nudge fills in a follow-up date for new rows (default: 4 days later).
3. On the follow-up date it opens your sent thread in Gmail:
   - Someone else replied: marks the row `REPLIED`, sends nothing.
   - Delivery failure: marks the row `BOUNCED`, sends nothing.
   - No reply: sends a follow-up inside the same thread and marks the row `DONE`.

![Sheet example](docs/sheet.png)

## Features

- True in-thread replies (same subject, same conversation), not new emails
- Reply and bounce detection
- Random 3 to 5 minute gap between sends
- Respects Gmail's daily send quota
- Lock to prevent double sends
- Flexible date parsing (`12 Oct 2026`, `12/10/2026`, `2026-10-12`, and more)
- Finds your columns by name, so you don't need to rebuild your sheet
- Draft mode for safe testing

## Setup

1. Open your Google Sheet and go to **Extensions > Apps Script**.
2. Paste the contents of `Code.gs` into the editor and save.
3. In the left menu, click **Services (+)**, choose **Gmail API**, and click **Add**.
4. At the top of the script, set `TIMEZONE` to yours (for example `'America/New_York'`) and `SHEET_NAME` to your tab name.
5. Leave `MODE = 'DRAFT'` for now. Select the `setup` function and click **Run**. Approve the permissions when asked.
6. Add a test row with your own second email address and a date of today, then check Gmail drafts after the follow-up date arrives.
7. When you're happy, change `MODE` to `'SEND'`.

A **Follow-ups** menu also appears in your sheet with **Run now**, **Status**, and **Start automation**.

## Your sheet

Nudge finds its columns by name. Order doesn't matter, and other columns are ignored.

| What it needs | Names it recognizes |
|---|---|
| Date you sent the email | `Date`, `Date Sent`, `Sent Date`, `Contacted`... |
| Email address | `Email`, `Email Address`, `Website/Email`... |
| Follow-up date | `NEXT_FOLLOW`, `Follow up`, `Follow-up Date`... |

- If there is no follow-up column, Nudge creates `NEXT_FOLLOW` for you.
- The header row can be anywhere in the first 10 rows.
- Cells that don't contain an email address (like "no email found") are skipped.
- If your headers are unusual, set the letters manually at the top of the script:

```js
const COLUMNS = { DATE: 'A', EMAIL: 'C', NEXT_FOLLOW: 'F' };
```

## Settings

| Setting | Default | What it does |
|---|---|---|
| `SHEET_NAME` | `'Sheet1'` | The tab to read. Falls back to the first tab |
| `TIMEZONE` | `'Asia/Karachi'` | Change this to yours |
| `FOLLOW_UP_DAYS` | `4` | Days between your email and the follow-up |
| `MODE` | `'DRAFT'` | `'DRAFT'` creates Gmail drafts, `'SEND'` sends for real |
| `DATE_ORDER` | `'DMY'` | How to read dates like 12/10/2026 (`'DMY'` or `'MDY'`) |
| `GAP_MINUTES` | `[3, 5]` | Random wait between sends |

The follow-up messages are in the `FOLLOWUPS` list. Edit them to sound like you.

## Limitations (honest ones)

- Apps Script and Gmail have daily quotas. Free Gmail accounts are limited, and Google changes the numbers from time to time, so check the current Apps Script quotas page.
- It finds your original email by searching your Sent folder for the recipient's address, then picking the thread closest to the date in your sheet. If you emailed the same person several times around the same date, double-check that row.
- It sends one follow-up per lead. There is no multi-step sequence yet.
- Replies are detected by looking for any message in the thread from someone other than you. Auto-replies and out-of-office messages count as replies.
- Plain-text messages only. No tracking pixels, no attachments.

## Responsible use

Only email addresses you found publicly or people who expect to hear from you. Follow the laws that apply to you, such as CAN-SPAM in the US: include a real sender identity and a way to opt out, and honor every "no" immediately. Start with `DRAFT` mode and low volume. You are responsible for what you send.

## Privacy

Nudge runs entirely inside your own Google account. Nothing is sent to any outside server, and no data is collected.

## License

MIT. See `LICENSE`.

Built by [Faisal](https://fais-tech.vercel.app).
