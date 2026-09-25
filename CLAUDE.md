# Onyx Accounting Group — Engagement Letter Workspace

This repo holds ONYX's engagement letter templates, client-specific letters in progress, and the
new-client intake workbooks. It is shared by **Steven Nikolov** (Principal) and **Josephine**.
Either of them may start a session. Treat them as equal users. When one of them relays a decision
("from Steven: …"), apply it.

Before starting client work, read `REWRITE-WALKTHROUGH.md` for the history of the template language.
For a new client letter, follow the `new-engagement-letter` skill in `.claude/skills/`.

## What's in the repo

| File | What it is |
|---|---|
| `Onyx_Engagement_Letter_HOURLY.docx` | **Template.** Hourly rate card. |
| `Onyx_Engagement_Letter_FIXED.docx` | **Template.** Fixed fee. Same content as HOURLY except for the fee section. |
| `Onyx_Engagement_Letter_<Client>_HOURLY.docx` / `_FIXED.docx` | Client letter in progress, copied from a template. |
| `Onyx_New_Client_Intake.xlsx` | Discovery-call intake form, filled in during or after the meeting. |
| `Onyx_Needs_Assessment.xlsx` | Needs checklist. Paste the call transcript into the "Transcript" tab and the "In transcript?" column flags the topics that came up. |
| `REWRITE-WALKTHROUGH.md` | Change history for the template language. Add an entry here whenever you change a template. |

## Standing rules (from past sessions; don't re-ask these)

**Templates**
- There are exactly two templates, HOURLY and FIXED. Any legal or scope edit goes into **both**,
  plus a line in `REWRITE-WALKTHROUGH.md` under "Later updates". The REVISED master is retired.
- Don't edit a client letter that has been signed. Signed or delivered letters come out of the repo
  (`git rm`; they stay in git history). Only remove one once the user says it has been sent or signed.

**Contact details** (cover page, signature block)
- Steven Nikolov, Principal. Steven@OnyxCFO.com. **Direct 480-999-5509**, mobile 480-772-5612.
  The old direct number, 480-442-3119, is wrong. Fix it wherever it appears.

**Fees**
- Rates are guaranteed only **through December 31 of the current year**. Fill the "Billing rates
  guaranteed through [Insert date]" line with Dec 31 of the letter's year.
- Current hourly rates: Staff Accountant $100, Controller $185, CFO $250.
- Retainer and time-estimate tables are **optional**. Ask before filling or removing them. Steven has
  often removed the hour estimates, and sometimes the retainer. If you remove the time estimates,
  reword the fee paragraph so it no longer refers to them.

**Amendments / flexibility**
- Scope and fee changes take effect once **confirmed in writing, including by email**. They don't
  need a signature, addendum, or re-execution. Keep all three amendment clauses consistent: the
  opening evergreen clause, "Adjusting your services as your needs change", and "Our service
  commitment".

**Record retention**
- The 7-year "Record retention." clause is the last paragraph of "Other provisions", right before
  "Employment provision". It must be in every template and new letter.

**Tailoring a letter to the actual engagement**
- Build Section 1 ("What we'll do for you") from what the client actually asked for in the intake
  call. Remove services they didn't ask for. Don't oversell.
- If ONYX won't be moving money (paying AP, running payroll, sending wires), **remove "Cash access
  services"**. Also cut "How we'll work together" down to what applies, usually: provide access;
  provide timely information and a point of contact; the Client keeps approval authority over its
  own transactions; the Client reviews and approves bank reconciliations.
- Remove add-on services (Payroll, Sage Intacct, Hosting, HR, WOTC) that don't apply. For example,
  drop Sage Intacct for a QuickBooks client.
- Set the Initial Period to the current fiscal year. Describe prior-year work as its own catch-up or
  clean-up item, and describe it accurately (e.g., "compile reports to finalize the 2024 review"
  versus "catch-up accounting for 2025").
- Use the voice that's already in the templates: sales-forward in Section 1, and factual and precise
  in Section 2 (Terms and Conditions) and in any money-authority language.

**Naming the client (who signs)**
- Name the party that owns the books and files the return. Use its **legal name** exactly as it
  appears on the bank statement or the Arizona Corporation Commission record, and the address on
  that record.
- Entity: address the letter "Attn: <Name>, <Title>" and put that title in the signature Title line.
  Trust: "The <Name> Living Trust dtd MM/DD/YYYY", signed by the Trustee.
  Individual: the person's legal name.
- The salutation can use their preferred first name ("Dear Tom:").
- If the name in the transcript conflicts with the documents, trust the documents and point out the
  conflict. Transcripts often mishear names ("Blusero" was really "Lucero").
- If ownership is unclear (e.g., the bank account is held by a trust but QBO shows an Inc.), raise it
  and ask. Don't guess.

**Word formatting**
- The templates have yellow highlights (customizable text) and gray **form-field shading**
  (placeholders). Gray shading appears only on screen and never prints.
- When a client letter is final, remove the yellow highlights and **convert the form fields to plain
  text**. This is the equivalent of Ctrl+A then Ctrl+Shift+F9 in Word. Check that the text is
  unchanged afterward.
- Before calling a letter signature-ready, check that no `[bracket]` placeholders are left.

## How to work in this repo

- Edit the .docx files directly in their XML (unzip, edit `word/document.xml`, re-zip) or with
  python-docx. Keep the existing styles, numbering, and run formatting. Load the `docx` skill if
  it's available.
- After every edit, extract the document text and re-read the changed sections before committing.
- Commit each logical change with a clear message and push. Name the client in the message
  (e.g., "Lucero: remove retainer section").
- Send the finished .docx to the user so they can download it. They usually review in Word and
  sometimes make edits themselves. If they ask for "the change to drop in", give them old → new text
  blocks they can paste.
- Write for non-programmers. Say what changed in the letter, not what changed in the XML.
