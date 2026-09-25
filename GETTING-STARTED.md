# Getting Started: Engagement Letters with Claude (Steven & Josephine)

This repo is set up so that either of you can open a new Claude session and pick up the engagement
letter work where it left off. Claude reads `CLAUDE.md` automatically at the start of every session.
That file has all the decisions made so far: phone numbers, rate-guarantee date, record retention,
email-confirmable amendments, how to name trusts and entities, and so on. You don't need to explain
any of that again.

## One-time setup for each person

1. **Claude account.** Each person needs their own login on the Onyx Claude plan (claude.ai).
2. **GitHub access to this repo.** The repo is `onyxcfo/engagement-letter-edits`. Right now the only
   account with access is the `Onyxcfo` GitHub account. For Josephine, choose one:
   - **Recommended:** Josephine creates her own GitHub account. Steven adds her on GitHub under
     **Settings → Collaborators → Add people** with *Write* access.
   - Or she signs in with the shared `Onyxcfo` GitHub account.
3. **Connect GitHub to Claude.** Each person goes to https://claude.ai/connect-github once and
   approves the Claude GitHub App for this repo.

## Starting a session

1. Go to **claude.ai/code** (or use the Claude app) and start a new session.
2. Select the repository **onyxcfo/engagement-letter-edits**.
3. Tell Claude what you need. Examples:
   - *"New engagement letter. Here is the intake call transcript: …"*
     Claude follows the `new-engagement-letter` playbook: it picks the HOURLY or FIXED template,
     fills in the client, cuts the scope to what was asked for, and lists anything still open.
   - *"Here's the name and address from the bank statement: …"*
   - *"Remove the retainer." / "Take out the hour estimates." / "This letter is final, remove the
     shading."*
   - *"Korman is signed, remove it from the repo."*
   - *"Update both templates to add this clause: …"*
4. Claude commits each change and sends you the .docx to download and review in Word.

## Tips

- **Gray shading in Word** is form-field shading. It only shows on screen and never prints. Ask
  Claude to "flatten the fields" when a letter is final. You can also do it yourself in Word:
  Ctrl+A, then Ctrl+Shift+F9.
- **Handing off.** Each session starts fresh, so pass along decisions in the chat, for example
  *"from Steven: the trust is the client."* Anything that should become a permanent rule
  (like "we only guarantee fees through year-end") can go into `CLAUDE.md`. Just ask Claude to
  "add this to the standing rules."
- **Intake sheets.** `Onyx_New_Client_Intake.xlsx` (the form) and `Onyx_Needs_Assessment.xlsx`
  (the checklist; paste the transcript into its "Transcript" tab) are in the repo root.
