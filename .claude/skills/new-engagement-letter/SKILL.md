---
name: new-engagement-letter
description: Draft a new ONYX client engagement letter from an intake-call transcript (or notes), starting from the HOURLY or FIXED template. Use when Steven or Josephine says they're writing a new engagement letter, pastes an onboarding/intake transcript, or asks to fill in a letter for a client.
---

# New engagement letter

Follow the standing rules in `CLAUDE.md` throughout. The steps:

## 1. Read the inputs
- The intake transcript or notes the user provides, plus any bank statement, ACC record, or email
  they paste.
- Pull out the following:
  - **Client legal name, entity type, and address** (from the documents; don't rely on the
    transcript's spelling).
  - **Contact person and title**, and the name they go by.
  - **Fee model**: hourly or fixed. If it wasn't said, ask.
  - **Services actually requested**: bookkeeping cadence, controller review, CFO, reporting,
    lender or borrowing-base reporting, catch-up years, and systems (QBO, etc.).
  - **Whether ONYX will move money**: AP, payroll, wires, or card use.
  - **Start date and fiscal year**, any **prior-year catch-up**, retainer, and hour estimates.
- If `Onyx_Needs_Assessment.xlsx` is useful, remind the user they can paste the transcript into
  its "Transcript" tab.

## 2. Confirm the gaps in one short message
Put everything that's missing or conflicting into a single short list of questions: legal name,
address, retainer (or none), hour estimates (or remove them), letter date, and who signs. If the user
says to go ahead, use sensible defaults and leave `[bracket]` placeholders for anything unknown.

## 3. Build the letter
1. Copy the template to `Onyx_Engagement_Letter_<ClientSurnameOrShortName>_<HOURLY|FIXED>.docx`.
   Don't touch the templates themselves.
2. Fill in the cover, the date, the recipient block, the salutation, the defined-terms opening, the
   Initial Period, the rate-guarantee date (Dec 31 of this year), and the Response line and Title.
3. Rebuild Section 1 around the requested services. Remove Cash access, any unused tiers, and
   add-ons that don't apply. Add a catch-up bullet for prior years if there's catch-up work.
4. Cut "How we'll work together" down to what matches the actual scope.
5. Fill in or remove the retainer and time estimates, as the user decided.
6. Leave Section 2 (Terms and Conditions) alone unless the user asks. It is standard legal language.

## 4. Verify
- Extract the full text and re-read it. Check for leftover `[brackets]`, the wrong phone number
  (480-442-3119), wording left over from removed sections, a wrong year or period, and duplicate
  words.
- Check that "Record retention." is present and the amendment clauses are email-confirmable.

## 5. Deliver
- Commit and push, then send the .docx to the user. Summarize what's in the letter in plain English,
  then list the open items.
- Keep the yellow highlights and form-field shading while the user is reviewing. When they say the
  letter is final, strip the highlights and flatten the form fields (see CLAUDE.md).
- After the user says the letter has been sent or signed, `git rm` it from the repo.
