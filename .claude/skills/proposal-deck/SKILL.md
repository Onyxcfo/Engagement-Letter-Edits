---
name: proposal-deck
description: Build the client proposal deck that accompanies an ONYX engagement letter, in the house layout first used for AJAC / George Brazil (14 slides: intro letter, contents, who we are, team table, understanding, services, how we work together, timeline, 30/60/90 plan, deliverables table, fees, why ONYX and next steps, thank-you). Use when Steven or Josephine asks for "the proposal" or "the presentation" to go with a letter.
---

# Proposal deck

The deck mirrors `Onyx_AJAC_Proposal.pptx` slide for slide. Nothing moves; only the text in each
slot changes (plus the client logo, the timeline bars and the footer). Follow `CLAUDE.md` rules
throughout (no em-dashes, CFO rate per the letter, no tax filing, ONYX does not move money).

## Steps

1. **Read the sources.** The engagement letter for the client (the deck must match its Section 1
   and "How we'll work together"), the intake-call recap or transcript, anything the client sent,
   and public facts about the client (site, ACC record). Never claim ONYX industry experience the
   firm does not have.
2. **Write the copy as a content JSON.** `slots.md` lists every slot id with the AJAC text and its
   length; the new text must keep the same number of lines and stay at or under the AJAC length
   (the AJAC deck already fills its boxes). `example_content_AZHandCenter.json` is a complete
   worked example. Conventions:
   - `**Lead-in.** text` marks a bold lead-in where slots.md says so.
   - `tl.rows`: ten lines `label | start | end | navy-or-red`, in months from the left of the first
     month column (0.0 to 6.0). Red = a fixed external date, drawn as a short bar.
   - `tl.checkpoints`: five lines `label | position`; `tl.months`: six month labels.
   - Open items the client must answer go in [square brackets] (e.g. `Mark Williamson, [Title]`).
3. **Build.**
   `python3 .claude/skills/proposal-deck/build_proposal.py content.json Onyx_<Client>_Proposal.pptx --client "<Client name>" [--logo logo.png] [--wordmark "LINE|LINE"] [--wordmark-plate "LINE|LINE|LINE"]`
   Without `--logo`, the two logo spots get a navy text wordmark from `--wordmark`. The footer,
   document title and cover note are set from `--client`.
4. **Check it.** Convert to PDF (`soffice --headless --convert-to pdf`), render the pages
   (`pdftoppm -jpeg -r 110`), and look at every slide next to the AJAC original. The usual
   problems: a card whose last bullet crosses the card border (shorten the bullet), a heading that
   wraps, a table row that grew, leftover AJAC words (George Brazil, Marc, Ian, Bell Bank, Sage
   Intacct, ServiceTitan, QStock, 401(k), S corporation, $275, $3,000). Run
   `python3 <pptx skill>/scripts/office/validate.py out.pptx --original Onyx_AJAC_Proposal.pptx`.
5. **Deliver.** Commit (name the client), push, send the .pptx, and list the open items.

## Notes
- `python-pptx` must be installed (`pip install python-docx python-pptx`).
- The AJAC deck's slide 13 contact line was too long for its column; the builder shortens it to
  `direct 480-999-5509`. Keep Steven's mobile on the cover.
- A `.jpg` or `.png` logo at the client's site can be passed with `--logo`; the builder fits it
  into the original logo box and centers it on the closing plate.
