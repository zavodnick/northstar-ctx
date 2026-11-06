---
name: northstar-ctx
description: "Use this skill when the user pastes or uploads a contract, agreement, DUA, SOW, CDA, NDA, MSA, CSA, amendment, or any other agreement for review in a client project chat. Triggers: 'review this', 'redline this', any uploaded .docx or .pdf that is an agreement alongside a forwarded email or request to review. Also triggers when the user references a prior round ('they came back with a revised draft', 'round 2'). Do NOT use for policy review, FMV assessments, monitoring tasks, ADP review, or general legal questions."
---

# NorthStar CTX Deliverable v3.4.0

The skill runs in the client project. The model makes the judgments: positions,
edit text, comments, the transmittal. Two commands do the mechanics:
ctx-redline (extract, accept, apply, check) and ctx-verify (the rule 11 pair).
They install from PyPI as the package northstar-ctx. No code lives in this
file and none is ever typed, repaired, or reconstructed by the model. Claude
never edits XML and never writes a full paragraph; every edit is a scoped
find/replace inside one paragraph ID.

## RULES: what the ctx skill tells every project
1. Do the redline here, in this project. Never send Dan to Claude Code.
2. The deliverable is the redlined Word file, built by ctx-redline and passed
 by ctx-verify. A docx the commands did not build is never a deliverable.
 If the commands cannot install or halt, the turn ends with the halt text
 and a plain list of the intended edits for Dan; never a hand-built docx,
 never XML edited by the model. Exceptions to the Word file only when Dan
 asks for a quick once-over or a brief list or chart of redlines.
3. Make the edit. Never write a comment saying an edit is needed, and never
 narrate an edit in a comment. Comments are only for a real question or a
 business decision the client must make.
4. Work out from the email and context whose edits and comments are whose. Do
 not ask. Our side's questions: answer them and/or redline as needed. The
 other side's edits: accept what we can accept; surgically redline what we
 cannot, as they wrote it. If Dan says at the start to leave their accepted
 edits visible, do that.
5. Never accept our own client's edits. Sponsor-side tracked changes stay
 visible.
6. Never edit inside a still-tracked insertion. Accept it, then redline it.
7. Keep the other side's comments. Reply in their thread, after the last
 existing reply. Mark resolved ones done. Never delete a chain.
8. File name: [Client] - [Counterparty] - [Agreement type] - dz - [MM-DD-YY].
 Example: Milestone - IQVIA - Clinical Services Agreement - dz - 10-20-26.
9. Every change and comment shows as Dan Zavodnick, Eastern time. Accepted
 text looks like original text, no red residue.
10. Every review is surgical. Never return an agreement that is all red.
 Redline and comment as needed to reasonably protect the client and get to
 execution, whatever the agreement's size or complexity.
11. Before presenting: confirm the file opens, the intended edits are actually
 in it, visible comments match the file's comment list, and every edit
 matches an instruction. Then present to Dan. No transmittal, no ledger
 entry until the redline is right.
12. When done, ask "hours for [file]?" Never guess hours or push billing on
 your own.

## ADDITIONAL RULES (from incident review, binding like rules 1 to 12)
13. Comment IDs are unstable. Word reassigns comment IDs when a user edits
 and saves. Never assume a comment ID persists across edit sessions. When
 checking whether a comment was preserved or removed, search by the
 comment's text content, not its numeric ID.
14. Verify the base text before prescribing an edit. Before adding a
 knowledge qualifier, a materiality qualifier, or any other modifier,
 extract and read the paragraph to confirm the modifier is not already
 there. Never prescribe an edit from memory of what the document says.
15. When a find string matches more than one paragraph, choose by context.
 Extract lists every paragraph with its ID and text. When the plan targets
 a defined term, a specific clause, or a schedule reference, use the
 paragraph ID where the term is defined or the clause appears, not the
 first match.
16. When Dan re-uploads a document after editing, re-extract and diff against
 the prior extract before proceeding. Report what changed (new tracked
 changes, changed comment IDs, new or removed paragraphs). Do not assume
 the document is the same.

## SETUP (once per session)
The commands install from PyPI. Run this once, read the two lines it prints.

```bash
pip install -q northstar-ctx==3.4.0 2>&1 | grep -v WARNING
pip show northstar-ctx 2>/dev/null | grep ^Version
export CTX_AUTHOR="Dan Zavodnick" PATH="$HOME/.local/bin:$PATH"; ctx-redline --help | head -1
```

Must print `Version: 3.4.0` and `usage: ctx-redline`. Anything else is a hard
stop: tell Dan "ctx 3.4.0 did not install" with the output, deliver the edit
list per rule 2, and stop. Never patch, retype, or substitute the commands.
Every later bash block that runs ctx-redline or ctx-verify starts with
`export CTX_AUTHOR="Dan Zavodnick" PATH="$HOME/.local/bin:$PATH"` (rule 9). The commands use the docx skill
at /mnt/skills/public/docx/scripts (DOCX_SKILL overrides).

## CORRECTION CHECK

Before producing any CTX output (redline, verification output,
comments, edit specs):

1. list_reminders filtered to tag "ctx-correction".
2. For each active correction: read the reminder body; apply the fix
   as a binding constraint on this output.
3. If client-specific corrections exist for the current client, those
   override general corrections on the same topic.
4. After output, verify each active correction was honored. If any
   was missed, fix before showing Dan.

Common correction areas: XML edit spec formatting, comment threading,
cross-author handling, redline paragraph structure, verify output
fidelity.

## WORKFLOW
1. Read the whole agreement. Note type, counterparty, who sent it, the
 client's leverage, the round, and what the agreement already handles well.
2. Client context: the client block (playbook and positions pointers,
 contacts, routing, billing rules); MCP for prior rounds and contracts.
3. Extract: `cp "/mnt/user-data/uploads/<agreement>.docx" /home/claude/in.docx`
 then `ctx-redline extract in.docx > extract.txt`. IDs are
 [P047], [T1R2C1], [TB3]. EXISTING TRACKED CHANGES lists rev id and author;
 COMMENTS lists author and thread root or reply. A STRUCTURE HALT stops the
 job: say the tool cannot address this document.
4. Round 2+ (tracked changes or comments present): sort the authors per rule
 4: counterparty, sponsor-side, Dan Zavodnick. Accept step: `ctx-redline accept
 in.docx base.docx --keep "Sponsor A,Sponsor B"`;
 counterparty changes become clean text, sponsor and Dan changes stay live
 (Dan is always kept). `--ids 12,15` accepts only chosen changes. Skip the
 step when Dan said to leave their edits visible. Extract base.docx again
 and plan against it. Round 1: base.docx is in.docx.
5. Decide per position: EDIT, SKIP, or COMMENT. Written privately, never shown.
6. Write plan.json (JSON EDIT SPEC below). Copy find and anchor text from the
 extract and keep each unique within its paragraph.
7. Apply: `ctx-redline apply base.docx plan.json --out
 "/mnt/user-data/outputs/[Client] - [Counterparty] - [Agreement type] - dz -
 [MM-DD-YY].docx"` (rule 8). A HALT writes nothing: fix the plan and rerun;
 a HALT is never a reason to edit the file another way. Read every line of
 the VERIFICATION REPORT (BEFORE, AFTER, MARKED per touched paragraph, each
 inserted paragraph, comment, reply, resolve, and any insertion accepted
 under rule 6), then VALIDATE, then CHECK. Paste the CHECK line verbatim.
 CHECK: FAIL ends the turn with the failure, not a file.
8. Completeness check, private. Positions sweep: each playbook position is
 addressed, already adequate in Section X, not worth it given leverage, or
 MISSED. Section sweep (liability, indemnification, term, termination, IP,
 confidentiality, payment, data, insurance, assignment, governing law,
 dispute resolution): any exposure left untouched is MISSED. Red flag scan:
 sole and exclusive remedy, AS IS, in perpetuity, irrevocable, one-sided
 indemnify and hold harmless, sole discretion, waive, liquidated damages,
 automatic renewal, unlimited audit, reimburse costs of audit, third-party
 beneficiary. Contradiction scan: for every paragraph the redline changed,
 check whether any other section in the document grants, limits, or assumes
 the opposite (e.g., an exclusive license in one section vs. co-ownership in
 another; a covenant not to sue that covers IP the redline just narrowed).
 Any MISSED or CONTRADICTION: add the edit, rerun apply, reread, recheck,
 and say in one line: "Completeness check added two edits: [what]."
9. Verify (rule 11): `ctx-verify out.docx in.docx base.docx`.
 REJECT must report mismatched: [] and the run must end VERIFY: PASS. Paste
 the VERIFY line verbatim. VERIFY: FAIL ends the turn with the failure, not a
 file. A validator or check that returns nonzero is read in full, never
 explained away.
10. Stale comment scan. Extract the output document. For every comment
 (original and new), read the comment text and the paragraph it is anchored
 to. If the comment asserts a fact about the paragraph ("caps liability at
 zero," "no knowledge qualifier," "this grants exclusivity") that the
 current paragraph text contradicts, flag it: "STALE: C[id] says [X] but
 the paragraph now reads [Y]." Fix or remove every stale comment before
 proceeding. This catches Round 1 comments that were correct when written
 but wrong after the redline.
11. Cross-reference audit. After any paragraph insertion, deletion, or
 renumbering edit: extract the output, list every reference to a section,
 article, or exhibit number (e.g., "Section 1.8", "Article III",
 "Exhibit B"), and verify each points to the section it names in the
 current numbering. If the redline inserted or deleted a section and
 shifted numbers, every downstream reference must be updated. List each
 wrong reference as: "XREF: [paragraph] references [old] but it is now
 [new]." Add the correcting edits, rerun apply, and reverify. Do not
 present until all cross-references are correct.
12. Substantive read (the "you sure?" gate). This is not a mechanical check.
 Read the accepted view of the entire document, as a lawyer reviewing it
 for the client. Ask:
 (a) Does each edit make sense in context, not just in isolation?
 (b) Do the edits create any internal inconsistency (a defined term changed
 in one place but used with the old meaning elsewhere; a carve-out
 referencing the wrong section; a new paragraph that conflicts with an
 existing one)?
 (c) Is every comment still accurate given what the redline changed?
 (d) Would Dan be embarrassed if the counterparty's lawyer read this?
 Present findings as a distinct block: "SUBSTANTIVE READ: [n] issues" or
 "SUBSTANTIVE READ: clean." Fix every issue before proceeding. Do not say
 "ready" until this step produces "clean."
13. Visual check: `python /mnt/skills/public/docx/scripts/office/soffice.py
 --headless --convert-to pdf out.docx` then `pdftoppm -jpeg -r 80 out.pdf
 page`. Look at the changed pages. LibreOffice does not render comments.
14. Deliver: present_files with the redline, then the transmittal body in chat.
 Never say "ready to send" without steps 10 through 12 completed and clean.
15. Records and billing through the connector: add_client_document (title
 `[CODE] EVIDENCE [counterparty] [agreement type] [version or date]`);
 never add_contract or update_contract (reviewed paper is evidence plus a
 review activity); log_activity per the logging skill; ask "hours for
 [filename]?" (rule 12); on the answer, queue_freshbooks_entry tagged
 Legal or Compliance, left drafted. Equity follows the client block.

## JSON EDIT SPEC (plan.json)
```json
{"edits": [
 {"id": "P047", "op": "replace", "find": "SHALL HV BE LIABLE", "with": "SHALL EITHER PARTY BE LIABLE"},
 {"id": "P047", "op": "insert_after", "find": "CONSEQUENTIAL DAMAGES", "text": ", EXCEPT FOR BREACHES OF SECTION 5"},
 {"id": "P050", "op": "delete", "find": " sole and exclusive"},
 {"id": "P052", "op": "insert_paragraph_after", "text": "10.11 Insurance. ..."},
 {"id": "P060", "op": "delete_paragraph"},
 {"id": "P049", "op": "comment", "anchor": "cost of any such audit", "text": "@Client, business decision: ...", "question": true},
 {"op": "reply", "parent": 12, "text": "..."},
 {"op": "resolve", "parent": 12}
], "questions": false}
```
Eight ops, no others. find and anchor match that one paragraph's rendered
text with interior whitespace normalized; leading and trailing spaces are
kept, so " sole and exclusive" removes the space and "sole and exclusive"
leaves it. Zero or multiple matches halt and print the paragraph. parent is
any comment id in the thread: reply lands after the thread's last existing
reply, never after the root; resolve marks the whole thread done. "question":
true (or "questions": true for the plan) marks a comment for the gate below.

## BUILD GATES AND CHECKS (the builder enforces these)
- Rule 6: a find range inside another author's still-tracked insertion
 accepts that insertion first, then applies the edit; the report lists it
 under "accepted before editing"; validate uses the input accepted the same.
- Cross-author refusal: a find range that runs from one author's insertion
 into another's prints both authors and stops. Split the edit at the boundary.
- Comment gate: more comments than text edits fails the build unless the
 comments are marked as questions. Replies and resolves are not counted.
- Post-build CHECK (also `ctx-redline check out.docx plan.json`):
 python-docx opens the file; LibreOffice headless converts it (NOT RUN where
 soffice is absent); commentRangeStart, commentRangeEnd, and commentReference
 counts equal the comment count in comments.xml; every plan item has its
 w:ins, w:del, comment, reply, or resolved thread. Ends CHECK: PASS or FAIL.
- ctx-verify: this turn's Dan Zavodnick revisions rejected must read
 as base.docx paragraph for paragraph; only the body and comment parts
 changed; original comments preserved verbatim; the four comment parts
 agree; new comments by Dan Zavodnick, no dashes; every reply displays last
 in its thread. Ends VERIFY: PASS or VERIFY: FAIL.

## COMMENTS
Business-decision comments are addressed to the client company
(`@Milestone`), never an individual; internal, so the transmittal tells the
client to address and delete them before sending on. Replies go through the
reply op; resolve the thread when the reply closes it. Never a dash in one.

## TRANSMITTAL
Clean: "Hi [name], Attached is our redline. Happy to discuss. Dan". Serious
issues: add "A few items to flag: [one sentence each, max 3]." Internal
@[Client] comments present (the email goes to the client): "Attached is our
redline of the [agreement]. We've left a few comments for your team marked
@[Client]; please address and delete those before sharing with
[counterparty]." Round 2+: open items first, no change inventory.

## PDF INPUT
Convert with soffice to .docx, extract, then produce a .docx where every change
is a comment (`Replace "[original]" with: "[new]"`); the transmittal says why.

## SCRIPTS
Source of record: github.com/zavodnick/northstar-ctx. Published to PyPI as
northstar-ctx; a push to main with a new version in pyproject.toml publishes
it. Changing the commands means a new version there and a new pin in SETUP,
in the same release. This skill file is deployed by uploading it in claude.ai
Settings under Skills; the in-chat skill card does not reach client projects.

## Version
3.4.0 (2026-11-06): commands install from PyPI (pip install northstar-ctx);
no code in the skill file, no fingerprints, no Drive step. Author name set by
CTX_AUTHOR. Rule 2 and steps 7 and 9 hardened after the Rice SRA round 2
incident: no hand-built docx ever, CHECK and VERIFY lines pasted verbatim,
FAIL ends the turn. Diagnostic 2026-11-06 showed client projects were still
loading 3.3.2 (inlined scripts) from the account skill library; 3.3.3 never
reached them.
3.3.3: scripts moved out of SKILL.md into scripts/ as bundled files; SETUP
locates and fingerprints the files instead of retyping them (fixes the
flattened-indentation failure every session). decompose_runs: mixed-content
runs (text + tabs + breaks) are decomposed into single-element runs before
splitting, so edits near tabs and line breaks no longer halt. people.xml:
created and populated when comments are added (fixes Word 365 repair prompt
on open). dom.toxml encoding: consistent UTF-8 on the apply path.
Workflow steps 10 to 12 added (stale comment scan, cross-reference audit,
substantive read gate). Additional rules 13 to 16 added (comment ID
instability, verify base text before edits, paragraph targeting, re-upload
diff). Incident review: SPA redline session, 12 failures traced.
3.3.2: CORRECTION CHECK gate before all output; Universal 3.2.7.
3.3.1: step 12 never writes contract rows (add_contract and update_contract
removed; decision 2026-09-28). 3.3.0: runs in the client project; the twelve
rules replace routing and workflow; minidom builder; round-2 helpers; find
refusals and gates; CHECK; verify_redline.py inlined. Full entries in
NorthStar CTX Skill Changelog. Placeholders: 0.
