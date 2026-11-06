# northstar-ctx

Deterministic tracked-change builder and verifier for Word .docx files, driven by a JSON edit plan.

    export CTX_AUTHOR="Your Name"
    ctx-redline extract in.docx
    ctx-redline apply in.docx plan.json --out out.docx
    ctx-verify out.docx in.docx base.docx

Requires the docx helper scripts at $DOCX_SKILL (default /mnt/skills/public/docx/scripts).
