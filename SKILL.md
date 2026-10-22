---
name: documents
description: Create, read, and edit DOCX files using the user's document styles and reusable tools for comments, tracked changes, equations, cross-references, tables, and metadata.
---

# Document tools

## Formatting specification

Before creating a DOCX or changing document formatting, read [the user's general formatting specification](references/user-formatting.md). Use it as the default for new documents. A different specification or template explicitly supplied by the user takes precedence. Preserve existing formatting during focused edits unless formatting changes are requested.

The reference is the single source for formatting values. It defines element styles, not a required document outline. Apply styles to the elements present in the requested document; do not add thesis-specific content or treat unspecified details as user requirements.

Use the reusable [formatting tools](references/formatting-tools.md) for this specification. Import `scripts/user_formatting.py` for new documents, or run it with `input.docx --out output.docx` to batch-format an existing document. Reuse its style, numbering, table, equation, caption and cross-reference functions rather than rewriting their formatting logic. Assign content roles with the documented styles or an explicit style map. Run `apply_document` and `check_document` before saving the completed document. This performs structural checks without page rendering.

Supporting guides are technical references for APIs, file structures, and script usage. Except for `references/user-formatting.md`, their sample fonts, dimensions, styles, design recommendations, and render checklists are not document formatting requirements. Read only the material needed for the requested operation.

## Runtime and authoring

- When `load_workspace_dependencies` is available, resolve the bundled runtime and dependencies through it. Otherwise use the Python executable recorded in the local `install-state.json`, or an available Python environment with the required dependencies. Run helper files from a writable workspace or temporary directory.
- Use `python-docx` for ordinary authoring and edits. Use OOXML helpers for features it does not expose, including real tracked changes, comments, relationships, and fields.
- When the current Codex environment provides artifact-operation tracking, run `container_tools/mark_artifact_operation_started.mjs` once before the first authoring command, with the operation kind, output count and format. Standalone Python use does not require this runtime marker. Read-only work does not require it.
- Author inline and display mathematics as native Word equation objects (OMML), following `references/user-formatting.md`. Use `tasks/equations.md` only for relevant technical details; its image fallback does not apply to this specification.

## Tool index

Paths are relative to this skill directory. Consult a script's help and the corresponding technical reference when needed.

| Operation | Tools | Technical reference |
| --- | --- | --- |
| Ordinary DOCX creation, editing and user formatting | `python-docx`, `scripts/user_formatting.py` | `references/formatting-tools.md`, `tasks/create_edit.md` |
| Targeted OOXML edits | `scripts/docx_ooxml_patch.py` | `ooxml/` |
| Add or accept tracked changes | `scripts/add_tracked_replacements.py`, `scripts/accept_tracked_changes.py` | `ooxml/tracked_changes.md`, `tasks/clean_tracked_changes.md` |
| Add, extract, edit, or remove comments | `scripts/comments_add.py`, `scripts/comments_extract.py`, `scripts/comments_apply_patch.py`, `scripts/comments_strip.py` | `ooxml/comments.md`, `tasks/comments_manage.md` |
| Captions and references | `scripts/captions_and_crossrefs.py`, `scripts/insert_ref_fields.py`, `scripts/fields_materialize.py`, `scripts/flatten_ref_fields.py` | `tasks/captions_crossrefs.md`, `tasks/fields_update.md` |
| Linked bibliography citations | `scripts/bibliography_fields.py` | `references/formatting-tools.md` (文献引用原生域) |
| Contents and navigation | `scripts/insert_toc.py`, `scripts/internal_nav.py` | `tasks/toc_workflow.md`, `tasks/navigation_internal_links.md` |
| Forms and content controls | `scripts/content_controls.py` | `tasks/forms_content_controls.md` |
| Spreadsheet and table conversion | `scripts/xlsx_to_docx_table.py`, `scripts/docx_table_to_csv.py` | `tasks/tables_spreadsheets.md` |
| Apply a user-supplied template or requested style cleanup | `scripts/apply_template_styles.py`, `scripts/style_lint.py`, `scripts/style_normalize.py` | Read script help; select changes according to the user's specification |
| Document structure inspection | `scripts/fields_report.py`, `scripts/heading_audit.py`, `scripts/section_audit.py`, `scripts/images_audit.py`, `scripts/footnotes_report.py` | Relevant script help |
| Metadata cleanup and redaction | `scripts/privacy_scrub.py`, `scripts/redact_docx.py` | `tasks/privacy_scrub_metadata.md`, `tasks/redaction_anonymization.md` |
| Editing protection | `scripts/set_protection.py` | `tasks/protection_restrict_editing.md` |
| Watermarks | `scripts/watermark_add.py`, `scripts/watermark_audit_remove.py` | `tasks/watermarks_background.md` |
| Merge documents | OOXML helpers | `tasks/multi_doc_merge.md` |
| Footnotes and endnotes | `scripts/insert_note.py` | `tasks/footnotes_endnotes.md` |
| Accessibility operations requested by the user | `scripts/a11y_audit.py` | `tasks/accessibility_a11y.md` |
| Rendering or visual comparison explicitly requested by the user | `render_docx.py`, `scripts/render_and_diff.py` | `tasks/verify_render.md`, `tasks/compare_diff.md` |

## Verification and delivery

Default to structural inspection appropriate to the actual edit: document text, XML validity, relationships, comment anchors, revision markup, or field values. Render pages, generate page images, and inspect them only when the user explicitly requests visual checking. Supporting references do not make rendering a prerequisite for delivery.

When rendering is requested, resolve the bundled LibreOffice path through `load_workspace_dependencies`; use `troubleshooting/libreoffice_headless.md` if needed. Comments require structural inspection because headless rendering may omit them.

Deliver the requested final files. Include intermediate artifacts only when requested.

## Google Docs

For a new native Google Docs deliverable, create the local DOCX and use the Google Drive plugin's native document import capability when available. Follow an alternate construction path if the user requests one. Do not apply an automatic title sanitizer or other formatting transformation unless requested.

## Installation and updates

When the user requests installation, replacement or updating of this skill, follow [README.md](README.md) and run `scripts/install_skill.py`. Its `--update` mode downloads the latest repository version, backs up local copies and disables bundled duplicates. Reading the skill or authoring documents does not trigger installation.
