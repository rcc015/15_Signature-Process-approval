# Signature Process Approval

Google Apps Script workflow for sequential approval of Google Docs with dual-signature capture, email notifications, rejection handling, and automatic Open Review List generation.

## Overview

This project extends a Google Doc with an approval workflow designed for controlled documents such as SOPs, QA templates, and internal review records.

Main capabilities:

- Sequential document approval based on the signer order found in the approval table.
- Dual signature capture:
  - `Draw`: handwritten signature using an HTML5 canvas.
  - `Type`: typed signature rendered with the `Dancing Script` font.
- Automatic signer identity validation using the active Google account.
- Signature placement directly into the document approval matrix.
- Timestamp logging for each approval.
- Email notifications for the next approver.
- Final PDF generation once all approvals are complete.
- Rejection flow that resets the approval cycle.
- Open Review List generation from open Google Docs comments.

## Files

- [Code.gs](./Code.gs): server-side Apps Script logic.
- [Sidebar.html](./Sidebar.html): sidebar UI for signing and rejection.
- [appsscript.json](./appsscript.json): Apps Script manifest and scopes.
- [.clasp.json](./.clasp.json): local binding to the Apps Script project.
- [.claspignore](./.claspignore): excludes local-only files from `clasp push`.

## Functionalities

### 1. Sequential Approval Workflow

The workflow is driven by the approval table inside the Google Doc.

- The script scans document tables to find approvers.
- Approvers are extracted from:
  - Google Docs `Person` chips.
  - `mailto:` hyperlinks.
- The extracted order defines the approval sequence.
- Workflow state is stored with `PropertiesService`.

State keys used:

- `paso_actual`: current approval step.
- `lista_correos`: ordered approver email list.
- `document_id`: source document ID.
- additional internal keys for approver metadata and audit state.

### 2. Automatic Signer Validation

Before a signature is accepted, the script verifies:

- the active user email via `Session.getActiveUser().getEmail()`
- the expected approver at the current step

If the active user is not the expected approver, the signature is rejected.

### 3. Auto-Opening Sidebar

When the document opens:

- a custom menu named `Approvals` is added
- if the active user is the current approver, the signature sidebar opens automatically

Menu actions:

- `Open Signature Panel`
- `Initialize Workflow`
- `Diagnose Links`
- `Reset Workflow`

### 4. Signature Sidebar

The sidebar includes two modes:

#### Draw mode

- HTML5 canvas for mouse or touch input
- clear button to reset the drawing

#### Type mode

- text input for typed signature
- live preview rendered with Google Font `Dancing Script`
- hidden export canvas for image generation

#### Signature export behavior

- the visual signature is normalized into a fixed image frame
- scaling is proportional
- the content is enlarged from the center to improve readability
- the final inserted image preserves aspect ratio

### 5. Signature Placement in the Document

When a signature is confirmed:

- the signature image is inserted into the signer’s column
- placement is resolved to the signature cell in the approval matrix
- any previous script-generated signature image in that cell is removed
- the signature note is written in the `Date:` cell of the same column

Current note format:

`Signed on yyyy-MM-dd HH:mm:ss via typed signature`

or

`Signed on yyyy-MM-dd HH:mm:ss via drawn signature`

### 6. Signature Cleanup and Reset

`Reset Workflow` does two things:

- clears workflow state
- removes all script-generated approval marks from the table:
  - signature images
  - `Signed on ...` notes

It preserves:

- approver names
- `Job position`
- static document content

The script restores `Date:` in the date cell after cleanup.

### 7. Email Notifications

#### Next approver email

After each successful signature, the next approver receives an email with:

- Conceivable branded header
- document name
- current workflow step
- list of completed approvers
- action button linking to the Google Doc

Subject:

`Action required: document approval`

#### Final approval email

When the last approver signs:

- the document is saved
- a final PDF is generated
- all approvers receive a completion email

The final email contains:

- branded header
- document link
- final PDF link
- list of approvers

Subject:

`Document fully approved`

### 8. Final PDF Generation

After the last signature:

- the active document is explicitly saved with `saveAndClose()`
- the document is reopened
- the PDF is exported from the saved state

This was added to ensure the final approver’s signature is included in the generated PDF.

Output:

- PDF file name: `Document Name - Approved.pdf`
- destination folder: `Approved PDFs`
- folder location: same parent folder as the source document

### 9. Rejection Workflow

The sidebar includes a `Reject document` button.

Behavior:

- only the current approver can reject
- all script-generated signatures are cleared
- all `Signed on ...` notes are removed
- workflow returns to step `0`
- the first approver is notified by email

Optional input:

- `Rejection note`

Rejection email includes:

- document name
- rejecting approver name
- optional rejection note
- document link
- Open Review List link when available

Subject:

`Document returned for rework`

### 10. Open Review List Generation

When a document is rejected, the script tries to build an ORL from open comments.

#### Source of comments

The ORL uses Google Drive API comments through the Apps Script advanced Drive service.

It pulls:

- open comments only
- non-deleted comments
- comment content
- quoted file content when available
- author display name

#### ORL document creation

The ORL is created in the same parent folder as the source document.

Naming format:

`ORL_NameOfDocument_yyyymmdd-hhmm`

#### ORL header behavior

The script copies the source document header and then remaps key fields:

- `DOCUMENT TYPE` → `Open Review List`
- central title area → `Open Review List of "Document Name"`
- `Issue date` value → generation date in `yyyy-MM-dd`
- `Expiration date` label → `Generated by`
- corresponding value below → rejecting approver name

The footer contains:

`Generated on yyyy-MM-dd HH:mm by Full Name.`

#### ORL body

The ORL contains a table with:

- `Finding`
- `Section`
- `Status`

Column definitions:

- `Finding`: comment author + comment content
- `Section`: best-effort section match based on quoted content and nearest numbered section or heading
- `Status`: initialized as `Open`

If there are no open comments but the rejector entered a rejection note, the ORL still includes one row with:

- `Finding`: rejection note
- `Section`: `General`
- `Status`: `Open`

### 11. Link Diagnostics

`Diagnose Links` inspects the document and reports:

- number of scanned tables
- number of detected approver links
- sample approver entries

This helps debug cases where the document uses plain text instead of person chips or hyperlinks.

## Technical Notes

### Supported approver sources

The approver extraction supports:

- Google Docs person chips through `DocumentApp.ElementType.PERSON`
- traditional `mailto:` links in text nodes

### Header and footer manipulation

The ORL copies the source document header and modifies specific cells in-place. This is template-sensitive, so major template changes may require small adjustments.

### Section matching limitations

Comment-to-section mapping is heuristic. It uses:

- quoted comment content from Drive comments
- nearest heading or numbered section found in the document body

This is usually good enough for ORL generation, but it is not a guaranteed exact anchor match.

## Manifest and Scopes

The script uses these main capabilities:

- Google Docs document read/write
- sidebar UI
- mail sending
- Drive file access
- active user email access
- Drive advanced service for comments

See [appsscript.json](./appsscript.json) for the exact scopes and advanced service configuration.

## Deployment

### Push to Apps Script

```bash
clasp push -f
```

### Push to GitHub

```bash
git add .
git commit -m "Describe approval workflow"
git push origin main
```

## Typical Usage Flow

1. Open the controlled document.
2. Run `Approvals > Initialize Workflow`.
3. The current approver signs from the sidebar.
4. The next approver receives an email.
5. Repeat until all approvers sign.
6. Final PDF is created and completion email is sent.

If rejected:

1. Current approver clicks `Reject document`.
2. Existing approvals are cleared.
3. ORL is generated from open comments.
4. First approver is notified to rework the document.
5. Workflow restarts from step `0`.

## Known Constraints

- `Session.getActiveUser().getEmail()` depends on Google Workspace visibility rules and may be restricted in some environments.
- The approval matrix structure is assumed to be consistent with the current template.
- ORL header remapping is tailored to the current document header layout.
- Reading comments depends on the Drive advanced service being enabled.

## References

Official documentation used for implementation details:

- Apps Script Document service:
  https://developers.google.com/apps-script/reference/document/document
- Apps Script Person element:
  https://developers.google.com/apps-script/reference/document/person
- Google Drive comments resource:
  https://developers.google.com/workspace/drive/api/reference/rest/v3/comments
- Google Drive comments list:
  https://developers.google.com/workspace/drive/api/reference/rest/v3/comments/list
