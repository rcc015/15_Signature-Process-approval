const STEP_KEY = 'paso_actual';
const EMAIL_LIST_KEY = 'lista_correos';
const DOCUMENT_ID_KEY = 'document_id';
const APPROVERS_KEY = 'approval_approvers_v1';
const UPDATED_AT_KEY = 'approval_updated_at_v1';
const LAST_SIGNED_BY_KEY = 'approval_last_signed_by_v1';
const SIGNATURE_IMAGE_WIDTH = 112;
const PDF_FOLDER_NAME = 'Approved PDFs';
const SIGNED_NOTE_PREFIX = 'Signed on ';
const ORL_STATUS_OPEN = 'Open';

function onOpen(e) {
  const menu = DocumentApp.getUi().createAddonMenu();
  menu.addItem('Open Signature Panel', 'showSignatureSidebar');
  menu.addItem('Initialize Workflow', 'initializeApprovalFlow');
  menu.addItem('Diagnose Links', 'debugApprovalLinks');
  menu.addItem('Reset Workflow', 'resetApprovalFlow');
  menu.addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function buildDocsHomepageCard() {
  const section = CardService.newCardSection()
    .addWidget(
      CardService.newTextParagraph()
        .setText(
          'Use this add-on from the <b>Extensions</b> menu inside Google Docs.' +
          '<br><br>Recommended flow:' +
          '<br>1. Initialize Workflow' +
          '<br>2. Open Signature Panel' +
          '<br>3. Sign or reject as the current approver'
        )
    );

  return CardService.newCardBuilder()
    .setHeader(
      CardService.newCardHeader()
        .setTitle('Approval Signature')
        .setSubtitle('Sequential document approval for Google Docs')
    )
    .addSection(section)
    .build();
}

function showSignatureSidebar() {
  const template = HtmlService.createTemplateFromFile('Sidebar');
  template.context = getApprovalContext_();
  const html = template.evaluate()
    .setTitle('Approval Signature')
    .setWidth(360);
  DocumentApp.getUi().showSidebar(html);
}

function initializeApprovalFlow() {
  const state = buildApprovalState_();
  saveApprovalState_(state);

  const currentApprover = state.approvers[state.currentStep] || null;
  const message = currentApprover
    ? 'Workflow initialized. Current approver: ' + currentApprover.name + ' (' + currentApprover.email + ').'
    : buildNoApproverMessage_();

  DocumentApp.getUi().alert(message);
}

function resetApprovalFlow() {
  clearAllApprovalMarks_();
  const props = PropertiesService.getDocumentProperties();
  [
    STEP_KEY,
    EMAIL_LIST_KEY,
    DOCUMENT_ID_KEY,
    APPROVERS_KEY,
    UPDATED_AT_KEY,
    LAST_SIGNED_BY_KEY
  ].forEach(function(key) {
    props.deleteProperty(key);
  });
  DocumentApp.getUi().alert('Workflow state reset.');
}

function getSidebarData() {
  return getApprovalContext_();
}

function confirmSignature(payload) {
  if (!payload || !payload.imageBase64) {
    throw new Error('No signature was received.');
  }

  const state = getApprovalState_();
  if (!state.initialized) {
    throw new Error('The workflow is not initialized for this document. Run Approvals > Initialize Workflow and try again.');
  }
  const approver = state.approvers[state.currentStep];
  if (!approver) {
    throw new Error('The workflow is already complete or has not been initialized.');
  }

  const activeEmail = getActiveUserEmail_();
  if (!activeEmail) {
    throw new Error('Could not identify the active user email.');
  }
  if (normalizeEmail_(activeEmail) !== normalizeEmail_(approver.email)) {
    throw new Error('Only the current approver can sign this document.');
  }

  const documentId = state.documentId || DocumentApp.getActiveDocument().getId();
  if (documentId !== DocumentApp.getActiveDocument().getId()) {
    throw new Error('The saved workflow state belongs to another document.');
  }

  try {
    insertSignatureIntoTable_(payload.imageBase64, approver, payload.mode);
  } catch (error) {
    throw buildDetailedError_(
      'The signature could not be inserted into the approval table.',
      error
    );
  }

  state.currentStep += 1;
  state.updatedAt = new Date().toISOString();
  state.lastSignedBy = approver.email;

  if (state.currentStep < state.approvers.length) {
    try {
      saveApprovalState_(state);
    } catch (error) {
      throw buildDetailedError_(
        'The document workflow state could not be updated after signing.',
        error
      );
    }

    try {
      sendNextApproverEmail_(state);
    } catch (error) {
      throw buildDetailedError_(
        'The signature was recorded, but the next approver email could not be sent.',
        error
      );
    }

    return {
      done: false,
      message: 'Signature recorded. ' + state.approvers[state.currentStep].name + ' has been notified.'
    };
  }

  try {
    saveApprovalState_(state);
  } catch (error) {
    throw buildDetailedError_(
      'The final approval state could not be saved.',
      error
    );
  }

  let pdfFile;
  try {
    pdfFile = finalizeAndExportApprovedPdf_();
  } catch (error) {
    throw buildDetailedError_(
      'The document was signed, but the final approved PDF could not be created.',
      error
    );
  }

  try {
    notifyAllApproved_(state, pdfFile);
  } catch (error) {
    throw buildDetailedError_(
      'The document was fully approved, but the completion email could not be sent.',
      error
    );
  }

  return {
    done: true,
    message: 'Signature recorded. The document is now fully approved.'
  };
}

function rejectDocument(payload) {
  const state = getApprovalState_();
  const approver = state.approvers[state.currentStep];
  if (!approver) {
    throw new Error('There is no active approver to reject this document.');
  }

  const activeEmail = getActiveUserEmail_();
  if (!activeEmail) {
    throw new Error('Could not identify the active user email.');
  }
  if (normalizeEmail_(activeEmail) !== normalizeEmail_(approver.email)) {
    throw new Error('Only the current approver can reject this document.');
  }

  clearAllApprovalMarks_();

  const reviewListFile = createOpenReviewListFromComments_(
    DocumentApp.getActiveDocument(),
    payload && payload.reason ? String(payload.reason) : '',
    approver
  );

  state.currentStep = 0;
  state.updatedAt = new Date().toISOString();
  state.lastSignedBy = '';
  saveApprovalState_(state);

  sendRejectionEmail_(state, approver, reviewListFile, payload && payload.reason ? String(payload.reason) : '');

  return {
    done: false,
    rejected: true,
    message: reviewListFile
      ? 'Document rejected. Prior signatures were cleared, the first approver was notified, and an Open Review List was created.'
      : 'Document rejected. Prior signatures were cleared and the first approver was notified.'
  };
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getApprovalContext_() {
  const state = getApprovalState_();
  const activeEmail = getActiveUserEmail_();
  const currentApprover = state.approvers[state.currentStep] || null;
  const canSign = Boolean(
    currentApprover &&
    activeEmail &&
    normalizeEmail_(activeEmail) === normalizeEmail_(currentApprover.email)
  );

  return {
    initialized: state.approvers.length > 0,
    currentStep: state.currentStep,
    totalSteps: state.approvers.length,
    activeUserEmail: activeEmail,
    currentApprover: currentApprover,
    canSign: canSign,
    documentName: DocumentApp.getActiveDocument().getName(),
    approvers: state.approvers
  };
}

function getApprovalState_() {
  const props = PropertiesService.getDocumentProperties();
  const approversRaw = props.getProperty(APPROVERS_KEY);
  if (approversRaw) {
    const approvers = JSON.parse(approversRaw) || [];
    const state = {
      currentStep: Number(props.getProperty(STEP_KEY) || 0),
      approvers: approvers,
      emailList: JSON.parse(props.getProperty(EMAIL_LIST_KEY) || '[]'),
      documentId: props.getProperty(DOCUMENT_ID_KEY) || DocumentApp.getActiveDocument().getId(),
      updatedAt: props.getProperty(UPDATED_AT_KEY) || '',
      lastSignedBy: props.getProperty(LAST_SIGNED_BY_KEY) || ''
    };
    if (!state.emailList.length) {
      state.emailList = approvers.map(function(approver) {
        return approver.email;
      });
    }
    return state;
  }

  const state = buildApprovalState_();
  state.initialized = false;
  return state;
}

function buildApprovalState_() {
  const doc = DocumentApp.getActiveDocument();
  const body = getDocumentBody_(doc);
  const approvers = extractApproversFromTables_(body);

  return {
    currentStep: 0,
    approvers: approvers,
    emailList: approvers.map(function(approver) {
      return approver.email;
    }),
    documentId: doc.getId(),
    updatedAt: new Date().toISOString(),
    initialized: true
  };
}

function saveApprovalState_(state) {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty(STEP_KEY, String(state.currentStep || 0));
  props.setProperty(EMAIL_LIST_KEY, JSON.stringify(state.emailList || []));
  props.setProperty(DOCUMENT_ID_KEY, state.documentId || DocumentApp.getActiveDocument().getId());
  props.setProperty(APPROVERS_KEY, JSON.stringify(state.approvers || []));
  props.setProperty(UPDATED_AT_KEY, state.updatedAt || new Date().toISOString());
  if (state.lastSignedBy) {
    props.setProperty(LAST_SIGNED_BY_KEY, state.lastSignedBy);
  }
}

function buildDetailedError_(prefix, error) {
  const detail = error && error.message ? String(error.message) : String(error || 'Unknown error');
  return new Error(prefix + ' Details: ' + detail);
}

function extractApproversFromTables_(body) {
  const approvers = [];
  const seenEmails = {};
  const tableCount = body.getNumChildren();

  for (let childIndex = 0; childIndex < tableCount; childIndex += 1) {
    const child = body.getChild(childIndex);
    if (child.getType() !== DocumentApp.ElementType.TABLE) {
      continue;
    }

    const table = child.asTable();
    for (let rowIndex = 0; rowIndex < table.getNumRows(); rowIndex += 1) {
      const row = table.getRow(rowIndex);
      for (let cellIndex = 0; cellIndex < row.getNumCells(); cellIndex += 1) {
        const cell = row.getCell(cellIndex);
        const links = getCellMailtoLinks_(cell);
        links.forEach(function(link) {
          const normalizedEmail = normalizeEmail_(link.email);
          if (!normalizedEmail || seenEmails[normalizedEmail]) {
            return;
          }

          seenEmails[normalizedEmail] = true;
          approvers.push({
            name: link.text || normalizedEmail,
            email: normalizedEmail,
            tableIndex: childIndex,
            rowIndex: rowIndex,
            nameCellIndex: cellIndex,
            signatureRowIndex: resolveSignatureRowIndex_(table, rowIndex, cellIndex),
            signatureCellIndex: cellIndex,
            noteRowIndex: resolveNoteRowIndex_(table, rowIndex, cellIndex),
            noteCellIndex: cellIndex
          });
        });
      }
    }
  }

  return approvers.slice(0, 4);
}

function buildNoApproverMessage_() {
  const diagnostics = inspectDocumentLinks_();
  return [
    'No approvers with mailto links or person chips were found inside tables.',
    'Tables scanned: ' + diagnostics.tableCount + '.',
    'Detected approver links across the document: ' + diagnostics.mailtoCount + '.',
    diagnostics.samples.length
      ? 'Samples: ' + diagnostics.samples.join(' | ')
      : 'Apps Script could not read any mailto links or person chips.',
    'Make sure each approver is inserted as a real person chip or a mailto hyperlink, not plain text or an embedded visual element.'
  ].join('\n');
}

function debugApprovalLinks() {
  const diagnostics = inspectDocumentLinks_();
  const lines = [
    'Tables scanned: ' + diagnostics.tableCount,
    'Detected approver links in the document: ' + diagnostics.mailtoCount
  ];

  if (diagnostics.samples.length) {
    lines.push('Samples:');
    diagnostics.samples.forEach(function(sample) {
      lines.push('- ' + sample);
    });
  } else {
    lines.push('Apps Script could not read any mailto links or person chips.');
  }

  DocumentApp.getUi().alert(lines.join('\n'));
}

function getCellMailtoLinks_(cell) {
  const links = [];
  collectMailtoLinksFromElement_(cell, links);
  return links;
}

function collectMailtoLinksFromElement_(element, links) {
  const type = element.getType();
  if (type === DocumentApp.ElementType.PERSON) {
    const person = element.asPerson();
    const email = normalizeEmail_(person.getEmail());
    if (email) {
      links.push({
        email: email,
        text: person.getName() || email
      });
    }
    return;
  }

  if (type === DocumentApp.ElementType.TEXT) {
    collectMailtoLinksFromText_(element.asText(), links);
    return;
  }

  if (typeof element.getNumChildren !== 'function') {
    return;
  }

  for (let i = 0; i < element.getNumChildren(); i += 1) {
    collectMailtoLinksFromElement_(element.getChild(i), links);
  }
}

function collectMailtoLinksFromText_(text, links) {
  const textContent = text.getText();
  if (!textContent) {
    return;
  }

  const indices = text.getTextAttributeIndices();
  if (!indices.length || indices[0] !== 0) {
    indices.unshift(0);
  }

  for (let i = 0; i < indices.length; i += 1) {
    const start = indices[i];
    const end = i + 1 < indices.length ? indices[i + 1] - 1 : textContent.length - 1;
    const url = text.getLinkUrl(start);
    if (!url || !/^mailto:/i.test(url)) {
      continue;
    }

    const segmentText = textContent.substring(start, end + 1).trim();
    const email = url.replace(/^mailto:/i, '').split('?')[0].trim();
    if (!email) {
      continue;
    }

    links.push({
      email: email,
      text: segmentText
    });
  }
}

function resolveSignatureRowIndex_(table, nameRowIndex, cellIndex) {
  if (nameRowIndex > 0) {
    const candidateCell = table.getRow(nameRowIndex - 1).getCell(cellIndex);
    if (isCellMostlyEmpty_(candidateCell)) {
      return nameRowIndex - 1;
    }
  }
  return nameRowIndex;
}

function resolveNoteRowIndex_(table, nameRowIndex, cellIndex) {
  for (let offset = 0; offset <= 2; offset += 1) {
    const candidateRowIndex = nameRowIndex + offset;
    if (candidateRowIndex >= table.getNumRows()) {
      break;
    }

    const candidateCell = table.getRow(candidateRowIndex).getCell(cellIndex);
    if (/date\s*:/i.test(candidateCell.getText())) {
      return candidateRowIndex;
    }
  }

  return nameRowIndex;
}

function insertSignatureIntoTable_(imageBase64, approver, mode) {
  const doc = DocumentApp.getActiveDocument();
  const body = getDocumentBody_(doc);
  const table = body.getChild(approver.tableIndex).asTable();
  const signatureRow = table.getRow(approver.signatureRowIndex);
  const signatureCell = signatureRow.getCell(approver.signatureCellIndex);
  const noteRow = table.getRow(approver.noteRowIndex);
  const noteCell = noteRow.getCell(approver.noteCellIndex);

  clearSignatureImagesFromCell_(signatureCell);
  clearManagedSignedNotes_(noteCell);
  replaceDateLineWithSignedNote_(noteCell);

  const imageBlob = Utilities.newBlob(
    Utilities.base64Decode(imageBase64),
    'image/png',
    'signature-' + approver.email + '.png'
  );

  const insertAt = Math.min(1, signatureCell.getNumChildren());
  const paragraph = signatureCell.insertParagraph(insertAt, '');
  paragraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  const image = paragraph.appendInlineImage(imageBlob);
  const originalWidth = image.getWidth();
  const originalHeight = image.getHeight();
  const scale = SIGNATURE_IMAGE_WIDTH / originalWidth;
  image.setWidth(Math.round(originalWidth * scale));
  image.setHeight(Math.round(originalHeight * scale));

  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone() || 'America/Mexico_City',
    "yyyy-MM-dd HH:mm:ss"
  );

  const detail = noteCell.appendParagraph(
    SIGNED_NOTE_PREFIX + timestamp + ' via ' + (mode === 'type' ? 'typed signature' : 'drawn signature')
  );
  detail.setFontSize(8);
  detail.setForegroundColor('#5f6368');
}

function sendNextApproverEmail_(state) {
  const nextApprover = state.approvers[state.currentStep];
  if (!nextApprover || !nextApprover.email) {
    return;
  }

  const doc = DocumentApp.getActiveDocument();
  const docName = doc.getName();
  const url = doc.getUrl();
  const signedApprovers = state.approvers.slice(0, state.currentStep).map(function(approver) {
    return approver.name;
  });
  const currentStepLabel = state.currentStep + 1;
  const htmlBody = [
    '<div style="font-family:Arial,sans-serif;background:#f6f8fb;padding:24px;color:#202124;">',
    '<div style="max-width:640px;margin:0 auto;background:#ffffff;border:1px solid #dadce0;border-radius:18px;overflow:hidden;">',
    getCompanyBrandHtml_(),
    '<div style="padding:24px;">',
    '<p style="margin:0 0 12px;">Hello ' + sanitizeHtml_(nextApprover.name) + ',</p>',
    '<p style="margin:0 0 12px;">It is now your turn to approve <strong>' + sanitizeHtml_(doc.getName()) + '</strong>.</p>',
    '<p style="margin:0 0 18px;">Current step: <strong>' + currentStepLabel + ' of ' + state.approvers.length + '</strong></p>',
    '<div style="margin:0 0 18px;padding:16px;border:1px solid #e3e7ee;border-radius:14px;background:#fafbff;">',
    '<div style="font-size:12px;color:#5f6368;margin-bottom:8px;">Completed approvals</div>',
    '<div style="font-size:14px;color:#202124;">' + sanitizeHtml_(signedApprovers.length ? signedApprovers.join(', ') : 'None yet') + '</div>',
    '</div>',
    '<p style="margin:0 0 20px;"><a href="' + url + '" style="display:inline-block;padding:12px 20px;background:#1a73e8;color:#ffffff;text-decoration:none;border-radius:999px;font-weight:600;">Open document</a></p>',
    '<p style="margin:0;color:#5f6368;">When you open the document, the signature sidebar will appear automatically if it is your turn.</p>',
    '</div>',
    '</div>',
    '</div>'
  ].join('');

  MailApp.sendEmail({
    to: nextApprover.email,
    subject: 'Action required: document approval - ' + docName,
    htmlBody: htmlBody
  });
}

function sendRejectionEmail_(state, rejectingApprover, reviewListFile, reason) {
  const firstApprover = state.approvers[0];
  if (!firstApprover || !firstApprover.email) {
    return;
  }

  const doc = DocumentApp.getActiveDocument();
  const docName = doc.getName();
  const reasonHtml = reason
    ? '<div style="margin:0 0 18px;padding:16px;border:1px solid #f1d7b4;border-radius:14px;background:#fff8ef;"><div style="font-size:12px;color:#8b5e1a;margin-bottom:8px;">Rejection note</div><div style="font-size:14px;color:#202124;">' + sanitizeHtml_(reason) + '</div></div>'
    : '';
  const reviewListHtml = reviewListFile
    ? '<p style="margin:0 0 12px;"><a href="' + reviewListFile.getUrl() + '" style="color:#1a73e8;text-decoration:none;">Open Review List</a></p>'
    : '';

  const htmlBody = [
    '<div style="font-family:Arial,sans-serif;background:#f6f8fb;padding:24px;color:#202124;">',
    '<div style="max-width:640px;margin:0 auto;background:#ffffff;border:1px solid #dadce0;border-radius:18px;overflow:hidden;">',
    getCompanyBrandHtml_(),
    '<div style="padding:24px;">',
    '<p style="margin:0 0 12px;">Hello ' + sanitizeHtml_(firstApprover.name) + ',</p>',
    '<p style="margin:0 0 12px;"><strong>' + sanitizeHtml_(doc.getName()) + '</strong> was rejected by ' + sanitizeHtml_(rejectingApprover.name) + ' and has been reset for rework.</p>',
    reasonHtml,
    '<p style="margin:0 0 12px;"><a href="' + doc.getUrl() + '" style="display:inline-block;padding:12px 20px;background:#1a73e8;color:#ffffff;text-decoration:none;border-radius:999px;font-weight:600;">Open document</a></p>',
    reviewListHtml,
    '<p style="margin:0;color:#5f6368;">All approval signatures generated by the workflow were cleared. Please update the document and restart the approval cycle when ready.</p>',
    '</div>',
    '</div>',
    '</div>'
  ].join('');

  MailApp.sendEmail({
    to: firstApprover.email,
    subject: 'Document returned for rework - ' + docName,
    htmlBody: htmlBody
  });
}

function notifyAllApproved_(state, pdfFile) {
  const doc = DocumentApp.getActiveDocument();
  const docName = doc.getName();
  const recipients = state.emailList.join(',');
  const completedBy = state.approvers.map(function(approver) {
    return approver.name;
  }).join(', ');
  const htmlBody = [
    '<div style="font-family:Arial,sans-serif;background:#f6f8fb;padding:24px;color:#202124;">',
    '<div style="max-width:640px;margin:0 auto;background:#ffffff;border:1px solid #dadce0;border-radius:18px;overflow:hidden;">',
    getCompanyBrandHtml_(),
    '<div style="padding:24px;">',
    '<p style="margin:0 0 12px;">The document <strong>' + sanitizeHtml_(doc.getName()) + '</strong> is now fully approved.</p>',
    '<div style="margin:0 0 18px;padding:16px;border:1px solid #e3e7ee;border-radius:14px;background:#fafbff;">',
    '<div style="font-size:12px;color:#5f6368;margin-bottom:8px;">Approved by</div>',
    '<div style="font-size:14px;color:#202124;">' + sanitizeHtml_(completedBy) + '</div>',
    '</div>',
    '<p style="margin:0 0 12px;"><a href="' + doc.getUrl() + '" style="color:#1a73e8;text-decoration:none;">Open document</a></p>',
    '<p style="margin:0;"><a href="' + pdfFile.getUrl() + '" style="color:#1a73e8;text-decoration:none;">Open final PDF</a></p>',
    '</div>',
    '</div>',
    '</div>'
  ].join('');

  MailApp.sendEmail({
    to: recipients,
    subject: 'Document fully approved - ' + docName,
    htmlBody: htmlBody
  });
}

function finalizeAndExportApprovedPdf_() {
  const activeDoc = DocumentApp.getActiveDocument();
  const documentId = activeDoc.getId();

  activeDoc.saveAndClose();

  // Reopen the document after flushing changes so the exported PDF includes
  // the last approver's signature.
  const doc = DocumentApp.openById(documentId);
  const pdfBlob = doc.getAs(MimeType.PDF)
    .setName(doc.getName() + ' - Approved.pdf');
  const folder = getOrCreatePdfFolder_(documentId);
  return folder.createFile(pdfBlob);
}

function getOrCreatePdfFolder_(documentId) {
  const docFile = DriveApp.getFileById(documentId);
  const parents = docFile.getParents();
  const parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const existing = parentFolder.getFoldersByName(PDF_FOLDER_NAME);
  return existing.hasNext() ? existing.next() : parentFolder.createFolder(PDF_FOLDER_NAME);
}

function getActiveUserEmail_() {
  return (Session.getActiveUser().getEmail() || '').trim();
}

function inspectDocumentLinks_() {
  const body = getDocumentBody_(DocumentApp.getActiveDocument());
  const links = [];
  collectMailtoLinksFromElement_(body, links);

  return {
    tableCount: countTables_(body),
    mailtoCount: links.length,
    samples: links.slice(0, 8).map(function(link) {
      return (link.text || '[sin texto]') + ' <' + link.email + '>';
    })
  };
}

function countTables_(body) {
  let count = 0;
  for (let i = 0; i < body.getNumChildren(); i += 1) {
    if (body.getChild(i).getType() === DocumentApp.ElementType.TABLE) {
      count += 1;
    }
  }
  return count;
}

function getDocumentBody_(doc) {
  if (typeof doc.getActiveTab === 'function') {
    return doc.getActiveTab().asDocumentTab().getBody();
  }
  return doc.getBody();
}

function isCellMostlyEmpty_(cell) {
  const text = normalizeWhitespace_(cell.getText());
  return !text;
}

function normalizeWhitespace_(value) {
  return String(value || '').replace(/\s+/g, ' ').trim();
}

function normalizeEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function sanitizeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function getCompanyBrandHtml_() {
  return [
    '<div style="padding:20px 24px;border-bottom:1px solid #eceff3;background:linear-gradient(180deg,#ffffff 0%,#fbfbfc 100%);">',
    '<div style="font-family:Georgia,Times,serif;font-size:36px;line-height:1;color:#1e1e1e;letter-spacing:-1px;">',
    '<span>conceiv</span><span style="color:#d79b43;">able</span>',
    '<span style="display:inline-block;margin-left:8px;font-family:Arial,sans-serif;font-size:11px;line-height:1.1;letter-spacing:.08em;color:#4b4b4b;vertical-align:middle;">LIFE<br>SCIENCES</span>',
    '</div>',
    '</div>'
  ].join('');
}

function createOpenReviewListFromComments_(sourceDoc, rejectionReason, rejectingApprover) {
  const comments = listOpenComments_(sourceDoc.getId());
  if (!comments.length && !rejectionReason) {
    return null;
  }

  const parentFolder = getPrimaryParentFolder_(sourceDoc.getId());
  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone() || 'America/Mexico_City',
    'yyyyMMdd-HHmm'
  );
  const fileName = 'ORL_' + sanitizeFileName_(sourceDoc.getName()) + '_' + timestamp;
  const orlDoc = DocumentApp.create(fileName);
  const orlFile = DriveApp.getFileById(orlDoc.getId());
  parentFolder.addFile(orlFile);
  DriveApp.getRootFolder().removeFile(orlFile);

  copyHeaderFromSource_(sourceDoc, orlDoc);
  applyOrlHeaderTemplate_(orlDoc, sourceDoc.getName(), rejectingApprover.name);
  setOrlFooterText_(orlDoc, 'Generated on ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/Mexico_City', 'yyyy-MM-dd HH:mm') + ' by ' + rejectingApprover.name + '.');

  const body = getDocumentBody_(orlDoc);
  body.clear();

  if (rejectionReason) {
    body.appendParagraph('Rejection note: ' + rejectionReason)
      .setForegroundColor('#8b5e1a');
  }

  const table = body.appendTable([
    ['Finding', 'Section', 'Status']
  ]);
  styleOrlHeaderRow_(table.getRow(0));

  const sourceBody = getDocumentBody_(sourceDoc);
  comments.forEach(function(comment) {
    const finding = buildCommentFindingText_(comment);
    const section = findSectionLabelForComment_(sourceBody, comment);
    const row = table.appendTableRow();
    row.appendTableCell(finding);
    row.appendTableCell(section);
    row.appendTableCell(ORL_STATUS_OPEN);
  });

  if (!comments.length && rejectionReason) {
    const row = table.appendTableRow();
    row.appendTableCell(rejectionReason);
    row.appendTableCell('General');
    row.appendTableCell(ORL_STATUS_OPEN);
  }

  orlDoc.saveAndClose();
  return orlFile;
}

function listOpenComments_(fileId) {
  if (typeof Drive === 'undefined' || !Drive.Comments || typeof Drive.Comments.list !== 'function') {
    return [];
  }

  const comments = [];
  let pageToken = null;
  do {
    const response = Drive.Comments.list(fileId, {
      pageToken: pageToken,
      includeDeleted: false,
      pageSize: 100,
      fields: 'nextPageToken,comments(id,content,quotedFileContent,resolved,deleted,anchor,author/displayName,author/emailAddress,createdTime,replies)'
    });
    const items = response.comments || [];
    items.forEach(function(comment) {
      if (!comment.deleted && !comment.resolved) {
        comments.push(comment);
      }
    });
    pageToken = response.nextPageToken || null;
  } while (pageToken);

  return comments;
}

function buildCommentFindingText_(comment) {
  const author = comment.author && comment.author.displayName ? comment.author.displayName : 'Reviewer';
  const content = normalizeWhitespace_(comment.content || '');
  return author + ': ' + (content || '[No comment text]');
}

function findSectionLabelForComment_(body, comment) {
  const quoted = normalizeWhitespace_(
    comment.quotedFileContent && comment.quotedFileContent.value
      ? comment.quotedFileContent.value
      : ''
  );

  if (quoted) {
    const located = findNearestSectionLabelForText_(body, quoted);
    if (located) {
      return located;
    }
  }

  return 'General';
}

function findNearestSectionLabelForText_(body, text) {
  const escaped = escapeRegExp_(text.substring(0, 80));
  const found = body.findText(escaped);
  if (!found) {
    return '';
  }

  let element = found.getElement();
  while (element && typeof element.getParent === 'function' && element.getType() !== DocumentApp.ElementType.PARAGRAPH && element.getType() !== DocumentApp.ElementType.LIST_ITEM) {
    element = element.getParent();
  }
  if (!element) {
    return '';
  }

  const paragraph = element.getType() === DocumentApp.ElementType.LIST_ITEM
    ? element.asListItem()
    : element.asParagraph();

  const container = paragraph.getParent();
  if (!container || typeof container.getChildIndex !== 'function') {
    return normalizeWhitespace_(paragraph.getText()) || '';
  }

  let index = container.getChildIndex(paragraph);
  while (index >= 0) {
    const candidate = container.getChild(index);
    const label = getSectionLabelFromElement_(candidate);
    if (label) {
      return label;
    }
    index -= 1;
  }

  return normalizeWhitespace_(paragraph.getText()) || '';
}

function getSectionLabelFromElement_(element) {
  const type = element.getType();
  if (type !== DocumentApp.ElementType.PARAGRAPH && type !== DocumentApp.ElementType.LIST_ITEM) {
    return '';
  }

  const paragraph = type === DocumentApp.ElementType.LIST_ITEM ? element.asListItem() : element.asParagraph();
  const text = normalizeWhitespace_(paragraph.getText());
  if (!text) {
    return '';
  }

  if (/^\d+(\.\d+)*\.?\s+/.test(text)) {
    return text;
  }

  if (type === DocumentApp.ElementType.PARAGRAPH) {
    const heading = paragraph.getHeading();
    if (heading && heading !== DocumentApp.ParagraphHeading.NORMAL) {
      return text;
    }
  }

  return '';
}

function copyHeaderFromSource_(sourceDoc, targetDoc) {
  const sourceHeader = sourceDoc.getHeader();
  if (!sourceHeader) {
    return;
  }

  const targetHeader = targetDoc.getHeader() || targetDoc.addHeader();
  targetHeader.clear();
  for (let i = 0; i < sourceHeader.getNumChildren(); i += 1) {
    const child = sourceHeader.getChild(i);
    const type = child.getType();
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      targetHeader.appendParagraph(child.asParagraph().copy());
    } else if (type === DocumentApp.ElementType.TABLE) {
      targetHeader.appendTable(child.asTable().copy());
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      targetHeader.appendListItem(child.asListItem().copy());
    }
  }
}

function applyOrlHeaderTemplate_(doc, sourceDocName, generatedByName) {
  const header = doc.getHeader();
  if (!header) {
    return;
  }

  const generationDate = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone() || 'America/Mexico_City',
    'yyyy-MM-dd'
  );

  for (let i = 0; i < header.getNumChildren(); i += 1) {
    const child = header.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.TABLE) {
      continue;
    }

    const table = child.asTable();
    for (let rowIndex = 0; rowIndex < table.getNumRows(); rowIndex += 1) {
      const row = table.getRow(rowIndex);
      for (let cellIndex = 0; cellIndex < row.getNumCells(); cellIndex += 1) {
        const cell = row.getCell(cellIndex);
        const text = normalizeWhitespace_(cell.getText());

        if (text === 'DOCUMENT TYPE') {
          setCellSingleParagraphText_(cell, 'Open Review List', false);
          continue;
        }

        if (text.indexOf('DOCUMENT TITLE') !== -1) {
          setCellSingleParagraphText_(cell, 'Open Review List of "' + sourceDocName + '"', true);
          continue;
        }

        if (text === 'Issue date') {
          continue;
        }

        if (text === 'Expiration date') {
          setCellSingleParagraphText_(cell, 'Generated by', false);
          continue;
        }

        if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
          const leftNeighborText = cellIndex > 0 ? normalizeWhitespace_(row.getCell(cellIndex - 1).getText()) : '';
          if (leftNeighborText === 'Issue date') {
            setCellSingleParagraphText_(cell, generationDate, true);
            continue;
          }
        }

        if (text === '2029-04-21' || text === generatedByName || text === 'Generated by') {
          const leftNeighborText = cellIndex > 0 ? normalizeWhitespace_(row.getCell(cellIndex - 1).getText()) : '';
          if (leftNeighborText === 'Generated by') {
            setCellSingleParagraphText_(cell, generatedByName, true);
            continue;
          }
        }

        if (text === 'DOCUMENT TITLE') {
          hideParagraphText_(cell);
          continue;
        }
      }
    }
  }
}

function setCellSingleParagraphText_(cell, text, bold) {
  clearCellChildren_(cell);
  const paragraph = cell.appendParagraph(text);
  paragraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  paragraph.setBold(Boolean(bold));
}

function hideParagraphText_(cell) {
  for (let i = 0; i < cell.getNumChildren(); i += 1) {
    const child = cell.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) {
      continue;
    }

    const paragraph = child.asParagraph();
    const textElement = paragraph.editAsText();
    textElement.setForegroundColor('#ffffff');
    textElement.setText('');
  }
}

function clearCellChildren_(cell) {
  for (let i = cell.getNumChildren() - 1; i >= 0; i -= 1) {
    cell.removeChild(cell.getChild(i));
  }
}

function setOrlFooterText_(doc, footerText) {
  const footer = doc.getFooter() || doc.addFooter();
  footer.clear();
  const paragraph = footer.appendParagraph(footerText);
  paragraph.setForegroundColor('#5f6368');
  paragraph.setFontSize(9);
}

function styleOrlHeaderRow_(row) {
  for (let i = 0; i < row.getNumCells(); i += 1) {
    const cell = row.getCell(i);
    cell.setBackgroundColor('#eef3fb');
    const text = cell.editAsText();
    text.setBold(true);
    text.setForegroundColor('#1f3b64');
  }
}

function getPrimaryParentFolder_(fileId) {
  const file = DriveApp.getFileById(fileId);
  const parents = file.getParents();
  return parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
}

function sanitizeFileName_(value) {
  return String(value || '')
    .replace(/[\\/:*?"<>|#%]/g, '')
    .replace(/\s+/g, '_')
    .substring(0, 80);
}

function escapeRegExp_(value) {
  return String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function clearAllApprovalMarks_() {
  const state = getApprovalState_();
  const doc = DocumentApp.getActiveDocument();
  const body = getDocumentBody_(doc);

  state.approvers.forEach(function(approver) {
    if (approver.tableIndex >= body.getNumChildren()) {
      return;
    }

    const tableElement = body.getChild(approver.tableIndex);
    if (tableElement.getType() !== DocumentApp.ElementType.TABLE) {
      return;
    }

    const table = tableElement.asTable();
    if (approver.signatureRowIndex < table.getNumRows()) {
      const signatureCell = table.getRow(approver.signatureRowIndex).getCell(approver.signatureCellIndex);
      clearSignatureImagesFromCell_(signatureCell);
    }
    if (approver.noteRowIndex < table.getNumRows()) {
      const noteCell = table.getRow(approver.noteRowIndex).getCell(approver.noteCellIndex);
      clearManagedSignedNotes_(noteCell);
      restoreEmptyDateLineIfNeeded_(noteCell);
    }
  });
}

function clearSignatureImagesFromCell_(cell) {
  for (let i = cell.getNumChildren() - 1; i >= 0; i -= 1) {
    const child = cell.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) {
      continue;
    }

    const paragraph = child.asParagraph();
    let removedInlineImage = false;
    for (let j = paragraph.getNumChildren() - 1; j >= 0; j -= 1) {
      if (paragraph.getChild(j).getType() === DocumentApp.ElementType.INLINE_IMAGE) {
        paragraph.removeChild(paragraph.getChild(j));
        removedInlineImage = true;
      }
    }

    if (removedInlineImage && !normalizeWhitespace_(paragraph.getText())) {
      cell.removeChild(paragraph);
    }
  }
}

function clearManagedSignedNotes_(cell) {
  for (let i = cell.getNumChildren() - 1; i >= 0; i -= 1) {
    const child = cell.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) {
      continue;
    }

    const paragraph = child.asParagraph();
    const text = normalizeWhitespace_(paragraph.getText());
    if (text.indexOf(SIGNED_NOTE_PREFIX) === 0) {
      cell.removeChild(paragraph);
    }
  }
}

function replaceDateLineWithSignedNote_(cell) {
  for (let i = cell.getNumChildren() - 1; i >= 0; i -= 1) {
    const child = cell.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) {
      continue;
    }

    const paragraph = child.asParagraph();
    const text = normalizeWhitespace_(paragraph.getText());
    if (/^date\s*:/i.test(text)) {
      cell.removeChild(paragraph);
      return;
    }
  }
}

function restoreEmptyDateLineIfNeeded_(cell) {
  if (/date\s*:/i.test(cell.getText()) || cell.getText().indexOf(SIGNED_NOTE_PREFIX) !== -1) {
    return;
  }

  const paragraph = cell.appendParagraph('Date:');
  paragraph.setFontSize(10);
}
