const SETTINGS = {
  columns: {
    domain: 1,          // A: ドメイン / URL / メール
    kanjiName: 2,       // B: 氏名
    email: 3,           // C: 送信先メールアドレス
    status: 4,          // D: ステータス
    company: 5,         // E: 会社名
    candidateStart: 6   // F以降: 候補メールアドレス
  },
  statuses: {
    sent: '送信済み',
    skip: 'スキップ',
    review: '要確認',
    noCandidate: '候補なし',
    ready: '送信準備完了'
  },
  autoSendFirstCandidate: false,
  subjectTemplate: 'ご連絡のお願い | 松本高専 電気電子工学科',
  bodyTemplate: [
    '{company} {kanji_name} 様',
    '',
    '突然のご連絡失礼いたします。松本高専 電気電子工学科の吉本舜一と申します。',
    '現在、学校で学んでいるビジネスコンテストやスタートアップに関する取り組みの一環として、企業の皆様へご相談のメールをお送りしております。',
    'このたび、貴社で学ぶ機会をぜひいただけないかと考え、ご連絡いたしました。大企業だけでなく、中小企業や地域企業についても広く知りたいと思っております。',
    '{kanji_name} 様のご都合がよろしければ、ぜひ一度、会社見学やキャリアについてお話を伺えますと幸いです。',
    '所要時間は 20 分から 30 分ほどを想定しております。オンライン・対面いずれでも対応可能です。',
    'お忙しいところ恐れ入りますが、ご検討のほどよろしくお願いいたします。',
    '',
    '吉本舜一'
  ].join('\n')
};

function runAutoMailSystem() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const context = getSheetContext_(sheet);
  if (!context) return;

  const summary = {
    sent: [],
    review: [],
    skipped: [],
    noCandidate: []
  };

  context.rows.forEach((rowData, index) => {
    const rowNumber = index + 2;
    const result = processRow_(sheet, rowNumber, rowData);
    if (result.sent) summary.sent.push(rowNumber);
    if (result.review) summary.review.push(rowNumber);
    if (result.skipped) summary.skipped.push(rowNumber);
    if (result.noCandidate) summary.noCandidate.push(rowNumber);
  });

  SpreadsheetApp.getUi().alert(
    [
      '自動送信処理が完了しました。',
      '送信: ' + summary.sent.length + ' 件',
      '要確認: ' + summary.review.length + ' 件',
      'スキップ: ' + summary.skipped.length + ' 件',
      '候補なし: ' + summary.noCandidate.length + ' 件'
    ].join('\n')
  );
}

function sendPendingEmails() {
  runAutoMailSystem();
}

function generateEmailCandidatesForActiveRow() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveCell().getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert('ヘッダー行ではなく、対象データ行を選択してください。');
    return;
  }

  const emails = writeCandidatesForRow_(sheet, row, true);
  if (emails.length > 0 && !String(sheet.getRange(row, SETTINGS.columns.email).getValue() || '').trim()) {
    sheet.getRange(row, SETTINGS.columns.email).setValue(emails[0]);
    sheet.getRange(row, SETTINGS.columns.status).setValue(SETTINGS.statuses.review);
  }
}

function generateEmailCandidatesForAllRows() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const context = getSheetContext_(sheet);
  if (!context) return;

  let processed = 0;
  context.rows.forEach((rowData, index) => {
    const rowNumber = index + 2;
    if (!rowData.domain || !rowData.kanjiName) return;
    const emails = writeCandidatesForRow_(sheet, rowNumber, false);
    if (emails.length > 0 && !rowData.email) {
      sheet.getRange(rowNumber, SETTINGS.columns.email).setValue(emails[0]);
      sheet.getRange(rowNumber, SETTINGS.columns.status).setValue(SETTINGS.statuses.review);
    }
    processed++;
  });

  SpreadsheetApp.getUi().alert(processed + ' 行の候補メールアドレスを生成しました。');
}

function previewEmailForActiveRow() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveCell().getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert('ヘッダー行ではなく、対象データ行を選択してください。');
    return;
  }

  const company = String(sheet.getRange(row, SETTINGS.columns.company).getValue() || '').trim();
  const kanjiName = String(sheet.getRange(row, SETTINGS.columns.kanjiName).getValue() || '').trim();
  const email = String(sheet.getRange(row, SETTINGS.columns.email).getValue() || '').trim();

  const subject = buildSubject_(company, kanjiName);
  const body = buildBody_(company, kanjiName);

  SpreadsheetApp.getUi().alert(
    [
      '宛先: ' + (email || '(未入力)'),
      '件名: ' + subject,
      '',
      body
    ].join('\n')
  );
}

function resetStatusForActiveRow() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveCell().getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert('ヘッダー行ではなく、対象データ行を選択してください。');
    return;
  }

  sheet.getRange(row, SETTINGS.columns.status).clearContent();
  SpreadsheetApp.getUi().alert('選択行のステータスをクリアしました。');
}

function processRow_(sheet, row, rowData) {
  const statuses = SETTINGS.statuses;
  const result = { sent: false, review: false, skipped: false, noCandidate: false };

  if (normalizeStatus_(rowData.status) === normalizeStatus_(statuses.sent)) {
    return result;
  }

  if (!rowData.kanjiName) {
    sheet.getRange(row, SETTINGS.columns.status).setValue(statuses.skip);
    result.skipped = true;
    return result;
  }

  let email = rowData.email;
  if (!isValidEmail_(email)) {
    const candidates = writeCandidatesForRow_(sheet, row, false);
    if (candidates.length === 0) {
      sheet.getRange(row, SETTINGS.columns.status).setValue(statuses.noCandidate);
      result.noCandidate = true;
      return result;
    }

    email = candidates[0];
    sheet.getRange(row, SETTINGS.columns.email).setValue(email);
    sheet.getRange(row, SETTINGS.columns.status).setValue(statuses.review);
    result.review = true;

    if (!SETTINGS.autoSendFirstCandidate) {
      return result;
    }
  }

  GmailApp.sendEmail(email, buildSubject_(rowData.company, rowData.kanjiName), buildBody_(rowData.company, rowData.kanjiName));
  sheet.getRange(row, SETTINGS.columns.status).setValue(statuses.sent);
  result.sent = true;
  result.review = false;
  return result;
}

function getSheetContext_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('送信対象データがありません。');
    return null;
  }

  const lastColumn = Math.max(sheet.getLastColumn(), SETTINGS.columns.candidateStart);
  const values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  const rows = values.map(row => ({
    domain: String(row[SETTINGS.columns.domain - 1] || '').trim(),
    kanjiName: String(row[SETTINGS.columns.kanjiName - 1] || '').trim(),
    email: String(row[SETTINGS.columns.email - 1] || '').trim(),
    status: String(row[SETTINGS.columns.status - 1] || '').trim(),
    company: String(row[SETTINGS.columns.company - 1] || '').trim()
  }));

  return { rows: rows };
}

function writeCandidatesForRow_(sheet, row, showAlert) {
  if (showAlert === undefined) showAlert = true;

  const domainInput = sheet.getRange(row, SETTINGS.columns.domain).getValue();
  const name = sheet.getRange(row, SETTINGS.columns.kanjiName).getValue();
  if (!domainInput || !name) {
    if (showAlert) {
      SpreadsheetApp.getUi().alert('A列のドメインと B列の氏名を入力してください。');
    }
    return [];
  }

  const emails = createEmailCandidates(name, domainInput);
  clearCandidateColumns_(sheet, row);

  if (emails.length > 0) {
    sheet.getRange(row, SETTINGS.columns.candidateStart, 1, emails.length).setValues([emails]);
  }

  if (showAlert) {
    SpreadsheetApp.getUi().alert(emails.length + ' 件の候補メールアドレスを生成しました。');
  }

  return emails;
}

function clearCandidateColumns_(sheet, row) {
  const maxColumns = sheet.getMaxColumns();
  const width = maxColumns - SETTINGS.columns.candidateStart + 1;
  if (width > 0) {
    sheet.getRange(row, SETTINGS.columns.candidateStart, 1, width).clearContent();
  }
}

function buildSubject_(company, kanjiName) {
  return applyTemplate_(SETTINGS.subjectTemplate, {
    company: company || 'ご担当者',
    kanji_name: kanjiName || 'ご担当者'
  });
}

function buildBody_(company, kanjiName) {
  return applyTemplate_(SETTINGS.bodyTemplate, {
    company: company || 'ご担当者',
    kanji_name: kanjiName || 'ご担当者'
  });
}

function applyTemplate_(template, values) {
  let output = String(template || '');
  Object.keys(values).forEach(key => {
    output = output.replace(new RegExp('\\{' + key + '\\}', 'g'), values[key]);
  });
  return output;
}

function createEmailCandidates(name, domainInput) {
  const parts = normalizeNameParts_(name);
  const first = parts.first;
  const last = parts.last;
  const firstInitial = first.charAt(0);
  const lastInitial = last.charAt(0);

  const domains = extractDomains_(domainInput);
  if (domains.length === 0) return [];

  const patterns = [];
  if (first && last) {
    patterns.push(first + '.' + last);
    patterns.push(first + last);
    patterns.push(first + '_' + last);
    patterns.push(first + '-' + last);
    patterns.push(firstInitial + last);
    patterns.push(firstInitial + '.' + last);
    patterns.push(firstInitial + '_' + last);
    patterns.push(firstInitial + '-' + last);
    patterns.push(first + lastInitial);
    patterns.push(first + '.' + lastInitial);
    patterns.push(first + '_' + lastInitial);
    patterns.push(first + '-' + lastInitial);
    patterns.push(last + '.' + first);
    patterns.push(last + first);
    patterns.push(last + '_' + first);
    patterns.push(last + '-' + first);
    patterns.push(last + firstInitial);
    patterns.push(last + '.' + firstInitial);
    patterns.push(first);
    patterns.push(last);
    patterns.push(firstInitial + lastInitial);
    patterns.push(lastInitial + firstInitial);
  } else if (first) {
    patterns.push(first);
    patterns.push(firstInitial);
  }

  const emails = [];
  const seen = new Set();
  patterns.forEach(pattern => {
    if (!pattern) return;
    domains.forEach(domain => {
      const email = pattern + '@' + domain;
      if (!seen.has(email)) {
        seen.add(email);
        emails.push(email);
      }
    });
  });

  return emails;
}

function normalizeNameParts_(name) {
  const raw = String(name || '').trim();
  if (!raw) return { first: '', last: '' };

  const cleaned = raw
    .replace(/[()\/]/g, ' ')
    .replace(/[,\u3001]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  const parts = cleaned.split(/[ \u3000]+/).filter(Boolean);
  const normalized = parts.map(part => {
    return part.toLowerCase().replace(/[^a-z0-9-]/g, '');
  }).filter(Boolean);

  if (normalized.length >= 2) {
    return { first: normalized[0], last: normalized[normalized.length - 1] };
  }

  if (normalized.length === 1) {
    return { first: normalized[0], last: '' };
  }

  const fallback = cleaned.toLowerCase().replace(/\s+/g, '').replace(/[^a-z0-9-]/g, '');
  return { first: fallback, last: '' };
}

function extractDomains_(input) {
  const raw = String(input || '').trim().toLowerCase();
  if (!raw) return [];

  const tokens = raw.split(/[,\s;]+/).filter(Boolean);
  const domains = [];

  tokens.forEach(tokenValue => {
    let token = tokenValue;
    const at = token.indexOf('@');
    if (at !== -1) token = token.slice(at + 1);

    if (/^https?:\/\//.test(token)) {
      const matched = token.match(/^https?:\/\/([^\/?#]+)/);
      if (matched && matched[1]) token = matched[1];
    }

    token = token.replace(/^www\./, '');

    if (!token.includes('.')) {
      token = token.replace(/\s+/g, '') + '.com';
    }

    if (/^[a-z0-9.-]+\.[a-z]{2,}$/.test(token)) {
      domains.push(token);
    }
  });

  return Array.from(new Set(domains));
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || '').trim());
}

function normalizeStatus_(status) {
  return String(status || '').trim().toLowerCase();
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('自動送信システム')
    .addItem('1. 選択行の候補メールを生成', 'generateEmailCandidatesForActiveRow')
    .addItem('2. 全行の候補メールを生成', 'generateEmailCandidatesForAllRows')
    .addItem('3. 選択行をプレビュー', 'previewEmailForActiveRow')
    .addSeparator()
    .addItem('4. 未送信メールを送信', 'sendPendingEmails')
    .addItem('5. 自動送信を実行', 'runAutoMailSystem')
    .addSeparator()
    .addItem('6. 選択行のステータスをクリア', 'resetStatusForActiveRow')
    .addToUi();
}
