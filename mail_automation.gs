const COL_DOMAIN = 1;           // A: ドメイン / URL / メール
const COL_NAME = 2;             // B: 氏名
const COL_EMAIL = 3;            // C: 送信先メール（確定）
const COL_STATUS = 4;           // D: 送信状態
const COL_COMPANY = 5;          // E: 会社名
const COL_CANDIDATE_START = 6;  // F以降: 候補メール

const STATUS_SENT = '送信済み';
const STATUS_SKIP = 'スキップ';
const STATUS_REVIEW = '要確認';
const AUTO_SEND_FIRST_CANDIDATE = false;

/**
 * メイン処理:
 * 1. C列にメールが入っていればそのまま送信
 * 2. C列が空なら候補を生成してF列以降へ出力
 * 3. 先頭候補をC列へ仮入力
 * 4. AUTO_SEND_FIRST_CANDIDATE が true のときだけ自動送信
 *
 * 想定シート構成:
 * A: ドメイン
 * B: 名前
 * C: 確定メールアドレス
 * D: ステータス
 * E: 会社名
 * F以降: 候補メールアドレス
 */
function runAutoMailSystem() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('送信対象データがありません。');
    return;
  }

  const sentRows = [];
  const skippedRows = [];
  const candidatePreparedRows = [];

  for (let row = 2; row <= lastRow; row++) {
    const company = String(sheet.getRange(row, COL_COMPANY).getValue() || '').trim();
    const name = String(sheet.getRange(row, COL_NAME).getValue() || '').trim();
    const status = String(sheet.getRange(row, COL_STATUS).getValue() || '').trim();
    let email = String(sheet.getRange(row, COL_EMAIL).getValue() || '').trim();

    if (status === STATUS_SENT) {
      continue;
    }

    if (!name) {
      skippedRows.push(row);
      sheet.getRange(row, COL_STATUS).setValue(STATUS_SKIP);
      continue;
    }

    if (!isValidEmail_(email)) {
      const candidates = writeCandidatesForRow_(sheet, row, false);
      if (candidates.length === 0) {
        skippedRows.push(row);
        sheet.getRange(row, COL_STATUS).setValue('候補なし');
        continue;
      }

      email = candidates[0];
      sheet.getRange(row, COL_EMAIL).setValue(email);
      candidatePreparedRows.push(row);

      if (!AUTO_SEND_FIRST_CANDIDATE) {
        sheet.getRange(row, COL_STATUS).setValue(STATUS_REVIEW);
        continue;
      }
    }

    GmailApp.sendEmail(email, buildSubject_(company, name), buildBody_(company, name));
    sheet.getRange(row, COL_STATUS).setValue(STATUS_SENT);
    sentRows.push(row);
  }

  SpreadsheetApp.getUi().alert(
    [
      '自動送信を完了しました。',
      '送信: ' + sentRows.length + '件',
      '候補をC列へ仮入力: ' + candidatePreparedRows.length + '件',
      'スキップ: ' + skippedRows.length + '件'
    ].join('\n')
  );
}

/**
 * 選択行の候補メールアドレスを生成
 */
function generateEmailCandidatesForActiveRow() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveCell().getRow();
  const emails = writeCandidatesForRow_(sheet, row);
  if (emails.length > 0 && !sheet.getRange(row, COL_EMAIL).getValue()) {
    sheet.getRange(row, COL_EMAIL).setValue(emails[0]);
  }
}

/**
 * 全行の候補メールアドレスを生成
 */
function generateEmailCandidatesForAllRows() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('シートにデータがありません。');
    return;
  }

  let count = 0;
  for (let row = 2; row <= lastRow; row++) {
    const domainInput = sheet.getRange(row, COL_DOMAIN).getValue();
    const name = sheet.getRange(row, COL_NAME).getValue();
    if (domainInput && name) {
      const emails = writeCandidatesForRow_(sheet, row, false);
      if (emails.length > 0 && !sheet.getRange(row, COL_EMAIL).getValue()) {
        sheet.getRange(row, COL_EMAIL).setValue(emails[0]);
      }
      count++;
    }
  }

  SpreadsheetApp.getUi().alert(count + ' 行に候補メールを生成しました。');
}

/**
 * 送信済みでない行だけ一括送信
 */
function sendPendingEmails() {
  runAutoMailSystem();
}

/**
 * 指定行に候補メールを横方向に書き出す
 */
function writeCandidatesForRow_(sheet, row, showAlert) {
  if (showAlert === undefined) showAlert = true;

  const domainInput = sheet.getRange(row, COL_DOMAIN).getValue();
  const name = sheet.getRange(row, COL_NAME).getValue();
  if (!domainInput || !name) {
    if (showAlert) {
      SpreadsheetApp.getUi().alert('この行の A列（ドメイン）と B列（名前）を入力してください。');
    }
    return [];
  }

  const emails = createEmailCandidates(name, domainInput);
  clearCandidateColumns_(sheet, row);

  if (emails.length > 0) {
    sheet.getRange(row, COL_CANDIDATE_START, 1, emails.length).setValues([emails]);
  }

  if (showAlert) {
    SpreadsheetApp.getUi().alert(emails.length + '件の候補メールアドレスを生成しました。');
  }
  return emails;
}

/**
 * 候補列(F列以降)をクリア
 */
function clearCandidateColumns_(sheet, row) {
  const maxColumns = sheet.getMaxColumns();
  const width = maxColumns - COL_CANDIDATE_START + 1;
  if (width > 0) {
    sheet.getRange(row, COL_CANDIDATE_START, 1, width).clearContent();
  }
}

/**
 * 件名
 */
function buildSubject_(company, name) {
  return '【ご相談のお願い】東京高専 吉本舜一';
}

/**
 * 本文
 */
function buildBody_(company, name) {
  const companyLabel = company || 'ご担当者';
  const nameLabel = name || '';

  return `${companyLabel} ${nameLabel}様

初めまして。東京高専2年の吉本舜一と申します。

現在、電気工学を学びながらビジネスコンテストやスタートアップに挑戦しております。将来は会社を経営し、社会に価値を生み出したいと本気で考えています。

ただ、高専で学ぶ技術をこのまま磨くべきか、大学へ進学し経営を学ぶべきか、進路について大きく迷っております。自分の選択が将来にどう繋がるのか、不安を感じることもあります。

${nameLabel || '皆様'}のご経歴を拝見し、ぜひ一度、進路選択やキャリアについて率直なお話を伺いたいと思いご連絡いたしました。

ご多忙のところ恐縮ですが、20〜30分ほどお時間をいただけませんでしょうか。オンライン・対面いずれでも可能です。

何卒よろしくお願いいたします。

吉本舜一`;
}

/**
 * 候補メールアドレス作成
 */
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
    patterns.push(`${first}.${last}`);
    patterns.push(`${first}${last}`);
    patterns.push(`${first}_${last}`);
    patterns.push(`${first}-${last}`);
    patterns.push(`${firstInitial}${last}`);
    patterns.push(`${firstInitial}.${last}`);
    patterns.push(`${firstInitial}_${last}`);
    patterns.push(`${firstInitial}-${last}`);
    patterns.push(`${first}${lastInitial}`);
    patterns.push(`${first}.${lastInitial}`);
    patterns.push(`${first}_${lastInitial}`);
    patterns.push(`${first}-${lastInitial}`);
    patterns.push(`${last}.${first}`);
    patterns.push(`${last}${first}`);
    patterns.push(`${last}_${first}`);
    patterns.push(`${last}-${first}`);
    patterns.push(`${last}${firstInitial}`);
    patterns.push(`${last}.${firstInitial}`);
    patterns.push(`${first}`);
    patterns.push(`${last}`);
    patterns.push(`${firstInitial}${lastInitial}`);
    patterns.push(`${lastInitial}${firstInitial}`);
  } else if (first) {
    patterns.push(first);
    patterns.push(firstInitial);
  }

  const emails = [];
  const seen = new Set();

  for (const pattern of patterns) {
    if (!pattern) continue;
    for (const domain of domains) {
      const email = `${pattern}@${domain}`;
      if (!seen.has(email)) {
        seen.add(email);
        emails.push(email);
      }
    }
  }

  return emails;
}

/**
 * 氏名をメール生成向けに正規化
 * 英字名を想定しつつ、日本語名が来ても空文字にはしない
 */
function normalizeNameParts_(name) {
  const raw = String(name || '').trim();
  if (!raw) return { first: '', last: '' };

  const cleaned = raw
    .replace(/[()（）]/g, ' ')
    .replace(/[,\u3001]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  const parts = cleaned.split(/[ 　]+/).filter(Boolean);
  const normalized = parts.map(part =>
    part
      .toLowerCase()
      .replace(/[^a-z0-9-]/g, '')
  ).filter(Boolean);

  if (normalized.length >= 2) {
    return { first: normalized[0], last: normalized[normalized.length - 1] };
  }
  if (normalized.length === 1) {
    return { first: normalized[0], last: '' };
  }

  const fallback = cleaned.toLowerCase().replace(/\s+/g, '').replace(/[^a-z0-9-]/g, '');
  return { first: fallback, last: '' };
}

/**
 * ドメイン抽出
 */
function extractDomains_(input) {
  const raw = String(input || '').trim().toLowerCase();
  if (!raw) return [];

  const tokens = raw.split(/[,\s;]+/).filter(Boolean);
  const domains = [];

  for (let token of tokens) {
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
  }

  return Array.from(new Set(domains));
}

/**
 * メールアドレス形式チェック
 */
function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || '').trim());
}

/**
 * カスタムメニュー
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('自動送信システム')
    .addItem('1. 選択行の候補メール生成', 'generateEmailCandidatesForActiveRow')
    .addItem('2. 全行の候補メール生成', 'generateEmailCandidatesForAllRows')
    .addSeparator()
    .addItem('3. 未送信メールを自動送信', 'sendPendingEmails')
    .addItem('4. 候補生成から送信まで一括実行', 'runAutoMailSystem')
    .addToUi();
}
