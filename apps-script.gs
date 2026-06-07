// PMP復習クイズ - Google Apps Script
// Google スプレッドシートで Extensions > Apps Script を開き、このコードを貼り付けてください

const SHEET_NAME = 'シート1';
const MASTERED_COL = 15; // 列O（選択肢6追加により1列シフト）
const WRITE_TOKEN = 'mVSNq3Pf-GSwhXvi1yL_nKxYfwvR1ocZ';

function doGet(e) {
  const action = e.parameter.action || '';
  let result;

  try {
    if (action === 'getQuestions') {
      result = getQuestions();
    } else if (action === 'recordAnswer') {
      if (e.parameter.token !== WRITE_TOKEN) return forbidden();
      result = recordAnswer(e.parameter.questionId, e.parameter.correct === 'true');
    } else if (action === 'resetAll') {
      if (e.parameter.token !== WRITE_TOKEN) return forbidden();
      result = resetAll();
    } else {
      result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function forbidden() {
  return ContentService
    .createTextOutput(JSON.stringify({ error: 'Forbidden' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

function ensureMasteredColumn(sheet) {
  const headerVal = sheet.getRange(1, MASTERED_COL).getValue();
  if (headerVal !== 'マスター済み') {
    sheet.getRange(1, MASTERED_COL).setValue('マスター済み');
  }
}

function getQuestions() {
  const sheet = getSheet();
  ensureMasteredColumn(sheet);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, MASTERED_COL).getValues();
  const questions = [];

  data.forEach((row) => {
    const questionText = row[3]; // D列: 問題文
    if (!questionText) return;

    // E〜J列: 選択肢1〜6（空欄は除外）
    const choices = [
      String(row[4] || ''), // E: 選択肢1
      String(row[5] || ''), // F: 選択肢2
      String(row[6] || ''), // G: 選択肢3
      String(row[7] || ''), // H: 選択肢4
      String(row[8] || ''), // I: 選択肢5
      String(row[9] || ''), // J: 選択肢6
    ].filter(c => c.trim() !== '');

    const correctAnswer = String(row[10] || ''); // K列: 正解（全文）

    // 「 / 」区切りで複数正解を分割
    const correctParts = correctAnswer.split(' / ').map(s => s.trim()).filter(Boolean);
    const correctIndices = correctParts
      .map(part => findBestMatch(part, choices))
      .filter(i => typeof i === 'number' && !isNaN(i) && i >= 0);

    if (choices.length < 2 || correctIndices.length === 0) return;

    const mastered = row[MASTERED_COL - 1]; // N列: マスター済み

    questions.push({
      id:             String(row[0]),  // A列: 問題番号
      difficulty:     String(row[1]),  // B列: 難易度
      domain:         String(row[2]),  // C列: PMBOK領域
      question:       String(row[3]),  // D列: 問題文
      choices,
      correctIndices,
      isMultiple:     correctIndices.length > 1,
      explanation:    String(row[12] || ''), // M列: 解説
      mastered:       mastered === true || mastered === 'TRUE',
    });
  });

  return questions;
}

function recordAnswer(questionId, correct) {
  if (!questionId) return { error: 'questionId is required' };

  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { error: 'No data' };

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === questionId) {
      if (correct) {
        sheet.getRange(i + 2, MASTERED_COL).setValue(true);
      }
      return { success: true };
    }
  }

  return { error: 'Question not found: ' + questionId };
}

function resetAll() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true };

  const numRows = lastRow - 1;
  const values = Array(numRows).fill([false]);
  sheet.getRange(2, MASTERED_COL, numRows, 1).setValues(values);

  return { success: true };
}

// 正解テキストと選択肢のあいまいマッチング（スプレッドシートの表記ゆれに対応）
function findBestMatch(correct, choices) {
  const exact = choices.indexOf(correct);
  if (exact !== -1) return exact;

  let bestIdx = -1;
  let bestScore = 0;

  for (let i = 0; i < choices.length; i++) {
    const score = bigramSimilarity(correct, choices[i]);
    if (score > bestScore) {
      bestScore = score;
      bestIdx = i;
    }
  }

  return bestScore >= 0.6 ? bestIdx : -1;
}

function bigramSimilarity(a, b) {
  a = a.replace(/\s/g, '');
  b = b.replace(/\s/g, '');
  if (!a || !b) return 0;

  const getBigrams = s => {
    const set = new Set();
    for (let i = 0; i < s.length - 1; i++) set.add(s.slice(i, i + 2));
    return set;
  };

  const bg1 = getBigrams(a);
  const bg2 = getBigrams(b);
  let intersection = 0;
  bg1.forEach(bg => { if (bg2.has(bg)) intersection++; });

  return (2 * intersection) / (bg1.size + bg2.size);
}
