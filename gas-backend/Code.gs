/**
 * MeterOps – Cloud Vision OCR backend (Google Apps Script Web App)
 *
 * Deploy: Extensions > Apps Script > Deploy > New deployment > Web app
 *   Execute as: Me
 *   Who has access: Anyone   (the client is an anonymous browser page)
 *
 * Script Properties required (Project Settings > Script Properties):
 *   VISION_API_KEY  – Cloud Vision API key, restricted (API restrictions) to "Cloud Vision API" only
 *   SHARED_SECRET   – any random string; must match the `token` field the client sends
 *   SHEET_ID        – (optional) target spreadsheet ID; omit to use the script's bound spreadsheet
 *
 * After deploying, copy the Web App URL and the SHARED_SECRET value into
 * index.html's settings panel (gear icon, top right).
 */

const SHEET_NAME = 'Readings';
const VISION_ENDPOINT = 'https://vision.googleapis.com/v1/images:annotate';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const body = JSON.parse(e.postData.contents);

    // Lightweight auth so the public Web App URL can't be scraped/abused.
    const props = PropertiesService.getScriptProperties();
    const expectedToken = props.getProperty('SHARED_SECRET');
    if (expectedToken && body.token !== expectedToken) {
      return jsonResponse({ ok: false, error: 'unauthorized' });
    }

    const {
      imageBase64,     // data URL or raw base64 of the (cropped) digit-window image
      room,             // e.g. "101"
      meterType,        // "water" | "electric"
      expectedDigits    // optional int — helps pick the right block when known
    } = body;

    if (!imageBase64) {
      return jsonResponse({ ok: false, error: 'imageBase64 is required' });
    }

    const cleanBase64 = imageBase64.replace(/^data:image\/\w+;base64,/, '');

    const visionResult = callVisionOCR(cleanBase64);
    const picked = pickDigitReading(visionResult, expectedDigits);

    appendRow({
      timestamp: new Date(),
      room: room || '',
      meterType: meterType || '',
      reading: picked.text,
      confidence: picked.confidence,
      rawText: visionResult.fullText,
      angleDeg: picked.angleDeg,
      candidateCount: visionResult.candidates.length
    });

    return jsonResponse({
      ok: true,
      reading: picked.text,
      confidence: picked.confidence,
      angleDeg: picked.angleDeg,
      rawText: visionResult.fullText,
      candidates: visionResult.candidates
    });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('MeterOps Vision OCR endpoint is running.');
}

/** Calls Cloud Vision images:annotate with DOCUMENT_TEXT_DETECTION (gives per-word confidence). */
function callVisionOCR(base64Image) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('VISION_API_KEY');
  if (!apiKey) throw new Error('VISION_API_KEY script property is not set');

  const payload = {
    requests: [{
      image: { content: base64Image },
      features: [{ type: 'DOCUMENT_TEXT_DETECTION' }],
      imageContext: { languageHints: ['en'] } // digits are language-agnostic; avoids Vision guessing another script
    }]
  };

  const response = UrlFetchApp.fetch(VISION_ENDPOINT + '?key=' + apiKey, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const json = JSON.parse(response.getContentText());
  const annotation = json.responses && json.responses[0];
  if (!annotation || annotation.error) {
    throw new Error('Vision API error: ' + JSON.stringify(annotation && annotation.error));
  }

  const fullText = (annotation.fullTextAnnotation && annotation.fullTextAnnotation.text) || '';
  const candidates = [];
  const pages = (annotation.fullTextAnnotation && annotation.fullTextAnnotation.pages) || [];

  pages.forEach(function (page) {
    (page.blocks || []).forEach(function (block) {
      (block.paragraphs || []).forEach(function (paragraph) {
        const words = (paragraph.words || []).map(function (word) {
          const wordText = (word.symbols || []).map(function (s) { return s.text; }).join('');
          const confidences = (word.symbols || []).map(function (s) { return s.confidence || 0; });
          const avgConf = confidences.length
            ? confidences.reduce(function (a, b) { return a + b; }, 0) / confidences.length
            : (word.confidence || 0);
          return { text: wordText, confidence: avgConf, boundingBox: word.boundingBox };
        });

        // Individual words as candidates (each roller group Vision segments on its own).
        words.forEach(function (w) {
          candidates.push({
            text: w.text,
            confidence: w.confidence,
            angleDeg: boundingBoxAngle(w.boundingBox),
            vertices: w.boundingBox ? w.boundingBox.vertices : null
          });
        });

        // Whole paragraph as one candidate too (handles a reading Vision split into
        // multiple "words" because of a visible gap between roller digits).
        const paraText = words.map(function (w) { return w.text; }).join('');
        const paraConf = paragraph.confidence ||
          (words.reduce(function (a, w) { return a + w.confidence; }, 0) / (words.length || 1));
        candidates.push({
          text: paraText,
          confidence: paraConf,
          angleDeg: boundingBoxAngle(paragraph.boundingBox),
          vertices: paragraph.boundingBox ? paragraph.boundingBox.vertices : null
        });
      });
    });
  });

  return { fullText: fullText, candidates: candidates };
}

/**
 * Picks the best numeric candidate:
 *  1. Keep only candidates that contain at least one digit (strip everything else).
 *  2. Prefer an exact match to expectedDigits length, if given.
 *  3. Otherwise prefer the longest digit run.
 *  4. Break ties by Vision's own confidence score.
 */
function pickDigitReading(visionResult, expectedDigits) {
  const numeric = visionResult.candidates
    .map(function (c) {
      return Object.assign({}, c, { digitsOnly: (c.text || '').replace(/[^0-9]/g, '') });
    })
    .filter(function (c) { return c.digitsOnly.length > 0; });

  if (numeric.length === 0) {
    return { text: '', confidence: 0, angleDeg: 0 };
  }

  numeric.sort(function (a, b) {
    if (expectedDigits) {
      const aMatch = a.digitsOnly.length === expectedDigits ? 1 : 0;
      const bMatch = b.digitsOnly.length === expectedDigits ? 1 : 0;
      if (aMatch !== bMatch) return bMatch - aMatch;
    }
    if (b.digitsOnly.length !== a.digitsOnly.length) {
      return b.digitsOnly.length - a.digitsOnly.length;
    }
    return (b.confidence || 0) - (a.confidence || 0);
  });

  const best = numeric[0];
  return { text: best.digitsOnly, confidence: best.confidence || 0, angleDeg: best.angleDeg || 0 };
}

/**
 * Recovers the text block's own rotation angle (degrees) from its boundingBox vertices.
 * This is the "deskew equivalent": Vision already read the rotated text correctly,
 * this angle is only for diagnostics (flag a too-tilted capture) or, if scanning a
 * full uncropped photo, for telling apart the real digit strip from stray rotated text.
 */
function boundingBoxAngle(boundingBox) {
  if (!boundingBox || !boundingBox.vertices || boundingBox.vertices.length < 2) return 0;
  const v = boundingBox.vertices;
  const dx = (v[1].x || 0) - (v[0].x || 0);
  const dy = (v[1].y || 0) - (v[0].y || 0);
  return Math.round(Math.atan2(dy, dx) * 180 / Math.PI * 10) / 10;
}

function appendRow(row) {
  const sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  const ss = sheetId ? SpreadsheetApp.openById(sheetId) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Room', 'MeterType', 'Reading', 'Confidence', 'RawText', 'AngleDeg', 'CandidateCount']);
  }
  sheet.appendRow([
    row.timestamp, row.room, row.meterType, row.reading,
    row.confidence, row.rawText, row.angleDeg, row.candidateCount
  ]);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
