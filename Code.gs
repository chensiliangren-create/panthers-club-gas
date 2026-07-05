/**
 * Panthers Club Code.gs
 * PDF取込 + Gemini解析 + PlayerStats_Import正式反映 完全版
 */

const STATS_UPLOAD_FOLDER_NAME = 'Panthers_Stats_Upload';
const GEMINI_MODEL_NAME = 'gemini-2.5-flash';
const GEMINI_MAX_RETRIES = 3;
const GEMINI_RETRY_SLEEP_MS = 2000;

/**
 * Driveフォルダ内のPDFから、未処理または最新の試合PDFを1件選んで取り込む。
 */
function processLatestStatsPdf() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const folder = getStatsUploadFolder_();
  const games = getGameMasterRows_(ss)
    .filter(function(game) {
      return game.GameID && game.試合日;
    })
    .sort(function(a, b) {
      return getDateValue_(b.試合日) - getDateValue_(a.試合日);
    });

  if (games.length === 0) {
    throw new Error('GameMasterに有効な試合情報がありません。');
  }

  const processedGameIds = getProcessedGameIdSet_(ss);
  const candidates = [];

  games.forEach(function(game) {
    const file = findPdfFileForGame_(folder, game);

    if (file) {
      candidates.push({
        game: game,
        file: file,
        processed: processedGameIds.has(normalizeGameId_(game.GameID))
      });
    }
  });

  if (candidates.length === 0) {
    throw new Error('Panthers_Stats_Upload内にGameMasterと照合できるPDFが見つかりません。');
  }

  const unprocessed = candidates.filter(function(candidate) {
    return !candidate.processed;
  });

  const selected = unprocessed.length > 0 ? unprocessed[0] : candidates[0];

  if (!selected || !selected.game || !selected.file) {
    throw new Error('処理対象PDFを特定できませんでした。');
  }

  processStatsUpload(selected.file.getBlob(), selected.game.GameID);
}

/**
 * 指定GameIDに対応するPDFをDriveフォルダから探して取り込む。
 */
function processStatsPdfByGameId(gameId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const folder = getStatsUploadFolder_();
  const targetGameId = normalizeGameId_(gameId);

  if (!targetGameId) {
    throw new Error('GameIDが空です。');
  }

  const game = getGameByGameId_(ss, targetGameId);

  if (!game) {
    throw new Error('GameMasterにGameIDが見つかりません: ' + targetGameId);
  }

  const file = findPdfFileForGame_(folder, game);

  if (!file) {
    throw new Error('GameIDに対応するPDFが見つかりません: ' + targetGameId);
  }

  processStatsUpload(file.getBlob(), targetGameId);
}

/**
 * PDFスタッツをGeminiで解析し、PlayerStats_Importへ書き込む。
 */
function processStatsUpload(pdfBlob, gameIdRaw) {
  const gameId = normalizeGameId_(gameIdRaw);
  const geminiResponse = analyzeStatsWithGemini(pdfBlob);
  const statsData = parseGeminiResponse(geminiResponse);

  if (!statsData || !Array.isArray(statsData)) {
    Logger.log('処理中断: 解析データを配列形式で取得できませんでした。');
    return;
  }

  writeToSheet(statsData, gameId);
}

/**
 * Geminiの解析結果をPlayerStats_Importへ書き込む。
 * PlayerStats_Import.PlayerNo にはPDF由来の背番号を入れる。
 */
function writeToSheet(statsData, gameId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getRequiredSheet_(ss, 'PlayerStats_Import');

  if (!statsData || !Array.isArray(statsData) || statsData.length === 0) {
    Logger.log('書き込み対象データがありません。');
    return;
  }

  const headers = getHeaders_(sheet);
  const seasonId = getSeasonIdForGame_(ss, gameId) || getActiveSeasonId_(ss);

  const rows = statsData
    .filter(function(p) {
      return !isTotalsRow_(p);
    })
    .map(function(p) {
      const rowObj = {};

      rowObj.PlayerStats = '';
      rowObj.SeasonID = seasonId;
      rowObj.GameID = normalizeGameId_(gameId);
      rowObj.PlayerNo = p.PlayerNo || p.PlayerID || p.No || p['No.'] || '';
      rowObj.GS = p.GS || 0;
      rowObj.PTS = p.PTS || 0;
      rowObj['eFG%'] = p['eFG%'] || 0;
      rowObj['3P/M'] = p['3P/M'] || 0;
      rowObj['3P/A'] = p['3P/A'] || 0;
      rowObj['3P%'] = p['3P%'] || 0;
      rowObj['2P/M'] = p['2P/M'] || 0;
      rowObj['2P/A'] = p['2P/A'] || 0;
      rowObj['2P%'] = p['2P%'] || 0;
      rowObj['FT/M'] = p['FT/M'] || 0;
      rowObj['FT/A'] = p['FT/A'] || 0;
      rowObj['FT%'] = p['FT%'] || 0;
      rowObj.OREB = p.OREB || 0;
      rowObj.DREB = p.DREB || 0;
      rowObj.TOTREB = p.TOTREB || 0;
      rowObj.AST = p.AST || 0;
      rowObj.STL = p.STL || 0;
      rowObj.BLK = p.BLK || 0;
      rowObj.TO = p.TO || 0;
      rowObj.PF = p.PF || 0;
      rowObj.TF = p.TF || 0;
      rowObj.OF = p.OF || 0;
      rowObj.MIN = p.MIN || 0;
      rowObj['+/-'] = p['+/-'] || p['±'] || p['ﾂｱ'] || 0;
      rowObj.PPP = p.PPP || 0;
      rowObj['TO%'] = p['TO%'] || 0;
      rowObj['OREB%'] = p['OREB%'] || 0;
      rowObj.FTR = p.FTR || 0;
      rowObj['確認ステータス'] = '未確認';
      rowObj['確認者'] = '';
      rowObj['確認日時'] = '';
      rowObj['取込済み'] = false;
      rowObj['承認'] = false;
      rowObj.ImportBatchID = '';
      rowObj['反映日時'] = '';
      rowObj['反映エラー'] = '';

      return headers.map(function(header) {
        return Object.prototype.hasOwnProperty.call(rowObj, header) ? rowObj[header] : '';
      });
    });

  if (rows.length === 0) {
    Logger.log('TOTALS行以外の書き込み対象データがありません。');
    return;
  }

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  Logger.log('成功: ' + rows.length + '件のデータをPlayerStats_Importへ書き込みました。');
}

/**
 * Gemini APIでPDFを解析する。
 */
function analyzeStatsWithGemini(pdfBlob) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

  if (!apiKey) {
    throw new Error('GEMINI_API_KEY がスクリプトプロパティに設定されていません。');
  }

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL_NAME + ':generateContent?key=' + apiKey;
  const base64Data = Utilities.base64Encode(pdfBlob.getBytes());

  const payload = {
    contents: [{
      parts: [
        {
          text: 'PDFの統計表から、個人スタッツだけをJSON配列として抽出してください。説明文は不要です。PlayerNoは背番号です。PlayerIDではありません。TOTALS行、TOTAL行、合計行、チーム合計行は除外してください。数値は数値として返してください。キー: PlayerNo, GS, PTS, eFG%, 3P/M, 3P/A, 3P%, 2P/M, 2P/A, 2P%, FT/M, FT/A, FT%, OREB, DREB, TOTREB, AST, STL, BLK, TO, PF, TF, OF, MIN, +/-, PPP, TO%, OREB%, FTR'
        },
        {
          inline_data: {
            mime_type: 'application/pdf',
            data: base64Data
          }
        }
      ]
    }]
  };

  let lastError = null;

  for (let attempt = 1; attempt <= GEMINI_MAX_RETRIES; attempt++) {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const statusCode = res.getResponseCode();
    const responseText = res.getContentText();

    if (statusCode >= 200 && statusCode < 300) {
      const json = JSON.parse(responseText);

      if (
        !json.candidates ||
        !json.candidates[0] ||
        !json.candidates[0].content ||
        !json.candidates[0].content.parts ||
        !json.candidates[0].content.parts[0]
      ) {
        throw new Error('Gemini APIの応答形式が想定外です: ' + responseText);
      }

      return json.candidates[0].content.parts[0].text;
    }

    lastError = new Error('Gemini APIエラー: HTTP ' + statusCode + ' / ' + responseText);

    if (statusCode === 503 || statusCode === 429 || statusCode >= 500) {
      if (attempt < GEMINI_MAX_RETRIES) {
        Utilities.sleep(GEMINI_RETRY_SLEEP_MS * attempt);
        continue;
      }
    }

    throw lastError;
  }

  throw lastError || new Error('Gemini APIエラー: 不明なエラー');
}

/**
 * Geminiのテキスト応答からJSON配列を取り出す。
 */
function parseGeminiResponse(rawText) {
  Logger.log(rawText);

  try {
    let text = String(rawText || '').trim();

    const codeBlockMatch = text.match(/```(?:json)?\s*([\s\S]*?)```/i);
    if (codeBlockMatch && codeBlockMatch[1]) {
      text = codeBlockMatch[1].trim();
    } else {
      const start = text.indexOf('[');
      const end = text.lastIndexOf(']');
      if (start >= 0 && end > start) {
        text = text.substring(start, end + 1);
      }
    }

    const parsed = JSON.parse(text);

    if (!Array.isArray(parsed)) {
      return null;
    }

    return parsed.filter(function(row) {
      return !isTotalsRow_(row);
    });
  } catch (e) {
    Logger.log('Gemini応答のJSON解析に失敗しました: ' + e.message);
    return null;
  }
}

/**
 * テスト用。
 */
function testProcess() {
  processStatsPdfByGameId('GAME0613');
}

/**
 * PlayerStats_ImportからPlayerStatsへ正式反映し、反映したGameIDのTeamGameSummaryを再集計する。
 */
function approveAndCommitPlayerStatsImport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const importSheet = getRequiredSheet_(ss, 'PlayerStats_Import');
  const statsSheet = getRequiredSheet_(ss, 'PlayerStats');
  const playerSheet = getRequiredSheet_(ss, 'PlayerMaster');
  const gameSheet = getRequiredSheet_(ss, 'GameMaster');
  const seasonSheet = getRequiredSheet_(ss, 'SeasonMaster');
  const batchSheet = getRequiredSheet_(ss, 'ImportBatch');
  const errorSheet = getRequiredSheet_(ss, 'ImportErrorLog');

  const batchId = createId_('BATCH');

  appendByHeader_(batchSheet, {
    ImportBatchID: batchId,
    処理種別: 'PlayerStats正式反映',
    開始日時: new Date(),
    終了日時: '',
    ステータス: '処理中',
    対象件数: 0,
    成功件数: 0,
    エラー件数: 0,
    メモ: ''
  });

  let targetCount = 0;
  let successCount = 0;
  let errorCount = 0;
  const affectedGameIds = new Set();

  try {
    const importData = readSheetObjects_(importSheet);
    const statsData = readSheetObjects_(statsSheet);

    const playerNoToPlayerId = buildPlayerNoToPlayerIdMap_(playerSheet);
    const validGameIds = buildIdSet_(gameSheet, 'GameID');
    const validSeasonIds = buildIdSet_(seasonSheet, 'SeasonID');
    const existingStatsKeys = buildExistingPlayerStatsKeys_(statsData);

    const importHeaderMap = getHeaderMap_(importSheet);
    const statsHeaders = getHeaders_(statsSheet);

    importData.rows.forEach(function(rowObj, index) {
      const rowNumber = importData.startRow + index;

      if (!isCommitTarget_(rowObj)) {
        return;
      }

      targetCount++;

      const seasonId = normalizeText_(rowObj.SeasonID);
      const gameId = normalizeGameId_(rowObj.GameID);
      const playerNo = normalizePlayerNo_(rowObj.PlayerNo);
      let resolvedPlayerId = '';

      try {
        resolvedPlayerId = resolvePlayerIdByPlayerNo_(playerNo, playerNoToPlayerId);

        validateImportRow_(rowObj, {
          gameId: gameId,
          playerNo: playerNo,
          resolvedPlayerId: resolvedPlayerId,
          validGameIds: validGameIds,
          validSeasonIds: validSeasonIds
        });

        const uniqueKey = buildPlayerStatsUniqueKey_(seasonId, gameId, resolvedPlayerId);

        if (existingStatsKeys.has(uniqueKey)) {
          throw createImportError_(
            'DUPLICATE_PLAYER_STATS',
            '同じ SeasonID + GameID + PlayerID の PlayerStats がすでに存在します。'
          );
        }

        const statsRow = buildPlayerStatsRow_(rowObj, statsHeaders, resolvedPlayerId);
        appendByHeader_(statsSheet, statsRow);

        existingStatsKeys.add(uniqueKey);

        updateImportRowSuccess_(importSheet, importHeaderMap, rowNumber, {
          ImportBatchID: batchId,
          取込済み: true,
          反映日時: new Date(),
          反映エラー: ''
        });

        affectedGameIds.add(gameId);
        successCount++;
      } catch (e) {
        errorCount++;

        const errorCode = e.errorCode || 'UNKNOWN_ERROR';
        const errorMessage = e.message || String(e);

        appendByHeader_(errorSheet, {
          ErrorID: createId_('ERR'),
          ImportBatchID: batchId,
          発生日時: new Date(),
          シート名: 'PlayerStats_Import',
          行番号: rowNumber,
          ErrorCode: errorCode,
          ErrorMessage: errorMessage,
          SeasonID: seasonId,
          GameID: gameId,
          PlayerID: resolvedPlayerId,
          PlayerNo: playerNo,
          RawData: JSON.stringify(rowObj)
        });

        updateImportRowError_(importSheet, importHeaderMap, rowNumber, {
          ImportBatchID: batchId,
          反映エラー: '[' + errorCode + '] ' + errorMessage
        });
      }
    });

    affectedGameIds.forEach(function(gameId) {
      recalcTeamGameSummary(gameId);
    });

    updateImportBatch_(batchSheet, batchId, {
      終了日時: new Date(),
      ステータス: errorCount > 0 ? '一部エラー' : '完了',
      対象件数: targetCount,
      成功件数: successCount,
      エラー件数: errorCount,
      メモ: affectedGameIds.size > 0 ? '再集計GameID: ' + Array.from(affectedGameIds).join(', ') : ''
    });

    Logger.log(
      'PlayerStats正式反映 完了: 対象=' + targetCount +
      ', 成功=' + successCount +
      ', エラー=' + errorCount
    );
  } catch (e) {
    updateImportBatch_(batchSheet, batchId, {
      終了日時: new Date(),
      ステータス: '失敗',
      対象件数: targetCount,
      成功件数: successCount,
      エラー件数: errorCount + 1,
      メモ: e.message || String(e)
    });

    appendByHeader_(errorSheet, {
      ErrorID: createId_('ERR'),
      ImportBatchID: batchId,
      発生日時: new Date(),
      シート名: '',
      行番号: '',
      ErrorCode: 'BATCH_FAILED',
      ErrorMessage: e.message || String(e),
      SeasonID: '',
      GameID: '',
      PlayerID: '',
      PlayerNo: '',
      RawData: ''
    });

    throw e;
  }
}

/**
 * PlayerStatsからTeamGameSummaryを再集計する。
 */
function recalcTeamGameSummary(gameId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const statsSheet = getRequiredSheet_(ss, 'PlayerStats');
  const summarySheet = getRequiredSheet_(ss, 'TeamGameSummary');

  const targetGameId = normalizeGameId_(gameId);

  if (!targetGameId) {
    throw new Error('GameID が空です。');
  }

  const statsData = readSheetObjects_(statsSheet);
  const targetRows = statsData.rows.filter(function(row) {
    return normalizeGameId_(row.GameID) === targetGameId &&
      normalizeText_(row.PlayerID) !== '' &&
      !isTotalsRow_(row);
  });

  if (targetRows.length === 0) {
    throw new Error('指定GameIDのPlayerStatsが見つかりません: ' + targetGameId);
  }

  const seasonId = normalizeText_(targetRows[0].SeasonID);

  if (!seasonId) {
    throw new Error('対象PlayerStatsのSeasonIDが空です: ' + targetGameId);
  }

  const totals = {
    PTS: 0,
    '3P/M': 0,
    '3P/A': 0,
    '2P/M': 0,
    '2P/A': 0,
    'FT/M': 0,
    'FT/A': 0,
    OREB: 0,
    DREB: 0,
    TOTREB: 0,
    AST: 0,
    TO: 0,
    STL: 0,
    BLK: 0,
    PF: 0,
    FGA: 0,
    FGM: 0
  };

  targetRows.forEach(function(row) {
    totals.PTS += toNumber_(row.PTS);
    totals['3P/M'] += toNumber_(row['3P/M']);
    totals['3P/A'] += toNumber_(row['3P/A']);
    totals['2P/M'] += toNumber_(row['2P/M']);
    totals['2P/A'] += toNumber_(row['2P/A']);
    totals['FT/M'] += toNumber_(row['FT/M']);
    totals['FT/A'] += toNumber_(row['FT/A']);
    totals.OREB += toNumber_(row.OREB);
    totals.DREB += toNumber_(row.DREB);
    totals.TOTREB += toNumber_(row.TOTREB);
    totals.AST += toNumber_(row.AST);
    totals.TO += toNumber_(row.TO);
    totals.STL += toNumber_(row.STL);
    totals.BLK += toNumber_(row.BLK);
    totals.PF += toNumber_(row.PF);
  });

  totals.FGA = totals['3P/A'] + totals['2P/A'];
  totals.FGM = totals['3P/M'] + totals['2P/M'];

  const poss = totals.FGA + 0.44 * totals['FT/A'] + totals.TO - totals.OREB;

  const summaryValues = {
    TeamGameSummaryID: 'TGS_' + seasonId + '_' + targetGameId,
    SeasonID: seasonId,
    GameID: targetGameId,
    PTS: round1_(totals.PTS),
    '3P/M': round1_(totals['3P/M']),
    '3P/A': round1_(totals['3P/A']),
    '3P%': safePercent_(totals['3P/M'], totals['3P/A']),
    '2P/M': round1_(totals['2P/M']),
    '2P/A': round1_(totals['2P/A']),
    '2P%': safePercent_(totals['2P/M'], totals['2P/A']),
    'FT/M': round1_(totals['FT/M']),
    'FT/A': round1_(totals['FT/A']),
    'FT%': safePercent_(totals['FT/M'], totals['FT/A']),
    'eFG%': safeDivide_(totals.FGM + 0.5 * totals['3P/M'], totals.FGA, 100),
    OREB: round1_(totals.OREB),
    DREB: round1_(totals.DREB),
    TOTREB: round1_(totals.TOTREB),
    AST: round1_(totals.AST),
    TO: round1_(totals.TO),
    STL: round1_(totals.STL),
    BLK: round1_(totals.BLK),
    PF: round1_(totals.PF),
    FGA: round1_(totals.FGA),
    FGM: round1_(totals.FGM),
    Poss: round1_(poss),
    PPP: safeDivide_(totals.PTS, poss, 1),
    'TO%': safeDivide_(totals.TO, poss, 100),
    FTR: safeDivide_(totals['FT/A'], totals.FGA, 100),
    'AST/TO': safeDivide_(totals.AST, totals.TO, 1)
  };

  const summaryData = readSheetObjects_(summarySheet);
  const headerMapAll = getHeaderMapAll_(summarySheet);
  let updated = false;

  for (let i = 0; i < summaryData.rows.length; i++) {
    if (normalizeGameId_(summaryData.rows[i].GameID) === targetGameId) {
      const rowNumber = summaryData.startRow + i;
      setRowValuesByHeaderAll_(summarySheet, headerMapAll, rowNumber, summaryValues);
      updated = true;
      break;
    }
  }

  if (!updated) {
    appendByHeader_(summarySheet, summaryValues);
  }

  Logger.log('TeamGameSummary再集計 完了: ' + targetGameId);
}

/**
 * 正式反映対象行かどうかを判定する。
 */
function isCommitTarget_(rowObj) {
  if (isTotalsRow_(rowObj)) {
    return false;
  }

  const imported = toBoolean_(rowObj['取込済み']);

  if (imported) {
    return false;
  }

  const approved = toBoolean_(rowObj['承認']);
  const status = normalizeText_(rowObj['確認ステータス']);

  return approved || status === '確認済み';
}

/**
 * TOTALS行、合計行、チーム合計行を判定する。
 */
function isTotalsRow_(rowObj) {
  const values = [
    rowObj.PlayerNo,
    rowObj.PlayerID,
    rowObj.No,
    rowObj['No.'],
    rowObj.Name,
    rowObj.PlayerName,
    rowObj.DisplayName,
    rowObj['名前'],
    rowObj['選手名']
  ];

  for (let i = 0; i < values.length; i++) {
    const text = normalizeText_(values[i]).toUpperCase();

    if (
      text === 'TOTALS' ||
      text === 'TOTAL' ||
      text === 'TEAM TOTALS' ||
      text === 'TEAM TOTAL' ||
      text === '合計' ||
      text === 'チーム合計'
    ) {
      return true;
    }
  }

  return false;
}

/**
 * PlayerStats_Import行の整合性を検証する。
 */
function validateImportRow_(rowObj, context) {
  const seasonId = normalizeText_(rowObj.SeasonID);
  const gameId = context.gameId || normalizeGameId_(rowObj.GameID);
  const playerNo = context.playerNo || normalizePlayerNo_(rowObj.PlayerNo);

  if (!seasonId) {
    throw createImportError_('MISSING_SEASON_ID', 'SeasonID が空です。');
  }

  if (!/^\d{4}$/.test(seasonId)) {
    throw createImportError_('INVALID_SEASON_ID_FORMAT', 'SeasonID は 2026 のような西暦4桁である必要があります: ' + seasonId);
  }

  if (!gameId) {
    throw createImportError_('MISSING_GAME_ID', 'GameID が空です。');
  }

  if (!/^GAME\d{4}$/.test(gameId)) {
    throw createImportError_('INVALID_GAME_ID_FORMAT', 'GameID は GAME0613 のような形式である必要があります: ' + gameId);
  }

  if (!playerNo) {
    throw createImportError_('MISSING_PLAYER_NO', 'PlayerNo が空です。');
  }

  if (!context.validSeasonIds.has(seasonId)) {
    throw createImportError_('INVALID_SEASON_ID', 'SeasonMaster に存在しない SeasonID です: ' + seasonId);
  }

  if (!context.validGameIds.has(gameId)) {
    throw createImportError_('INVALID_GAME_ID', 'GameMaster に存在しない GameID です: ' + gameId);
  }

  if (!context.resolvedPlayerId) {
    throw createImportError_(
      'INVALID_PLAYER_NO',
      'PlayerMaster.No. に一致しない背番号です: ' + playerNo
    );
  }

  if (!/^P\d{3}$/.test(context.resolvedPlayerId)) {
    throw createImportError_(
      'INVALID_PLAYER_ID_FORMAT',
      '解決されたPlayerIDが P001 形式ではありません: ' + context.resolvedPlayerId
    );
  }
}

/**
 * PlayerStatsへ書き込む1行分のオブジェクトを作る。
 */
function buildPlayerStatsRow_(importRow, statsHeaders, resolvedPlayerId) {
  const row = {};

  statsHeaders.forEach(function(header) {
    if (!header) {
      return;
    }

    if (header === 'PlayerStatsID' || header === 'PlayerStats') {
      row[header] = createId_('PSTAT');
      return;
    }

    if (header === 'SeasonID') {
      row[header] = normalizeText_(importRow.SeasonID);
      return;
    }

    if (header === 'GameID') {
      row[header] = normalizeGameId_(importRow.GameID);
      return;
    }

    if (header === 'PlayerID') {
      row[header] = resolvedPlayerId;
      return;
    }

    if (header === 'PlayerNo') {
      return;
    }

    if (
      header === '確認ステータス' ||
      header === '確認者' ||
      header === '確認日時' ||
      header === '承認' ||
      header === '取込済み' ||
      header === 'ImportBatchID' ||
      header === '反映日時' ||
      header === '反映エラー'
    ) {
      return;
    }

    if (Object.prototype.hasOwnProperty.call(importRow, header)) {
      row[header] = importRow[header];
    }
  });

  return row;
}

/**
 * PlayerMaster.No. から PlayerMaster.PlayerID を引くMapを作る。
 */
function buildPlayerNoToPlayerIdMap_(playerSheet) {
  const data = readSheetObjects_(playerSheet);
  const map = {};

  data.rows.forEach(function(row) {
    const playerId = normalizeText_(row.PlayerID);
    const playerNo = normalizePlayerNo_(row['No.'] || row.No || row['背番号']);

    if (playerId && playerNo) {
      map[playerNo] = playerId;
    }
  });

  return map;
}

/**
 * 背番号から正式PlayerIDを取得する。
 */
function resolvePlayerIdByPlayerNo_(playerNo, playerNoToPlayerId) {
  const normalizedPlayerNo = normalizePlayerNo_(playerNo);
  return playerNoToPlayerId[normalizedPlayerNo] || '';
}

/**
 * 背番号を比較用に正規化する。
 */
function normalizePlayerNo_(value) {
  if (value === null || value === undefined) {
    return '';
  }

  const text = String(value).trim();

  if (!text) {
    return '';
  }

  const numberValue = Number(text);

  if (!isNaN(numberValue)) {
    return String(numberValue);
  }

  return text;
}

/**
 * 既存PlayerStatsの重複判定キーを作成する。
 */
function buildExistingPlayerStatsKeys_(statsData) {
  const set = new Set();

  statsData.rows.forEach(function(row) {
    const seasonId = normalizeText_(row.SeasonID);
    const gameId = normalizeGameId_(row.GameID);
    const playerId = normalizeText_(row.PlayerID);

    if (seasonId && gameId && playerId) {
      set.add(buildPlayerStatsUniqueKey_(seasonId, gameId, playerId));
    }
  });

  return set;
}

/**
 * PlayerStatsの一意キーを作成する。
 */
function buildPlayerStatsUniqueKey_(seasonId, gameId, playerId) {
  return [
    normalizeText_(seasonId),
    normalizeGameId_(gameId),
    normalizeText_(playerId)
  ].join('|');
}

/**
 * 成功時にPlayerStats_Importを更新する。
 */
function updateImportRowSuccess_(sheet, headerMap, rowNumber, values) {
  setRowValuesByHeader_(sheet, headerMap, rowNumber, values);
}

/**
 * エラー時にPlayerStats_Importを更新する。
 */
function updateImportRowError_(sheet, headerMap, rowNumber, values) {
  setRowValuesByHeader_(sheet, headerMap, rowNumber, values);
}

/**
 * ImportBatchの該当行を更新する。
 */
function updateImportBatch_(batchSheet, batchId, values) {
  const data = readSheetObjects_(batchSheet);
  const headerMap = getHeaderMap_(batchSheet);

  for (let i = 0; i < data.rows.length; i++) {
    if (normalizeText_(data.rows[i].ImportBatchID) === normalizeText_(batchId)) {
      const rowNumber = data.startRow + i;
      setRowValuesByHeader_(batchSheet, headerMap, rowNumber, values);
      return;
    }
  }

  throw new Error('ImportBatch が見つかりません: ' + batchId);
}

/**
 * 指定シートの指定ID列からSetを作る。
 */
function buildIdSet_(sheet, idHeader) {
  const data = readSheetObjects_(sheet);
  const set = new Set();

  data.rows.forEach(function(row) {
    const id = idHeader === 'GameID'
      ? normalizeGameId_(row[idHeader])
      : normalizeText_(row[idHeader]);

    if (id) {
      set.add(id);
    }
  });

  return set;
}

/**
 * ヘッダー名に基づいて1行追記する。
 */
function appendByHeader_(sheet, values) {
  const headers = getHeaders_(sheet);
  const row = headers.map(function(header) {
    return Object.prototype.hasOwnProperty.call(values, header) ? values[header] : '';
  });

  sheet.appendRow(row);
}

/**
 * ヘッダー名に基づいて指定行の複数列を更新する。
 */
function setRowValuesByHeader_(sheet, headerMap, rowNumber, values) {
  Object.keys(values).forEach(function(header) {
    if (!headerMap[header]) {
      return;
    }

    sheet.getRange(rowNumber, headerMap[header]).setValue(values[header]);
  });
}

/**
 * 重複ヘッダーを含めて、ヘッダー名 -> 列番号配列 のMapを取得する。
 */
function getHeaderMapAll_(sheet) {
  const headers = getHeaders_(sheet);
  const map = {};

  headers.forEach(function(header, index) {
    if (!header) {
      return;
    }

    if (!map[header]) {
      map[header] = [];
    }

    map[header].push(index + 1);
  });

  return map;
}

/**
 * 重複ヘッダーを含めて、指定行の複数列を更新する。
 */
function setRowValuesByHeaderAll_(sheet, headerMapAll, rowNumber, values) {
  Object.keys(values).forEach(function(header) {
    const columns = headerMapAll[header];

    if (!columns || columns.length === 0) {
      return;
    }

    columns.forEach(function(column) {
      sheet.getRange(rowNumber, column).setValue(values[header]);
    });
  });
}

/**
 * シートをオブジェクト配列として読む。
 */
function readSheetObjects_(sheet) {
  const values = sheet.getDataRange().getValues();

  if (values.length < 1) {
    return {
      headers: [],
      rows: [],
      startRow: 2
    };
  }

  const headers = values[0].map(function(value) {
    return normalizeText_(value);
  });

  const rows = [];

  for (let r = 1; r < values.length; r++) {
    const obj = {};

    headers.forEach(function(header, c) {
      if (header) {
        obj[header] = values[r][c];
      }
    });

    rows.push(obj);
  }

  return {
    headers: headers,
    rows: rows,
    startRow: 2
  };
}

/**
 * ヘッダー一覧を取得する。
 */
function getHeaders_(sheet) {
  const lastColumn = sheet.getLastColumn();

  if (lastColumn === 0) {
    return [];
  }

  return sheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0]
    .map(function(value) {
      return normalizeText_(value);
    });
}

/**
 * ヘッダー名 -> 列番号 のMapを取得する。
 */
function getHeaderMap_(sheet) {
  const headers = getHeaders_(sheet);
  const map = {};

  headers.forEach(function(header, index) {
    if (header) {
      map[header] = index + 1;
    }
  });

  return map;
}

/**
 * 必須シートを取得する。
 */
function getRequiredSheet_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('必要なシートが見つかりません: ' + sheetName);
  }

  return sheet;
}

/**
 * ActiveなSeasonIDを取得する。
 */
function getActiveSeasonId_(ss) {
  const sheet = ss.getSheetByName('SeasonMaster');

  if (!sheet) {
    return '';
  }

  const data = readSheetObjects_(sheet);

  for (let i = 0; i < data.rows.length; i++) {
    if (toBoolean_(data.rows[i].Active)) {
      return normalizeText_(data.rows[i].SeasonID);
    }
  }

  if (data.rows.length > 0) {
    return normalizeText_(data.rows[0].SeasonID);
  }

  return '';
}

/**
 * GameIDに対応するSeasonIDをGameMasterから取得する。
 */
function getSeasonIdForGame_(ss, gameId) {
  const game = getGameByGameId_(ss, gameId);
  return game ? normalizeText_(game.SeasonID) : '';
}

/**
 * GameMasterからGameID指定で1行取得する。
 */
function getGameByGameId_(ss, gameId) {
  const targetGameId = normalizeGameId_(gameId);
  const games = getGameMasterRows_(ss);

  for (let i = 0; i < games.length; i++) {
    if (normalizeGameId_(games[i].GameID) === targetGameId) {
      return games[i];
    }
  }

  return null;
}

/**
 * GameMasterを読み込む。
 */
function getGameMasterRows_(ss) {
  const sheet = getRequiredSheet_(ss, 'GameMaster');
  const data = readSheetObjects_(sheet);

  return data.rows.map(function(row) {
    row.GameID = normalizeGameId_(row.GameID);
    return row;
  });
}

/**
 * 処理済みGameIDを取得する。
 */
function getProcessedGameIdSet_(ss) {
  const set = new Set();

  const importSheet = ss.getSheetByName('PlayerStats_Import');
  if (importSheet) {
    const importData = readSheetObjects_(importSheet);
    importData.rows.forEach(function(row) {
      const gameId = normalizeGameId_(row.GameID);
      if (gameId) {
        set.add(gameId);
      }
    });
  }

  const statsSheet = ss.getSheetByName('PlayerStats');
  if (statsSheet) {
    const statsData = readSheetObjects_(statsSheet);
    statsData.rows.forEach(function(row) {
      const gameId = normalizeGameId_(row.GameID);
      if (gameId) {
        set.add(gameId);
      }
    });
  }

  return set;
}

/**
 * PDFアップロードフォルダを取得する。
 */
function getStatsUploadFolder_() {
  const folders = DriveApp.getFoldersByName(STATS_UPLOAD_FOLDER_NAME);

  if (!folders.hasNext()) {
    throw new Error('Driveフォルダが見つかりません: ' + STATS_UPLOAD_FOLDER_NAME);
  }

  return folders.next();
}

/**
 * GameMasterの試合情報に対応するPDFを探す。
 */
function findPdfFileForGame_(folder, game) {
  const files = folder.getFiles();
  const gameId = normalizeGameId_(game.GameID);
  const ymd = formatGameDateYmd_(game.試合日);
  const opponent = normalizeFileToken_(game.対戦相手);
  const displayGame = normalizeFileToken_(game.DisplayGame);

  let fallback = null;

  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    const lower = name.toLowerCase();

    if (lower.indexOf('.pdf') === -1) {
      continue;
    }

    const normalizedName = normalizeFileToken_(name);

    if (gameId && normalizedName.indexOf(normalizeFileToken_(gameId)) >= 0) {
      return file;
    }

    if (ymd && normalizedName.indexOf(ymd) >= 0) {
      if (opponent && normalizedName.indexOf(opponent) >= 0) {
        return file;
      }

      if (displayGame && hasAnyDisplayGameToken_(normalizedName, displayGame)) {
        return file;
      }

      if (!fallback) {
        fallback = file;
      }
    }
  }

  return fallback;
}

/**
 * DisplayGame由来の一部トークンがファイル名に含まれるか判定する。
 */
function hasAnyDisplayGameToken_(normalizedName, displayGame) {
  const tokens = String(displayGame || '')
    .split(/[^0-9A-Z\u3040-\u30ff\u3400-\u9fff]+/i)
    .filter(function(token) {
      return token && token.length >= 2;
    });

  for (let i = 0; i < tokens.length; i++) {
    if (normalizedName.indexOf(tokens[i]) >= 0) {
      return true;
    }
  }

  return false;
}

/**
 * ファイル名照合用に文字列を正規化する。
 */
function normalizeFileToken_(value) {
  return normalizeText_(value)
    .toUpperCase()
    .replace(/\s+/g, '')
    .replace(/[._\-ー―‐/\\:：()（）\[\]【】]/g, '');
}

/**
 * 試合日をyyyyMMddへ変換する。
 */
function formatGameDateYmd_(value) {
  if (!value) {
    return '';
  }

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyyMMdd');
  }

  const text = normalizeText_(value);

  if (/^\d{8}$/.test(text)) {
    return text;
  }

  const compact = text.replace(/[^\d]/g, '');

  if (compact.length >= 8) {
    return compact.substring(0, 8);
  }

  return '';
}

/**
 * 日付比較用の値を取得する。
 */
function getDateValue_(value) {
  if (!value) {
    return 0;
  }

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return value.getTime();
  }

  const ymd = formatGameDateYmd_(value);

  if (!ymd) {
    return 0;
  }

  const year = Number(ymd.substring(0, 4));
  const month = Number(ymd.substring(4, 6)) - 1;
  const day = Number(ymd.substring(6, 8));

  return new Date(year, month, day).getTime();
}

/**
 * TRUE判定。
 */
function toBoolean_(value) {
  if (value === true) {
    return true;
  }

  const text = normalizeText_(value).toUpperCase();

  return text === 'TRUE' ||
    text === '1' ||
    text === 'YES' ||
    text === 'Y' ||
    text === '承認' ||
    text === '済' ||
    text === '確認済み';
}

/**
 * 文字列正規化。
 */
function normalizeText_(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return String(value).trim();
}

/**
 * GameIDを正規化する。
 */
function normalizeGameId_(value) {
  const text = normalizeText_(value).toUpperCase();

  if (!text) {
    return '';
  }

  return text.replace(/^GANE/, 'GAME');
}

/**
 * 数値化する。
 */
function toNumber_(value) {
  if (value === null || value === undefined || value === '') {
    return 0;
  }

  const numberValue = Number(String(value).replace(/,/g, ''));

  if (isNaN(numberValue)) {
    return 0;
  }

  return numberValue;
}

/**
 * 小数1桁に丸める。
 */
function round1_(value) {
  const numberValue = toNumber_(value);
  return Math.round(numberValue * 10) / 10;
}

/**
 * 0除算を避けて比率を計算する。
 */
function safeDivide_(numerator, denominator, multiplier) {
  const den = toNumber_(denominator);

  if (den === 0) {
    return 0;
  }

  return round1_(toNumber_(numerator) / den * multiplier);
}

/**
 * 0除算を避けて百分率を計算する。
 */
function safePercent_(numerator, denominator) {
  return safeDivide_(numerator, denominator, 100);
}

/**
 * ID生成。
 */
function createId_(prefix) {
  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyyMMddHHmmss'
  );

  const random = Math.floor(Math.random() * 1000000)
    .toString()
    .padStart(6, '0');

  return prefix + '_' + timestamp + '_' + random;
}

/**
 * エラーコード付き例外を作る。
 */
function createImportError_(errorCode, message) {
  const error = new Error(message);
  error.errorCode = errorCode;
  return error;
}

/**
 * 既存互換テスト用。
 */
function testRecalcTeamGameSummary_GAME0603() {
  recalcTeamGameSummary('GAME0603');
}

/**
 * 既存互換テスト用。
 */
function testProcessStatsUpload_GAME0603() {
  processStatsPdfByGameId('GAME0603');
}

/**
 * 既存互換テスト用。
 */
function testProcessStatsUpload_GAME0524() {
  processStatsPdfByGameId('GAME0524');
}
function testProcessStatsUpload_GAME0704() {
  processStatsPdfByGameId('GAME0704');
}
