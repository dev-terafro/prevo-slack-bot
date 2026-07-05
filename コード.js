/**
   * リマインダーBotプラットフォーム — メインスクリプト
   * バージョン: 1.1.0（2026-05-15 朝稽古Bot用に skip機能追加）
   */

  // ===== 定数 =====
  var TZ = 'Asia/Tokyo';

  /** _bots シートの列インデックス（0始まり） */
  var COL_BOT_NAME       = 0;  // A
  var COL_ENABLED        = 1;  // B
  var COL_CHANNEL_ID     = 2;  // C
  var COL_SCHEDULE       = 3;  // D
  var COL_SOURCE_SHEET   = 4;  // E
  var COL_ROTATION_UNIT  = 5;  // F
  var COL_TEMPLATE       = 6;  // G
  var COL_LAST_PERIOD    = 7;  // H
  var COL_LAST_RUN       = 8;  // I
  var COL_SKIP_OFFSET    = 9;  // J ← 新規追加

  /** _bots の読み込み列数（新J列を含めて10） */
  var BOTS_NUM_COLS      = 10; // ← 新規（旧コードでは 9 だった箇所をこれに置換）

  /** _state シートの列インデックス（0始まり） */
  var COL_STATE_SOURCE   = 0;
  var COL_STATE_PERIOD   = 1;

  var DAY_MAP = { '日': 0, '月': 1, '火': 2, '水': 3, '木': 4, '金': 5, '土': 6 };
  var FULLWIDTH_DIGITS = { '0':'０','1':'１','2':'２','3':'３','4':'４','5':'５','6':'６','7':'７','8':'８','9':'９' };
  var KANJI_NUMS = ['〇','一','二','三','四','五','六','七','八','九','十','十一','十二'];
  var DAY_KANJI = ['日','月','火','水','木','金','土'];


  // ===== メイン処理 =====

  function dispatcher() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var botsSheet = ss.getSheetByName('_bots');
    var stateSheet = ss.getSheetByName('_state');

    if (!botsSheet) { Logger.log('[ERROR] _bots シートが見つかりません'); return; }
    if (!stateSheet) { Logger.log('[ERROR] _state シートが見つかりません'); return; }

    var now = new Date();
    var nowDateStr = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
    var nowHour = parseInt(Utilities.formatDate(now, TZ, 'HH'), 10);
    var nowDayOfWeek = parseInt(Utilities.formatDate(now, TZ, 'u'), 10) % 7;

    Logger.log('[INFO] dispatcher 起動: ' + Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm') + ' (曜日=' + DAY_KANJI[nowDayOfWeek] + ' 時='
  + nowHour + ')');

    var lastRow = botsSheet.getLastRow();
    if (lastRow < 2) { Logger.log('[INFO] _bots シートにデータがありません'); return; }
    var botsData = botsSheet.getRange(2, 1, lastRow - 1, BOTS_NUM_COLS).getValues();

    var stateMap = loadStateMap_(stateSheet);

    for (var i = 0; i < botsData.length; i++) {
      var row = botsData[i];
      var rowNum = i + 2;
      try {
        processBotRow_(ss, botsSheet, stateSheet, row, rowNum, stateMap, now, nowDateStr, nowHour, nowDayOfWeek);
      } catch (e) {
        Logger.log('[ERROR] Bot行 ' + rowNum + ' (' + String(row[COL_BOT_NAME]) + ') でエラー: ' + e.message);
      }
    }

    Logger.log('[INFO] dispatcher 完了');
  }


  function processBotRow_(ss, botsSheet, stateSheet, row, rowNum, stateMap, now, nowDateStr, nowHour, nowDayOfWeek) {
    var botName     = String(row[COL_BOT_NAME]).trim();
    var enabled     = String(row[COL_ENABLED]).trim().toUpperCase();
    var channelId   = String(row[COL_CHANNEL_ID]).trim();
    var schedule    = String(row[COL_SCHEDULE]).trim();
    var sourceSheet = String(row[COL_SOURCE_SHEET]).trim();
    var rotUnit     = String(row[COL_ROTATION_UNIT]).trim().toLowerCase();
    var template    = String(row[COL_TEMPLATE]).trim();
    var lastRun     = String(row[COL_LAST_RUN]).trim();
    var skipOffset  = String(row[COL_SKIP_OFFSET]).trim();  // ← 新規

    if (!botName) return;

    if (enabled !== 'TRUE') {
      Logger.log('[SKIP] ' + botName + ': enabled=FALSE');
      return;
    }

    var scheduleParts = schedule.split(/\s+/);
    if (scheduleParts.length < 2) throw new Error('schedule の形式が不正: ' + schedule);
    var schedDay = scheduleParts[0];
    var schedTime = scheduleParts[1];
    var schedHour = parseInt(schedTime.split(':')[0], 10);

    if (!(schedDay in DAY_MAP)) throw new Error('schedule の曜日が不正: ' + schedDay);
    if (DAY_MAP[schedDay] !== nowDayOfWeek) return;
    if (schedHour !== nowHour) return;

    var guardKey = nowDateStr + ' ' + pad2_(nowHour);
    if (lastRun === guardKey) {
      Logger.log('[SKIP] ' + botName + ': 本日この時間帯は実行済み (last_run=' + lastRun + ')');
      return;
    }

    // --- スキップ日チェック（ローテも投稿も全部スキップ）← 新規追加 ---
    if (skipOffset !== '' && skipOffset.toLowerCase() !== 'none') {
      var offsetDays = parseInt(skipOffset, 10);
      if (!isNaN(offsetDays)) {
        var targetDate = new Date(now.getTime() + offsetDays * 24 * 60 * 60 * 1000);
        var targetDateStr = Utilities.formatDate(targetDate, TZ, 'yyyy-MM-dd');
        if (isSkipDate_(ss, targetDateStr)) {
          Logger.log('[SKIP] ' + botName + ': 朝稽古お休み日のためスキップ（対象日=' + targetDateStr + '）');
          return;  // ローテーションも投稿もしない
        }
      }
    }

    Logger.log('[RUN] ' + botName + ' を発火します');

    // ローテーション処理
    if (rotUnit !== 'none' && sourceSheet && sourceSheet.toLowerCase() !== 'none') {
      var rosterSheet = ss.getSheetByName(sourceSheet);
      if (!rosterSheet) throw new Error('source_sheet が見つからない: ' + sourceSheet);
      var currentPeriod = getCurrentPeriod_(now, rotUnit);
      var lastRotatedPeriod = stateMap[sourceSheet] || '';

      if (currentPeriod !== lastRotatedPeriod) {
        Logger.log('[ROTATE] ' + sourceSheet + ': ' + lastRotatedPeriod + ' → ' + currentPeriod);
        rotateRoster_(rosterSheet);
        updateStateSheet_(stateSheet, sourceSheet, currentPeriod);
        stateMap[sourceSheet] = currentPeriod;
      } else {
        Logger.log('[NO-ROTATE] ' + sourceSheet + ': 既に今期ローテーション済み (' + currentPeriod + ')');
      }
    }

    var rosterSheet = (sourceSheet && sourceSheet.toLowerCase() !== 'none') ? ss.getSheetByName(sourceSheet) : null;
    var message = resolveTemplate_(template, now, rosterSheet);

    postToSlack_(channelId, message);
    Logger.log('[POST] ' + botName + ' → ' + channelId + '\n' + message);

    botsSheet.getRange(rowNum, COL_LAST_RUN + 1).setValue(guardKey);
  }


  // ===== スキップ日チェック ← 新規追加 =====
  /**
   * _skip シートの A 列に指定日付（YYYY-MM-DD）があるかチェック
   * シートが存在しない場合は false（=スキップしない）
   */
  function isSkipDate_(ss, dateStr) {
    var skipSheet = ss.getSheetByName('_skip');
    if (!skipSheet) return false;
    var lastRow = skipSheet.getLastRow();
    if (lastRow < 2) return false;
    var values = skipSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < values.length; i++) {
      var v = values[i][0];
      var s;
      if (v instanceof Date) {
        s = Utilities.formatDate(v, TZ, 'yyyy-MM-dd');  // 日付セルの場合
      } else {
        s = String(v).trim();
      }
      if (s === dateStr) return true;
    }
    return false;
  }


  // ===== ローテーション処理 =====

  function rotateRoster_(rosterSheet) {
    var lastRow = rosterSheet.getLastRow();
    if (lastRow < 3) {
      Logger.log('[WARN] ロスターのデータ行が少なすぎてローテ不可: ' + rosterSheet.getName());
      return;
    }
    var numDataRows = lastRow - 1;
    var bValues = rosterSheet.getRange(2, 2, numDataRows, 1).getValues();
    // 全員が1つ下の担当へ進み、最後尾（末尾）が先頭に戻る：[a,b,c,d] → [d,a,b,c]
    var last = bValues[numDataRows - 1][0];
    var newValues = [[last]];
    for (var i = 0; i < numDataRows - 1; i++) newValues.push([bValues[i][0]]);
    rosterSheet.getRange(2, 2, numDataRows, 1).setValues(newValues);
  }


  // ===== 期間計算 =====

  function getCurrentPeriod_(now, rotUnit) {
    if (rotUnit === 'month') return Utilities.formatDate(now, TZ, 'yyyy-MM');
    if (rotUnit === 'week')  return getISOWeekString_(now);
    if (rotUnit === 'day')   return Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
    return '';
  }

  function getISOWeekString_(date) {
    var d = new Date(Date.UTC(
      parseInt(Utilities.formatDate(date, TZ, 'yyyy'), 10),
      parseInt(Utilities.formatDate(date, TZ, 'MM'), 10) - 1,
      parseInt(Utilities.formatDate(date, TZ, 'dd'), 10)
    ));
    var dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    var weekNum = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
    return d.getUTCFullYear() + '-W' + pad2_(weekNum);
  }


  // ===== テンプレート変数置換 =====

  function resolveTemplate_(template, now, rosterSheet) {
    var text = template;
    text = text.replace(/\\n/g, '\n');
    var tomorrow = new Date(now.getTime() + 24 * 60 * 60 * 1000);

    var yearStr  = Utilities.formatDate(now, TZ, 'yyyy');
    var monthNum = parseInt(Utilities.formatDate(now, TZ, 'M'), 10);
    var dayNum   = parseInt(Utilities.formatDate(now, TZ, 'd'), 10);
    var dowNum   = parseInt(Utilities.formatDate(now, TZ, 'u'), 10) % 7;

    text = text.replace(/\{\{年\}\}/g, yearStr);
    text = text.replace(/\{\{月\}\}/g, toFullWidth_(String(monthNum)));
    text = text.replace(/\{\{月_漢数字\}\}/g, monthNum <= 12 ? KANJI_NUMS[monthNum] : String(monthNum));
    text = text.replace(/\{\{日\}\}/g, String(dayNum));
    text = text.replace(/\{\{曜日\}\}/g, DAY_KANJI[dowNum]);

    var tMonthNum = parseInt(Utilities.formatDate(tomorrow, TZ, 'M'), 10);
    var tDayNum   = parseInt(Utilities.formatDate(tomorrow, TZ, 'd'), 10);
    var tDowNum   = parseInt(Utilities.formatDate(tomorrow, TZ, 'u'), 10) % 7;

    text = text.replace(/\{\{翌日:曜日\}\}/g, DAY_KANJI[tDowNum]);
    text = text.replace(/\{\{翌日:日\}\}/g, String(tDayNum));
    text = text.replace(/\{\{翌日:月\}\}/g, toFullWidth_(String(tMonthNum)));

    if (rosterSheet) {
      var lastRow = rosterSheet.getLastRow();

      if (text.indexOf('{{top}}') !== -1) {
        var topValue = lastRow >= 2 ? String(rosterSheet.getRange(2, 2).getValue()).trim() : '（未設定）';
        text = text.replace(/\{\{top\}\}/g, topValue);
      }

      if (text.indexOf('{{table}}') !== -1) {
        var tableText = '';
        if (lastRow >= 2) {
          var numDataRows = lastRow - 1;
          var tableValues = rosterSheet.getRange(2, 1, numDataRows, 2).getValues();
          var lines = [];
          for (var i = 0; i < tableValues.length; i++) {
            var a = String(tableValues[i][0]).trim();
            var b = String(tableValues[i][1]).trim();
            if (!a && !b) continue;
            lines.push((a ? a + '：' : '') + b);
          }
          tableText = lines.join('\n');
        }
        text = text.replace(/\{\{table\}\}/g, tableText);
      }

      text = text.replace(/\{\{cell:(\d+):(\d+)\}\}/g, function(match, r, c) {
        var rNum = parseInt(r, 10), cNum = parseInt(c, 10);
        if (rNum < 1 || cNum < 1) return match;
        return String(rosterSheet.getRange(rNum, cNum).getValue()).trim();
      });
    }

    return text;
  }


  // ===== Slack 投稿 =====

  function postToSlack_(channelId, text) {
    var token = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');
    if (!token) throw new Error('Script Properties に SLACK_BOT_TOKEN が未設定');
    if (!channelId) throw new Error('channel_id が空');

    var response = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
      method: 'post',
      contentType: 'application/json; charset=utf-8',
      headers: { 'Authorization': 'Bearer ' + token },
      payload: JSON.stringify({ channel: channelId, text: text, unfurl_links: false, unfurl_media: false }),
      muteHttpExceptions: true
    });

    var data;
    try { data = JSON.parse(response.getContentText()); }
    catch (e) { throw new Error('Slack応答解析失敗: ' + response.getContentText()); }

    if (!data.ok) {
      var hint = '';
      switch (data.error) {
        case 'not_in_channel':   hint = '→ ボットを招待してください (/invite @ボット名)'; break;
        case 'channel_not_found': hint = '→ プライベートチャンネルならボット招待を / channel_id が正しいか確認'; break;
        case 'invalid_auth':
        case 'token_revoked':
        case 'token_expired':    hint = '→ SLACK_BOT_TOKEN を確認 (xoxb-...)'; break;
        case 'is_archived':      hint = '→ そのチャンネルはアーカイブ済み'; break;
      }
      throw new Error('Slack投稿失敗: ' + data.error + ' ' + hint);
    }
  }


  // ===== 状態管理 =====

  function loadStateMap_(stateSheet) {
    var map = {};
    var lastRow = stateSheet.getLastRow();
    if (lastRow < 2) return map;
    var values = stateSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (var i = 0; i < values.length; i++) {
      var sheetName = String(values[i][COL_STATE_SOURCE]).trim();
      var period    = String(values[i][COL_STATE_PERIOD]).trim();
      if (sheetName) map[sheetName] = period;
    }
    return map;
  }

  function updateStateSheet_(stateSheet, sourceSheetName, newPeriod) {
    var lastRow = stateSheet.getLastRow();
    if (lastRow >= 2) {
      var values = stateSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < values.length; i++) {
        if (String(values[i][0]).trim() === sourceSheetName) {
          stateSheet.getRange(i + 2, 2).setValue(newPeriod);
          return;
        }
      }
    }
    var newRow = stateSheet.getLastRow() + 1;
    stateSheet.getRange(newRow, 1, 1, 2).setValues([[sourceSheetName, newPeriod]]);
  }


  // ===== ユーティリティ =====

  function pad2_(n) { return ('0' + n).slice(-2); }
  function toFullWidth_(s) { return String(s).replace(/[0-9]/g, function(d) { return FULLWIDTH_DIGITS[d]; }); }


  // ===== テスト・動作確認用関数 =====

  function dryRun() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var botsSheet = ss.getSheetByName('_bots');
    if (!botsSheet) { Logger.log('[ERROR] _bots シートなし'); return; }
    var now = new Date();
    Logger.log('=== dryRun 実行 (' + Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm') + ') ===');
    var lastRow = botsSheet.getLastRow();
    if (lastRow < 2) { Logger.log('_bots にデータなし'); return; }
    var botsData = botsSheet.getRange(2, 1, lastRow - 1, BOTS_NUM_COLS).getValues();
    for (var i = 0; i < botsData.length; i++) {
      var row = botsData[i];
      var botName     = String(row[COL_BOT_NAME]).trim();
      var enabled     = String(row[COL_ENABLED]).trim().toUpperCase();
      var sourceSheet = String(row[COL_SOURCE_SHEET]).trim();
      var template    = String(row[COL_TEMPLATE]).trim();
      if (!botName) continue;
      Logger.log('--- Bot: ' + botName + ' (enabled=' + enabled + ') ---');
      try {
        var rosterSheet = (sourceSheet && sourceSheet.toLowerCase() !== 'none') ? ss.getSheetByName(sourceSheet) : null;
        if (sourceSheet && sourceSheet.toLowerCase() !== 'none' && !rosterSheet) {
          Logger.log('  [WARNING] source_sheet "' + sourceSheet + '" 未存在');
          continue;
        }
        var message = resolveTemplate_(template, now, rosterSheet);
        Logger.log('  生成メッセージ:\n' + message);
      } catch (e) { Logger.log('  [ERROR] ' + e.message); }
    }
    Logger.log('=== dryRun 完了（投稿はしていません） ===');
  }

  function testDispatcherNow() {
    Logger.log('[TEST] testDispatcherNow: dispatcher() 呼び出し（実投稿あり）');
    dispatcher();
  }

  function testSpecificBot() {
    var botNameToTest = '朝稽古-金夕'; // ← テストしたいBot名

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var botsSheet = ss.getSheetByName('_bots');
    var stateSheet = ss.getSheetByName('_state');
    if (!botsSheet || !stateSheet) { Logger.log('[ERROR] _bots/_state なし'); return; }

    var now = new Date();
    var stateMap = loadStateMap_(stateSheet);
    var lastRow = botsSheet.getLastRow();
    if (lastRow < 2) return;
    var botsData = botsSheet.getRange(2, 1, lastRow - 1, BOTS_NUM_COLS).getValues();

    for (var i = 0; i < botsData.length; i++) {
      var row = botsData[i];
      if (String(row[COL_BOT_NAME]).trim() !== botNameToTest) continue;

      var sourceSheet = String(row[COL_SOURCE_SHEET]).trim();
      var rotUnit     = String(row[COL_ROTATION_UNIT]).trim().toLowerCase();
      var template    = String(row[COL_TEMPLATE]).trim();
      var channelId   = String(row[COL_CHANNEL_ID]).trim();
      var skipOffset  = String(row[COL_SKIP_OFFSET]).trim();

      Logger.log('[TEST] Bot: ' + botNameToTest + ' を強制実行');

      // スキップ日チェック（テストでも反映）
      if (skipOffset !== '' && skipOffset.toLowerCase() !== 'none') {
        var off = parseInt(skipOffset, 10);
        if (!isNaN(off)) {
          var targetDate = new Date(now.getTime() + off * 24 * 60 * 60 * 1000);
          var targetDateStr = Utilities.formatDate(targetDate, TZ, 'yyyy-MM-dd');
          if (isSkipDate_(ss, targetDateStr)) {
            Logger.log('[SKIP] 朝稽古お休み日のため停止（対象日=' + targetDateStr + '）');
            return;
          }
        }
      }

      if (rotUnit !== 'none' && sourceSheet && sourceSheet.toLowerCase() !== 'none') {
        var rosterSheet = ss.getSheetByName(sourceSheet);
        if (rosterSheet) {
          var currentPeriod = getCurrentPeriod_(now, rotUnit);
          var lastRotatedPeriod = stateMap[sourceSheet] || '';
          if (currentPeriod !== lastRotatedPeriod) {
            Logger.log('[ROTATE] ローテーション実行');
            rotateRoster_(rosterSheet);
            updateStateSheet_(stateSheet, sourceSheet, currentPeriod);
          }
        }
      }

      var roster = (sourceSheet && sourceSheet.toLowerCase() !== 'none') ? ss.getSheetByName(sourceSheet) : null;
      var message = resolveTemplate_(template, now, roster);
      Logger.log('[MESSAGE]\n' + message);
      postToSlack_(channelId, message);
      postToSlack_(channelId, message);
      Logger.log('[DONE] 投稿完了');
      return;
    }
    Logger.log('[ERROR] Bot "' + botNameToTest + '" が見つかりません');
  }

  function resetLastRun() {
    var botNameToReset = '掃除当番Bot'; // ← リセット対象のBot名
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var botsSheet = ss.getSheetByName('_bots');
    if (!botsSheet) return;
    var lastRow = botsSheet.getLastRow();
    if (lastRow < 2) return;
    var botsData = botsSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < botsData.length; i++) {
      if (String(botsData[i][0]).trim() === botNameToReset) {
        botsSheet.getRange(i + 2, COL_LAST_RUN + 1).setValue('');
        Logger.log('[RESET] ' + botNameToReset + ' の last_run クリア');
        return;
      }
    }
  }