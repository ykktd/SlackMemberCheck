/**
 * フォーム回答とSlackメンバーを突合し、A列（加入準備WS参加）を更新する。
 *
 * シート列定義:
 * A: 加入準備WS参加（チェックボックス）
 * B: タイムスタンプ
 * C: メールアドレス
 * D: お名前（漢字）
 */
function checkSlackMembership() {
  var SHEET_NAME = "フォームの回答 1";
  var UNMATCHED_SHEET_NAME = "Slack未対応";
  var TOKEN_KEY = "SLACK_BOT_TOKEN";
  var CHECKED = true;
  var UNCHECKED = false;

  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error("対象シートが見つかりません: " + SHEET_NAME);
    }

    var token = PropertiesService.getScriptProperties().getProperty(TOKEN_KEY);
    if (!token) {
      throw new Error(
        "スクリプトプロパティにトークンがありません: " + TOKEN_KEY,
      );
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log("処理対象データがありません。");
      return;
    }

    var rowCount = lastRow - 1;
    var values = sheet.getRange(2, 1, rowCount, 4).getValues();

    var slackMembers = fetchSlackMembers_(token);
    var slackIndex = buildSlackIndex_(slackMembers);
    var responseIndex = buildResponseIndex_(values);
    var unmatchedSlackMembers = findUnmatchedSlackMembers_(
      slackMembers,
      responseIndex,
    );

    var outputStatuses = [];
    var updatedCount = 0;
    var matchedByEmailCount = 0;
    var matchedByNameCount = 0;
    var notMatchedCount = 0;

    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var currentStatus = row[0];

      // 既にTRUEの行は確定済みとして保持する。FALSE/空欄は毎回再判定する。
      if (currentStatus === true) {
        outputStatuses.push([currentStatus]);
        continue;
      }

      var formEmail = normalizeEmail_(row[2]);
      var formNameNormalized = normalizeName_(row[3]);
      var nextStatus = UNCHECKED;

      if (formEmail && slackIndex.emailSet.has(formEmail)) {
        nextStatus = CHECKED;
        matchedByEmailCount++;
      } else if (
        formNameNormalized &&
        slackIndex.normalizedNameSet.has(formNameNormalized)
      ) {
        nextStatus = CHECKED;
        matchedByNameCount++;
      } else {
        notMatchedCount++;
      }

      outputStatuses.push([nextStatus]);
      updatedCount++;
    }

    var statusRange = sheet.getRange(2, 1, rowCount, 1);
    statusRange.insertCheckboxes();
    statusRange.setValues(outputStatuses);

    writeUnmatchedSlackMembers_(
      spreadsheet,
      UNMATCHED_SHEET_NAME,
      unmatchedSlackMembers,
    );

    Logger.log(
      "Slack突合完了: totalRows=%s, updated=%s, emailMatch=%s, nameMatch=%s, notMatched=%s, slackMembers=%s, unmatchedSlackMembers=%s, unmatchedSheet=%s",
      rowCount,
      updatedCount,
      matchedByEmailCount,
      matchedByNameCount,
      notMatchedCount,
      slackMembers.length,
      unmatchedSlackMembers.length,
      UNMATCHED_SHEET_NAME,
    );
  } catch (error) {
    Logger.log(
      "checkSlackMembership failed: %s",
      error && error.stack ? error.stack : error,
    );
    throw error;
  }
}

/**
 * 回答シートの照合用インデックスを構築する。
 *
 * @param {Array<Array<*>>} rows 回答シートの行データ（A〜D列）
 * @return {{emailSet: Set<string>, normalizedNameSet: Set<string>}}
 */
function buildResponseIndex_(rows) {
  var emailSet = new Set();
  var normalizedNameSet = new Set();

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (!row) {
      continue;
    }

    var email = normalizeEmail_(row[2]);
    var normalizedName = normalizeName_(row[3]);

    if (email) {
      emailSet.add(email);
    }
    if (normalizedName) {
      normalizedNameSet.add(normalizedName);
    }
  }

  return {
    emailSet: emailSet,
    normalizedNameSet: normalizedNameSet,
  };
}

/**
 * Slackメンバーのうち、回答シートで対応者が見つからない人を抽出する。
 *
 * @param {Array<{email: string, realName: string, normalizedName: string}>} slackMembers
 * @param {{emailSet: Set<string>, normalizedNameSet: Set<string>}} responseIndex
 * @return {Array<{email: string, realName: string, normalizedName: string}>}
 */
function findUnmatchedSlackMembers_(slackMembers, responseIndex) {
  var unmatched = [];

  for (var i = 0; i < slackMembers.length; i++) {
    var member = slackMembers[i];
    if (!member) {
      continue;
    }

    var matchedByEmail =
      member.email && responseIndex.emailSet.has(member.email);
    var matchedByName =
      !matchedByEmail &&
      member.normalizedName &&
      responseIndex.normalizedNameSet.has(member.normalizedName);

    if (!matchedByEmail && !matchedByName) {
      unmatched.push(member);
    }
  }

  return unmatched;
}

/**
 * Slack未対応者一覧を別シートへ上書き出力する。
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @param {string} sheetName
 * @param {Array<{email: string, realName: string, normalizedName: string}>} unmatchedMembers
 */
function writeUnmatchedSlackMembers_(spreadsheet, sheetName, unmatchedMembers) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  sheet.clearContents();

  var now = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy/MM/dd HH:mm:ss",
  );
  var headers = [
    ["最終更新", "Slack表示名", "Slackメールアドレス", "正規化氏名"],
  ];
  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);

  if (!unmatchedMembers || unmatchedMembers.length === 0) {
    sheet
      .getRange(2, 1, 1, headers[0].length)
      .setValues([[now, "未対応者は0件です", "", ""]]);
    return;
  }

  var output = [];
  for (var i = 0; i < unmatchedMembers.length; i++) {
    var member = unmatchedMembers[i];
    output.push([
      now,
      toSafeString_(member.realName),
      toSafeString_(member.email),
      toSafeString_(member.normalizedName),
    ]);
  }

  sheet.getRange(2, 1, output.length, headers[0].length).setValues(output);
}

/**
 * Slack users.list をページネーションで全件取得し、照合に必要なデータへ整形する。
 *
 * @param {string} token Slack Bot Token
 * @return {Array<{email: string, realName: string, normalizedName: string}>}
 */
function fetchSlackMembers_(token) {
  var members = [];
  var cursor = "";
  var endpoint = "https://slack.com/api/users.list";

  while (true) {
    var params = {
      method: "get",
      muteHttpExceptions: true,
      headers: {
        Authorization: "Bearer " + token,
      },
    };

    var query = "?limit=200";
    if (cursor) {
      query += "&cursor=" + encodeURIComponent(cursor);
    }

    var response = UrlFetchApp.fetch(endpoint + query, params);
    var statusCode = response.getResponseCode();
    var bodyText = response.getContentText();

    if (statusCode === 429) {
      var retryAfter = response.getHeaders()["Retry-After"];
      throw new Error(
        "Slack API rate limited (429). Retry-After=" + retryAfter,
      );
    }

    if (statusCode < 200 || statusCode >= 300) {
      throw new Error(
        "Slack API request failed: status=" + statusCode + ", body=" + bodyText,
      );
    }

    var json;
    try {
      json = JSON.parse(bodyText);
    } catch (parseError) {
      throw new Error(
        "Slack API response parse failed: " + parseError + ", body=" + bodyText,
      );
    }

    if (!json.ok) {
      throw new Error("Slack API returned ok=false: error=" + json.error);
    }

    var pageMembers = json.members || [];
    for (var i = 0; i < pageMembers.length; i++) {
      var user = pageMembers[i];
      if (!user || user.deleted || user.is_bot) {
        continue;
      }

      var profile = user.profile || {};
      var email = normalizeEmail_(profile.email);
      if (!email) {
        continue;
      }

      var realName = toSafeString_(
        user.real_name || profile.real_name || user.name,
      );
      var normalizedName = normalizeName_(realName);

      members.push({
        email: email,
        realName: realName,
        normalizedName: normalizedName,
      });
    }

    cursor = "";
    if (json.response_metadata && json.response_metadata.next_cursor) {
      cursor = String(json.response_metadata.next_cursor).trim();
    }

    if (!cursor) {
      break;
    }
  }

  return members;
}

/**
 * 氏名を照合用に正規化する。
 * 1) 全角/半角スペース削除
 * 2) Unicode正規化(NFKC)
 *
 * @param {*} value 氏名文字列
 * @return {string}
 */
function normalizeName_(value) {
  var name = toSafeString_(value);
  if (!name) {
    return "";
  }

  var noSpaces = name.replace(/[ \u3000]/g, "");
  return noSpaces.normalize("NFKC").trim();
}

/**
 * メールアドレスを比較用に正規化する。
 *
 * @param {*} value メールアドレス
 * @return {string}
 */
function normalizeEmail_(value) {
  return toSafeString_(value).trim().toLowerCase();
}

/**
 * Slackメンバー照合用のインデックスを構築する。
 *
 * @param {Array<{email: string, normalizedName: string}>} members
 * @return {{emailSet: Set<string>, normalizedNameSet: Set<string>}}
 */
function buildSlackIndex_(members) {
  var emailSet = new Set();
  var normalizedNameSet = new Set();

  for (var i = 0; i < members.length; i++) {
    var member = members[i];
    if (!member) {
      continue;
    }

    if (member.email) {
      emailSet.add(member.email);
    }
    if (member.normalizedName) {
      normalizedNameSet.add(member.normalizedName);
    }
  }

  return {
    emailSet: emailSet,
    normalizedNameSet: normalizedNameSet,
  };
}

/**
 * null/undefined を安全に文字列化する。
 *
 * @param {*} value
 * @return {string}
 */
function toSafeString_(value) {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value);
}
