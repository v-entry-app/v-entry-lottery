const MAX_API_PAGES = 20; // 最大2,000人まで取得

/**
 * スプレッドシートを開いたときにメニューを追加
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("V-entry抽選")
    .addItem("① シートを初期設定", "setupLotterySheet")
    .addItem("② Xから応募者を取り込む", "importLikingUsers")
    .addSeparator()
    .addItem("③ 当選3名＋補欠2名を抽選", "runLottery")
    .addToUi();
}

/**
 * シートのレイアウトを作成
 * 既存の応募者や結果は消しません
 */
function setupLotterySheet() {
  const sheet = SpreadsheetApp.getActiveSheet();

  sheet.getRange("A1:D1").setValues([[
    "応募者X ID",
    "除外",
    "Xリンク",
    "表示名"
  ]]);

  sheet.getRange("E1:F6").setValues([
    ["区分", "抽選結果"],
    ["当選1", ""],
    ["当選2", ""],
    ["当選3", ""],
    ["補欠1", ""],
    ["補欠2", ""]
  ]);

  // G列はF列（抽選結果）からD列（表示名）を引く
  applyWinnerNameFormulas_(sheet);

  sheet.getRange("H1:H5").setValues([
    ["投稿URL"],
    ["取込日時"],
    ["取得人数"],
    ["抽選日時"],
    ["抽選シード"]
  ]);

  sheet.getRange("J1").setValue("YouTube / リンク（参考）");

  // B列を除外用チェックボックスにする
  sheet
    .getRange(2, 2, sheet.getMaxRows() - 1, 1)
    .insertCheckboxes();

  sheet.getRange("A1:J1")
    .setFontWeight("bold")
    .setBackground("#d9ead3");

  sheet.getRange("I1")
    .setBackground("#fff2cc")
    .setNote("抽選対象となるX投稿のURLを入力してください");

  sheet.getRange("I2").setNumberFormat("yyyy-MM-dd HH:mm:ss");
  sheet.getRange("I4").setNumberFormat("yyyy-MM-dd HH:mm:ss");

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 70);
  sheet.setColumnWidth(3, 240);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 90);
  sheet.setColumnWidth(6, 180);
  sheet.setColumnWidth(7, 180);
  sheet.setColumnWidth(8, 100);
  sheet.setColumnWidth(9, 320);
  sheet.setColumnWidth(10, 320);
  sheet.setFrozenRows(1);

  SpreadsheetApp.getUi().alert(
    "初期設定が完了しました",
    "I1セルに募集投稿のURLを入力してください。",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * X投稿に「いいね」したユーザーを取り込む
 */
function importLikingUsers() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();

  try {
    // 抽選後の名簿差し替えを防止
    if (hasLotteryResults_(sheet)) {
      ui.alert(
        "抽選済みです",
        "抽選結果が確定しているため、応募者を取り込み直せません。",
        ui.ButtonSet.OK
      );
      return;
    }

    validateCredentials_();

    const postUrl = String(sheet.getRange("I1").getValue()).trim();
    const tweetId = extractTweetId_(postUrl);

    if (!tweetId) {
      ui.alert(
        "投稿URLを確認してください",
        "I1セルにXの投稿URLを入力してください。",
        ui.ButtonSet.OK
      );
      return;
    }

    const existingIds = sheet
      .getRange(2, 1, sheet.getMaxRows() - 1, 1)
      .getDisplayValues()
      .flat()
      .some(value => value.trim() !== "");

    if (existingIds) {
      const answer = ui.alert(
        "応募者を取り込み直しますか？",
        "現在の応募者一覧と除外設定は消去されます。",
        ui.ButtonSet.YES_NO
      );

      if (answer !== ui.Button.YES) return;
    }

    const baseUrl =
      `https://api.x.com/2/tweets/${tweetId}/liking_users`;

    let paginationToken = "";
    let pageCount = 0;
    let allUsers = [];

    do {
      const query = {
        max_results: "100",
        "user.fields": "description,entities,location,url"
      };

      if (paginationToken) {
        query.pagination_token = paginationToken;
      }

      const response = oauth1Get_(baseUrl, query);
      const statusCode = response.getResponseCode();
      const responseText = response.getContentText();

      let body;

      try {
        body = JSON.parse(responseText);
      } catch (error) {
        throw new Error(
          `X APIの応答を読み取れませんでした（${statusCode}）。`
        );
      }

      if (statusCode < 200 || statusCode >= 300) {
        throw createXApiError_(statusCode, body);
      }

      if (Array.isArray(body.data)) {
        allUsers = allUsers.concat(body.data);
      }

      paginationToken =
        body.meta && body.meta.next_token
          ? body.meta.next_token
          : "";

      pageCount++;

    } while (paginationToken && pageCount < MAX_API_PAGES);

    if (allUsers.length === 0) {
      ui.alert(
        "応募者が見つかりませんでした",
        "投稿URL、いいね数、投稿の公開状態を確認してください。",
        ui.ButtonSet.OK
      );
      return;
    }

    // XのユーザーIDを基準に重複を除外
    const uniqueUsers = [];
    const seenUserIds = new Set();

    allUsers.forEach(user => {
      if (!user.id || !user.username) return;
      if (seenUserIds.has(user.id)) return;

      seenUserIds.add(user.id);
      uniqueUsers.push(user);
    });

    uniqueUsers.sort((a, b) =>
      a.username.localeCompare(b.username, "en", {
        sensitivity: "base"
      })
    );

    // 行数が足りなければ追加
    const requiredRows = uniqueUsers.length + 1;

    if (requiredRows > sheet.getMaxRows()) {
      sheet.insertRowsAfter(
        sheet.getMaxRows(),
        requiredRows - sheet.getMaxRows()
      );
    }

    // 古い応募者・除外設定を消去
    sheet
      .getRange(2, 1, sheet.getMaxRows() - 1, 4)
      .clearContent();

    sheet
      .getRange(2, 2, sheet.getMaxRows() - 1, 1)
      .insertCheckboxes();

    // C列は本人のXプロフィールへのリンク
    const rows = uniqueUsers.map(user => [
      `@${user.username}`,
      false,
      `https://x.com/${user.username}`,
      user.name || ""
    ]);

    sheet.getRange(2, 1, rows.length, 4).setValues(rows);

    // J列は参考情報。空欄でも抽選対象です
    // YouTubeが見つからない場合はプロフィールに設定されたURL（lit.linkなど）
    sheet.getRange("J1").setValue("YouTube / リンク（参考）");
    sheet
      .getRange(2, 10, sheet.getMaxRows() - 1, 1)
      .clearContent();

    const linkRows = uniqueUsers.map(user => [
      findProfileLink_(user)
    ]);

    sheet
      .getRange(2, 10, linkRows.length, 1)
      .setValues(linkRows);

    sheet.setColumnWidth(3, 240);
    sheet.setColumnWidth(10, 320);

    sheet.getRange("I2").setValue(new Date());
    sheet.getRange("I3").setValue(uniqueUsers.length);

    SpreadsheetApp.flush();

    let message =
      `${uniqueUsers.length}名を取り込みました。\n\n` +
      "応募条件を満たさない方は、B列の「除外」にチェックしてください。";

    if (paginationToken) {
      message +=
        `\n\n※${MAX_API_PAGES * 100}名で取得を打ち切っています。`;
    }

    ui.alert("取り込み完了", message, ui.ButtonSet.OK);

  } catch (error) {
    ui.alert(
      "取り込みに失敗しました",
      error.message || String(error),
      ui.ButtonSet.OK
    );
  }
}

/**
 * 当選3名＋補欠2名を一度だけ抽選
 */
function runLottery() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();

  if (hasLotteryResults_(sheet)) {
    ui.alert(
      "抽選済みです",
      "すでに結果が確定しています。再抽選は行いません。",
      ui.ButtonSet.OK
    );
    return;
  }

  const answer = ui.alert(
    "最終確認",
    "当選3名＋補欠2名を抽選します。\n一度確定すると、ボタンを押しても再抽選されません。\n\n実行しますか？",
    ui.ButtonSet.YES_NO
  );

  if (answer !== ui.Button.YES) return;

  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(5000)) {
    ui.alert("ほかの抽選処理が実行中です。");
    return;
  }

  try {
    // 二重クリック対策
    if (hasLotteryResults_(sheet)) {
      ui.alert("すでに抽選済みです。");
      return;
    }

    const rows = sheet
      .getRange(2, 1, sheet.getMaxRows() - 1, 2)
      .getValues();

    const candidates = [];
    const seen = new Set();

    rows.forEach(([rawId, excluded]) => {
      const displayId = String(rawId).trim();

      if (!displayId || excluded === true) return;

      const normalizedId = normalizeXId_(displayId);

      if (!normalizedId || seen.has(normalizedId)) return;

      seen.add(normalizedId);

      candidates.push({
        displayId: displayId.startsWith("@")
          ? displayId
          : `@${normalizedId}`,
        normalizedId
      });
    });

    if (candidates.length < 5) {
      ui.alert(
        "抽選できません",
        `有効な応募者は${candidates.length}名です。\n` +
        "当選3名＋補欠2名を選ぶには、5名以上必要です。",
        ui.ButtonSet.OK
      );
      return;
    }

    /*
     * 抽選用のランダムシードを作成。
     * 各応募者のIDとシードからSHA-256値を作り、
     * 値の小さい順に並べます。
     */
    const seed =
      Utilities.getUuid().replace(/-/g, "") +
      Utilities.getUuid().replace(/-/g, "");

    const rankedCandidates = candidates
      .map(candidate => ({
        displayId: candidate.displayId,
        sortKey: sha256Hex_(
          `${seed}|${candidate.normalizedId}`
        )
      }))
      .sort((a, b) => a.sortKey.localeCompare(b.sortKey));

    const selected = rankedCandidates.slice(0, 5);
    const labels = ["当選1", "当選2", "当選3", "補欠1", "補欠2"];

    sheet.getRange("E1:F6").setValues([
      ["区分", "抽選結果"],
      ...labels.map((label, index) => [
        label,
        selected[index].displayId
      ])
    ]);

    // G列にVLOOKUPで表示名を表示
    applyWinnerNameFormulas_(sheet);

    sheet.getRange("I4").setValue(new Date());
    sheet.getRange("I5").setValue(seed);

    SpreadsheetApp.flush();

    const names = sheet
      .getRange("G2:G6")
      .getDisplayValues()
      .flat();

    const message = labels
      .map((label, index) => {
        const name = String(names[index] || "").trim();

        return name
          ? `${label}：${selected[index].displayId}（${name}）`
          : `${label}：${selected[index].displayId}`;
      })
      .join("\n");

    ui.alert("🎉 抽選結果 🎉", message, ui.ButtonSet.OK);

  } catch (error) {
    ui.alert(
      "抽選に失敗しました",
      error.message || String(error),
      ui.ButtonSet.OK
    );
  } finally {
    lock.releaseLock();
  }
}

/**
 * G列にF列（抽選結果のX ID）からD列（表示名）を引くVLOOKUPを設定
 * A列が「@あり」「@なし」どちらでも引けるようにしています
 */
function applyWinnerNameFormulas_(sheet) {
  sheet.getRange("G1").setValue("表示名");

  const formulas = [];

  for (let row = 2; row <= 6; row++) {
    formulas.push([
      `=IF($F${row}="","",` +
      `IFERROR(VLOOKUP($F${row},$A:$D,4,FALSE),` +
      `IFERROR(VLOOKUP(SUBSTITUTE($F${row},"@",""),$A:$D,4,FALSE),"")))`
    ]);
  }

  sheet.getRange("G2:G6").setFormulas(formulas);
}

/**
 * OAuth 1.0a署名付きGETリクエスト
 */
function oauth1Get_(baseUrl, queryParams) {
  const properties =
    PropertiesService.getScriptProperties().getProperties();

  const apiKey = properties.X_API_KEY;
  const apiSecret = properties.X_API_SECRET;
  const accessToken = properties.X_ACCESS_TOKEN;
  const accessTokenSecret =
    properties.X_ACCESS_TOKEN_SECRET;

  const oauthParams = {
    oauth_consumer_key: apiKey,
    oauth_nonce: Utilities.getUuid().replace(/-/g, ""),
    oauth_signature_method: "HMAC-SHA1",
    oauth_timestamp: Math.floor(Date.now() / 1000).toString(),
    oauth_token: accessToken,
    oauth_version: "1.0"
  };

  const signatureParams = Object.assign(
    {},
    queryParams,
    oauthParams
  );

  const parameterString = Object.keys(signatureParams)
    .map(key => [
      percentEncode_(key),
      percentEncode_(signatureParams[key])
    ])
    .sort((a, b) => {
      const keyComparison = a[0].localeCompare(b[0]);
      return keyComparison !== 0
        ? keyComparison
        : a[1].localeCompare(b[1]);
    })
    .map(pair => `${pair[0]}=${pair[1]}`)
    .join("&");

  const signatureBaseString = [
    "GET",
    percentEncode_(baseUrl),
    percentEncode_(parameterString)
  ].join("&");

  const signingKey =
    `${percentEncode_(apiSecret)}&` +
    `${percentEncode_(accessTokenSecret)}`;

  const signatureBytes = Utilities.computeHmacSignature(
    Utilities.MacAlgorithm.HMAC_SHA_1,
    signatureBaseString,
    signingKey,
    Utilities.Charset.UTF_8
  );

  oauthParams.oauth_signature =
    Utilities.base64Encode(signatureBytes);

  const authorizationHeader =
    "OAuth " +
    Object.keys(oauthParams)
      .sort()
      .map(key =>
        `${percentEncode_(key)}="${percentEncode_(oauthParams[key])}"`
      )
      .join(", ");

  const queryString = Object.keys(queryParams)
    .map(key =>
      `${percentEncode_(key)}=${percentEncode_(queryParams[key])}`
    )
    .join("&");

  return UrlFetchApp.fetch(
    `${baseUrl}?${queryString}`,
    {
      method: "get",
      headers: {
        Authorization: authorizationHeader,
        Accept: "application/json"
      },
      muteHttpExceptions: true
    }
  );
}

function validateCredentials_() {
  const properties =
    PropertiesService.getScriptProperties().getProperties();

  const required = [
    "X_API_KEY",
    "X_API_SECRET",
    "X_ACCESS_TOKEN",
    "X_ACCESS_TOKEN_SECRET"
  ];

  const missing = required.filter(name => !properties[name]);

  if (missing.length > 0) {
    throw new Error(
      "スクリプトプロパティが不足しています：\n" +
      missing.join("\n")
    );
  }
}

function extractTweetId_(value) {
  const text = String(value).trim();

  if (/^\d{5,25}$/.test(text)) {
    return text;
  }

  const match = text.match(/\/status(?:es)?\/(\d{5,25})/i);
  return match ? match[1] : "";
}

function normalizeXId_(value) {
  return String(value)
    .trim()
    .replace(
      /^https?:\/\/(?:www\.)?(?:x\.com|twitter\.com)\//i,
      ""
    )
    .split(/[/?#]/)[0]
    .replace(/^@/, "")
    .toLowerCase();
}

function hasLotteryResults_(sheet) {
  return sheet
    .getRange("F2:F6")
    .getDisplayValues()
    .flat()
    .some(value => value.trim() !== "");
}

function sha256Hex_(value) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    value,
    Utilities.Charset.UTF_8
  );

  return bytes
    .map(byte => {
      const unsignedByte = (byte + 256) % 256;
      return unsignedByte.toString(16).padStart(2, "0");
    })
    .join("");
}

function percentEncode_(value) {
  return encodeURIComponent(String(value))
    .replace(/[!'()*]/g, character =>
      "%" +
      character
        .charCodeAt(0)
        .toString(16)
        .toUpperCase()
    );
}

function createXApiError_(statusCode, body) {
  let detail = "";

  if (body && body.detail) {
    detail = body.detail;
  } else if (body && body.title) {
    detail = body.title;
  } else if (
    body &&
    Array.isArray(body.errors) &&
    body.errors.length > 0
  ) {
    detail = body.errors
      .map(error => error.message || error.detail || String(error))
      .join("\n");
  }

  const guidance = {
    401:
      "API Key、Secret Key、Access Token、Access Token Secretを確認してください。",
    403:
      "Xアプリのユーザー認証設定、Read権限、APIプランを確認してください。",
    429:
      "X APIの利用回数制限に達しました。しばらく待ってから再実行してください。"
  };

  return new Error(
    `X APIエラー（${statusCode}）\n` +
    `${guidance[statusCode] || "X APIから拒否されました。"}\n` +
    (detail ? `\n${detail}` : "")
  );
}

/**
 * J列に入れるURLを決める。
 * YouTubeが見つかればそれを、なければ本人が
 * プロフィールに設定しているURL（lit.linkなど）を返す。
 */
function findProfileLink_(user) {
  return findYouTubeUrl_(user) || findWebsiteUrl_(user);
}

/**
 * Xプロフィールの説明・場所・Webサイトから
 * 直接記載されたYouTube URLを探す
 */
function findYouTubeUrl_(user) {
  const candidates = [];

  function addEntityUrls(items) {
    (items || []).forEach(item => {
      candidates.push(
        item.unwound_url,
        item.expanded_url,
        item.display_url,
        item.url
      );
    });
  }

  if (user.entities) {
    if (user.entities.url) {
      addEntityUrls(user.entities.url.urls);
    }

    if (user.entities.description) {
      addEntityUrls(user.entities.description.urls);
    }
  }

  candidates.push(
    user.url,
    user.description,
    user.location
  );

  for (const candidate of candidates) {
    if (!candidate) continue;

    const match = String(candidate).match(
      /(?:https?:\/\/)?(?:www\.|m\.)?(?:youtube\.com|youtu\.be)(?:\/[^\s<>"'）)\]}]+)?/i
    );

    if (!match) continue;

    let url = match[0].replace(/[.,。、]+$/, "");

    if (!/^https?:\/\//i.test(url)) {
      url = `https://${url}`;
    }

    return url;
  }

  return "";
}

/**
 * 本人がプロフィールに設定しているURLを探す。
 * Webサイト欄 → 自己紹介のリンク → 自己紹介の本文 の順。
 * t.coの短縮URLは展開後のURLを優先します。
 */
function findWebsiteUrl_(user) {
  const candidates = [];

  function addEntityUrls(items) {
    (items || []).forEach(item => {
      candidates.push(
        item.unwound_url,
        item.expanded_url,
        item.display_url
      );
    });
  }

  // Webサイト欄
  if (user.entities && user.entities.url) {
    addEntityUrls(user.entities.url.urls);
  }

  // 自己紹介に貼られたリンク
  if (user.entities && user.entities.description) {
    addEntityUrls(user.entities.description.urls);
  }

  // 展開済みURLが取れなかったときの保険
  candidates.push(user.url, user.description);

  const urls = [];

  candidates.forEach(candidate => {
    if (!candidate) return;

    const url = extractUrl_(String(candidate));

    if (url) urls.push(url);
  });

  const expanded = urls.find(
    url => !/^https?:\/\/t\.co\//i.test(url)
  );

  return expanded || urls[0] || "";
}

/**
 * 文字列から最初のURLらしき部分を取り出し、
 * https:// を補って返す
 */
function extractUrl_(text) {
  const match = text.match(
    /(?:https?:\/\/)?(?:[a-z0-9-]+\.)+[a-z]{2,}(?:\/[^\s<>"'）)\]}]*)?/i
  );

  if (!match) return "";

  let url = match[0].replace(/[.,。、]+$/, "");

  if (!/^https?:\/\//i.test(url)) {
    url = `https://${url}`;
  }

  return url;
}
