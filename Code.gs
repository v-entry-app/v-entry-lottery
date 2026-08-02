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
