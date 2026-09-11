/* global Office, Excel, document, console */
"use strict";

var quickFillBusy = false;
var officeReady = Office.onReady();

function setText(id, message) {
  var element = document.getElementById(id);
  if (element) element.textContent = message;
}

function describeError(error) {
  var code = error && error.code ? error.code : "UnknownError";
  if (code === "InvalidOperationInCellEditMode") {
    return "セルの入力を確定してから、もう一度押してください。";
  }
  return "処理に失敗しました（" + code + "）。" +
    (error && error.message ? " " + error.message : "");
}

function showFailure(error) {
  console.error("Quick Fill", error);
  setText("status", describeError(error));

  // 色変更に失敗した場合だけ、状態確認用のパネルを表示する。
  if (Office.addin && Office.addin.showAsTaskpane) {
    return Office.addin.showAsTaskpane().catch(function paneFailure(paneError) {
      console.error("Quick Fill: could not open status pane", paneError);
    });
  }
}

function completeCommand(event) {
  if (event && typeof event.completed === "function") event.completed();
}

function applyFill(color, event) {
  // 連打を待ち行列に入れない。遅延実行時に別の選択範囲を塗る事故を避ける。
  if (quickFillBusy) {
    setText("status", "前の処理中です。完了後にもう一度押してください。");
    completeCommand(event);
    return Promise.resolve();
  }

  quickFillBusy = true;

  var operation = officeReady.then(function runFill() {
    setText("status", "反映中…");
    return Excel.run(function updateSelection(context) {
      var selected = context.workbook.getSelectedRanges();
      if (color === null) {
        selected.format.fill.clear();
      } else {
        selected.format.fill.color = color;
      }
      return context.sync();
    });
  }).then(function fillSucceeded() {
    setText("status", "反映しました。");
  });

  // Officeのコマンド完了通知を先に済ませてから、必要ならエラーUIを表示する。
  return operation.then(function finishSuccess() {
    quickFillBusy = false;
    completeCommand(event);
  }, function finishFailure(error) {
    quickFillBusy = false;
    completeCommand(event);
    return showFailure(error);
  });
}

function fillYellow(event) { return applyFill("#FFFF00", event); }
function fillOrange(event) { return applyFill("#FFA500", event); }
function fillBlue(event) { return applyFill("#00B0F0", event); }
function fillGray(event) { return applyFill("#A9A9A9", event); }
function clearFill(event) { return applyFill(null, event); }

// 起動中でもコマンドを見つけられるよう、ハンドラーは直ちに登録する。
Office.actions.associate("fillYellow", fillYellow);
Office.actions.associate("fillOrange", fillOrange);
Office.actions.associate("fillBlue", fillBlue);
Office.actions.associate("fillGray", fillGray);
Office.actions.associate("clearFill", clearFill);

function ensureStartupPreload() {
  // 設定ボタンを置かず、このブックでは先読みを自動的に有効にする。
  // 既に有効なら書き込みは行わない。
  if (!Office.addin ||
      !Office.StartupBehavior ||
      typeof Office.addin.getStartupBehavior !== "function" ||
      typeof Office.addin.setStartupBehavior !== "function") {
    console.warn("Quick Fill: startup preload API is unavailable.");
    return Promise.resolve();
  }

  return Office.addin.getStartupBehavior().then(function checkStartup(behavior) {
    if (behavior === Office.StartupBehavior.load) {
      return;
    }
    return Office.addin.setStartupBehavior(Office.StartupBehavior.load);
  }).catch(function preloadFailure(error) {
    // 先読みに失敗しても色ボタン自体は使えるため、操作を妨げない。
    console.warn("Quick Fill: could not enable startup preload.", error);
  });
}

function initializeRuntime(info) {
  if (info.host !== Office.HostType.Excel) {
    setText("status", "Excelから開いてください。");
    return;
  }

  setText("status", "準備できました。ホームの色ボタンを使えます。");
  return ensureStartupPreload();
}

officeReady.then(initializeRuntime).catch(function initializationFailed(error) {
  setText("status", describeError(error));
  console.error("Quick Fill initialization", error);
});
