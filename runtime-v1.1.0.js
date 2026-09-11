/* global Office, Excel, document, console */
"use strict";
var quickFillBusy = false;
var startupBusy = false;
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
  // Display the real task pane, without repeating the failed write.
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
  // Do not queue clicks: a queued operation could target a newer selection.
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
      if (color === null) selected.format.fill.clear();
      else selected.format.fill.color = color;
      return context.sync();
    });
  }).then(function fillSucceeded() {
    setText("status", "反映しました。");
  });
  // Finish the Office event before attempting any failure UI.
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

// Associate immediately, so command dispatch can find handlers during startup.
Office.actions.associate("fillYellow", fillYellow);
Office.actions.associate("fillOrange", fillOrange);
Office.actions.associate("fillBlue", fillBlue);
Office.actions.associate("fillGray", fillGray);
Office.actions.associate("clearFill", clearFill);

function showStartup(behavior) {
  setText("startup-status", behavior === Office.StartupBehavior.load ?
    "先読み：有効。このブックを次回開くときに準備を始めます。" :
    "先読み：無効。下のボタンで、このブックの先読みを有効にできます。");
}
function changeStartup(behavior) {
  if (startupBusy) return Promise.resolve();
  startupBusy = true;
  return officeReady.then(function saveStartup() {
    return Office.addin.setStartupBehavior(behavior);
  }).then(function savedStartup() {
    showStartup(behavior);
    startupBusy = false;
  }, function failedStartup(error) {
    startupBusy = false;
    setText("startup-status", describeError(error));
  });
}
function enableStartup() { return changeStartup(Office.StartupBehavior.load); }
function disableStartup() { return changeStartup(Office.StartupBehavior.none); }
function initializePane(info) {
  if (info.host !== Office.HostType.Excel) {
    setText("status", "Excelから開いてください。");
    return;
  }
  document.getElementById("enable-startup").onclick = enableStartup;
  document.getElementById("disable-startup").onclick = disableStartup;
  document.getElementById("enable-startup").disabled = false;
  document.getElementById("disable-startup").disabled = false;
  setText("status", "準備できました。ホームの色ボタンを使えます。");
  return Office.addin.getStartupBehavior().then(showStartup, function startupReadFailed(error) {
    setText("startup-status", describeError(error));
  });
}
officeReady.then(initializePane).catch(function initializationFailed(error) {
  setText("status", describeError(error));
  console.error("Quick Fill initialization", error);
});
