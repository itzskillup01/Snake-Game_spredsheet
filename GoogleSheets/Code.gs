function onOpen() {
  SpreadsheetApp.getUi().createMenu('Snake Game')
    .addItem('Play Classic Snake', 'showSnakeGame').addToUi();
}
function showSnakeGame() {
  const html = HtmlService.createHtmlOutputFromFile('Game')
    .setTitle('Classic Snake').setWidth(420).setHeight(620);
  SpreadsheetApp.getUi().showSidebar(html);
}
function saveSnakeHighScore(score) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Snake');
  if (!sh) sh = ss.insertSheet('Snake');
  sh.getRange('A1').setValue('Classic Snake High Score');
  const best = Number(sh.getRange('B1').getValue()) || 0;
  if (Number(score) > best) sh.getRange('B1').setValue(Number(score));
}
