function onOpen() {
  const customMenu = SpreadsheetApp.getUi()
  customMenu.createMenu('必要人数割り出し')//メニューバーに表示するカスタムメニュー名
      .addItem('(1) csvをインポート', 'main')//メニューアイテムを追加
      .addItem('(2) 日ごとの入力に割り当て', 'getSheetData')//メニューアイテムを追加
      .addToUi()
}