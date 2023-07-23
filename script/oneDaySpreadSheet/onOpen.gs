function onOpen() {
  const customMenu = SpreadsheetApp.getUi()
  customMenu.createMenu('OnDay作成')//メニューバーに表示するカスタムメニュー名
      .addItem('(1) 実行', 'main')//メニューアイテムを追加
      .addToUi()
}