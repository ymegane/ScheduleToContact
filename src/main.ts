/**
 * スプレッドシートが開かれたときにカスタムメニューを追加する関数
 */
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function onOpen(): void {
  const ui = SpreadsheetApp.getUi()
  ui.createMenu('欠席連絡')
    .addItem('連絡テキスト生成 (今月分)', 'generateAbsenceTextForThisMonth')
    .addToUi()
}

/**
 * メインの処理：今月の予定から連絡テキストを生成します。
 * (グループ化して書き出すバージョン)
 */
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function generateAbsenceTextForThisMonth(): void {
  const ui = SpreadsheetApp.getUi()

  // --- 1. 設定の読み込み ---
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const ruleSheet = ss.getSheetByName('ルール設定')
  const resultSheet = ss.getSheetByName('生成結果')

  // ★TS変更点: シートが存在するかチェック
  if (!ruleSheet) {
    ui.alert('エラー: 「ルール設定」シートが見つかりません。')
    return
  }
  if (!resultSheet) {
    ui.alert('エラー: 「生成結果」シートが見つかりません。')
    return
  }

  // ルール設定シートからルールを取得 (A列: キーワード, B列: 行動)
  // getValues() は any[][] を返すため、string[][] として扱う
  const rules: string[][] = ruleSheet
    .getRange(2, 1, ruleSheet.getLastRow() - 1, 2)
    .getValues() as string[][]

  // --- 2. カレンダーの読み込み ---
  const calendarId = PropertiesService.getScriptProperties().getProperty('CALENDAR_ID')
  let calendar: GoogleAppsScript.Calendar.Calendar

  if (calendarId) {
    calendar = CalendarApp.getCalendarById(calendarId)
    if (!calendar) {
      ui.alert(`エラー: ID「${calendarId}」のカレンダーが見つかりません。スクリプトプロパティを確認してください。`)
      return
    }
  } else {
    calendar = CalendarApp.getDefaultCalendar()
  }

  const today: Date = new Date()
  const startDate: Date = new Date(today.getFullYear(), today.getMonth(), 1)
  const endDate: Date = new Date(today.getFullYear(), today.getMonth() + 1, 0)

  const events: GoogleAppsScript.Calendar.CalendarEvent[] = calendar.getEvents(
    startDate,
    endDate
  )

  // --- 3. テキストの生成 (★TS変更点: Mapの型を明記) ---
  // key: 行動 (string), value: 連絡文の配列 (string[])
  const groupedResults = new Map<string, string[]>()

  // カレンダーの予定を一つずつチェック
  for (const event of events) {
    const eventTitle: string = event.getTitle() // 予定のタイトル

    // ルールを一つずつチェック
    for (const rule of rules) {
      const keyword: string = rule[0] // A列のキーワード
      const action: string = rule[1] // B列の行動

      // キーワードが空でなく、予定のタイトルにキーワードが含まれていたら
      if (keyword && eventTitle.includes(keyword)) {
        // 日付をフォーマット (例: "10月26日 (日)")
        const eventDateStr: string = Utilities.formatDate(
          event.getStartTime(),
          'Asia/Tokyo',
          'M月d日 (E)'
        )

        const line: string = `${eventDateStr} ${keyword}のため`

        // Mapにデータを格納
        if (!groupedResults.has(action)) {
          groupedResults.set(action, []) // actionをキーにして新しい配列を作成
        }
        // groupedResults.get(action) が undefined でないことを TypeScript に伝える ( ! )
        groupedResults.get(action)!.push(line)

        break
      }
    }
  }

  // --- 4. スプレッドシートへの書き込み ---
  resultSheet.clear()
  const outputData: string[][] = [] // スプレッドシートに書き込むための2D配列

  if (groupedResults.size > 0) {
    // Mapからキー（行動）ごとに処理
    for (const [action, lines] of groupedResults.entries()) {
      outputData.push([`[${action}]`])

      for (const line of lines) {
        outputData.push([`    ${line}`])
      }

      outputData.push([''])
    }

    // データをA1セルから一括書き込み
    resultSheet.getRange(1, 1, outputData.length, 1).setValues(outputData)
    ui.alert('連絡テキストを生成しました！')
  } else {
    resultSheet.getRange(1, 1).setValue('対象の予定は見つかりませんでした。')
    ui.alert('対象の予定は見つかりませんでした。')
  }
}
