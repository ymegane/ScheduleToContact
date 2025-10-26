function doGet(e: GoogleAppsScript.Events.DoGet) {
  return HtmlService.createHtmlOutputFromFile('index');
}

/**
 * ----------------------------------------------------------------
 * Core Logic Function
 * ----------------------------------------------------------------
 */
function _generateContactTextData() {
  // --- 1. 設定の読み込み ---
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const ruleSheet = ss.getSheetByName('ルール設定')
  if (!ruleSheet) {
    throw new Error('エラー: 「ルール設定」シートが見つかりません。')
  }
  const rules: any[][] = ruleSheet
    .getRange(2, 1, ruleSheet.getLastRow() - 1, 4)
    .getValues()

  // --- 2. カレンダーの読み込み ---
  const calendarId = PropertiesService.getScriptProperties().getProperty('CALENDAR_ID')
  let calendar: GoogleAppsScript.Calendar.Calendar
  if (calendarId) {
    calendar = CalendarApp.getCalendarById(calendarId)
    if (!calendar) {
      throw new Error(`エラー: ID「${calendarId}」のカレンダーが見つかりません。スクリプトプロパティを確認してください。`)
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

  // --- 3. テキストの生成 ---
  const groupedResults = new Map<string, [string, Date][]>()
  const requiredKeywords = rules.filter(rule => rule[3] === true).map(rule => rule[0] as string)
  const foundRequiredKeywords = new Set<string>()

  for (const event of events) {
    const eventTitle: string = event.getTitle()
    for (const rule of rules) {
      const keyword: string = rule[0]
      const outputWord: string = rule[1]
      const action: string = rule[2]
      const isRequired: boolean = rule[3]

      if (keyword && eventTitle.includes(keyword)) {
        const startTime = event.getStartTime()
        const wordToUse = outputWord || keyword
        const line: string = `${wordToUse}のため`

        if (!groupedResults.has(action)) {
          groupedResults.set(action, [])
        }
        groupedResults.get(action)!.push([line, startTime as any])

        if (isRequired) {
          foundRequiredKeywords.add(keyword)
        }
        break
      }
    }
  }

  const missingKeywords = requiredKeywords.filter(keyword => !foundRequiredKeywords.has(keyword))

  return { groupedResults, events, missingKeywords }
}


/**
 * ----------------------------------------------------------------
 * Web App Endpoint
 * ----------------------------------------------------------------
 */
function generateTextForWebApp(): { mainOutput: string, debugEvents: {time: string, title: string}[], missingKeywordsWarning: string } {
  try {
    const { groupedResults, events, missingKeywords } = _generateContactTextData();

    // --- 4. 出力文字列の組み立て ---
    let mainOutput = ''
    if (groupedResults.size > 0) {
      for (const [action, lines] of groupedResults.entries()) {
        mainOutput += `[${action}]\n`
        for (const line of lines) {
          const description = line[0]
          const date = line[1]
          const dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'M/d HH:mm')
          mainOutput += `    ${description} (${dateStr})\n`
        }
        mainOutput += '\n'
      }
    } else {
      mainOutput += '対象の予定は見つかりませんでした。\n\n'
    }

    // --- 5. デバッグ情報 ---
    const debugEvents = events
      .sort((a, b) => a.getStartTime().getTime() - b.getStartTime().getTime())
      .map(event => {
        const startTime = event.getStartTime();
        const dayOfWeek = startTime.getDay();
        const dayOfWeekStr = ['日', '月', '火', '水', '木', '金', '土'][dayOfWeek];
        return {
          date: Utilities.formatDate(startTime, 'Asia/Tokyo', `M/d (${dayOfWeekStr})`),
          time: Utilities.formatDate(startTime, 'Asia/Tokyo', 'HH:mm'),
          title: event.getTitle()
        }
      });

    // --- 6. 必須予定のチェックと警告 ---
    let missingKeywordsWarning = '';
    if (missingKeywords.length > 0) {
      missingKeywordsWarning = `警告: 以下の必須予定が見つかりませんでした。\n・${missingKeywords.join('\n・')}`
    }

    return { mainOutput, debugEvents, missingKeywordsWarning }
  } catch (e) {
    if (e instanceof Error) {
      // Return error in the same format
      return { mainOutput: e.message, debugEvents: [], missingKeywordsWarning: '' };
    }
    return { mainOutput: String(e), debugEvents: [], missingKeywordsWarning: '' };
  }
}

/**
 * ----------------------------------------------------------------
 * Spreadsheet-bound Functions
 * ----------------------------------------------------------------
 */
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function onOpen(): void {
  const ui = SpreadsheetApp.getUi()
  ui.createMenu('欠席連絡')
    .addItem('連絡テキスト生成 (今月分)', 'generateAbsenceTextForThisMonth')
    .addToUi()
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function generateAbsenceTextForThisMonth(): void {
  const ui = SpreadsheetApp.getUi()
  try {
    const resultSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('生成結果')
    if (!resultSheet) {
      ui.alert('エラー: 「生成結果」シートが見つかりません。')
      return
    }

    const { groupedResults, events, missingKeywords } = _generateContactTextData();

    // --- 4. スプレッドシートへの書き込み ---
    resultSheet.clear()
    const outputData: (string | Date)[][] = []

    if (groupedResults.size > 0) {
      for (const [action, lines] of groupedResults.entries()) {
        outputData.push([`[${action}]`, ''])
        for (const line of lines) {
          const description = line[0]
          const date = line[1]
          const dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'M/d HH:mm')
          outputData.push([`    ${description}`, dateStr])
        }
        outputData.push(['', ''])
      }
    } else {
      outputData.push(['対象の予定は見つかりませんでした。', ''])
      outputData.push(['', ''])
    }

    // --- 5. デバッグ用に取得した予定を書き出す ---
    outputData.push(['--- 取得したカレンダーの予定 --- ', ''])
    if (events.length > 0) {
      for (const event of events) {
        const startTime = Utilities.formatDate(event.getStartTime(), 'Asia/Tokyo', 'M/d HH:mm');
        outputData.push([`${startTime} ${event.getTitle()}`, '']);
      }
    } else {
      outputData.push(['（予定なし）', ''])
    }

    // データをA1セルから一括書き込み
    resultSheet.getRange(1, 1, outputData.length, 2).setValues(outputData)
    
    ui.alert('連絡テキストを生成しました！')

    // --- 6. 必須予定のチェックと警告 ---
    if (missingKeywords.length > 0) {
      ui.alert(`警告: 以下の必須予定が見つかりませんでした。\n・${missingKeywords.join('\n・')}`)
    }

    resultSheet.activate()

  } catch (e) {
    if (e instanceof Error) {
      ui.alert(e.message);
    } else {
      ui.alert(String(e));
    }
  }
}