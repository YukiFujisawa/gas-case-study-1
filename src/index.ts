/* eslint-disable @typescript-eslint/no-unused-vars */

const PREFIX_MTG = 'MTG';

/**
 * 自分のカレンダーから会議予定(イベント)を取得します。
 * @returns object[] 会議予定の配列
 *  [0:Event 1:ID	2:タイトル	3:開始日時	4:終了日時	5:場所	6:説明	7:URL]
 */
function getMeetingEventValues_() {
  const calendarId = Session.getActiveUser().getEmail();
  const calendar = CalendarApp.getCalendarById(calendarId);
  const startTime = new Date();
  const entTime = new Date();
  entTime.setDate(startTime.getDate() + 30);
  const events = calendar.getEvents(startTime, entTime);
  return events
    .filter(event => {
      return event.getTitle().includes(PREFIX_MTG);
    })
    .map(event => {
      return [
        event.getId(),
        event.getTitle(),
        event.getStartTime(),
        event.getEndTime(),
        event.getLocation(),
        event.getDescription(),
        getCalenderEventLink_(calendarId, event),
      ];
    });
}

/**
 * 予定へのリンクを取得します。
 * @param {string} calendarId カレンダーID
 * @param { GoogleAppsScript.Calendar.CalendarEvent} event イベント
 * @returns {string} 予定へのリンク
 */
function getCalenderEventLink_(
  calendarId: string,
  calendarEvent: GoogleAppsScript.Calendar.CalendarEvent
) {
  const baseUrl = 'https://www.google.com/calendar/event?eid=';
  const splitEventId = calendarEvent.getId().split('@');
  const eventUrl = `${baseUrl}${Utilities.base64Encode(
    splitEventId[0] + ' ' + calendarId
  )}`;
  return eventUrl;
}

/**
 * Googleカレンダーのイベントをスプレッドシート(eventsシート)に連携する
 */
function linkMeetingSchedules() {
  const eventsSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('events');
  if (!eventsSheet) {
    throw new Error('events sheet not found');
  }
  const eventIdSet = new Set();
  eventsSheet
    .getRange(5, 1, eventsSheet.getLastRow(), 1)
    .getValues()
    .forEach(row => {
      if (row[0]) {
        eventIdSet.add(row[0]);
      }
    });
  const events = getMeetingEventValues_();
  events
    .filter(event => {
      return !eventIdSet.has(event[0]);
    })
    .forEach(event => {
      eventsSheet.appendRow(event);
    });
}

/**
 * スプレッドシート(guestsシート)にしたがってカレンダー予定の参加者を更新する
 */
function updateMeetingGuests() {
  const guestsSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('guests');
  if (!guestsSheet) {
    throw new Error('guests sheet not found');
  }
  const eventsSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('events');
  if (!eventsSheet) {
    throw new Error('events sheet not found');
  }

  const guestValues = guestsSheet
    .getRange(5, 1, guestsSheet.getLastRow(), guestsSheet.getLastColumn())
    .getValues();

  // カレンダーの参加者にいない場合は追加する
  guestValues.forEach(row => {
    // row[0:Event ID	1:タイトル	2:開始時間	3:終了時間	4:メールアドレス	5:部署	6:名前]
    const eventId = row[0];
    if (!eventId) {
      return;
    }
    const event = CalendarApp.getEventById(eventId);
    if (!event) {
      return;
    }
    const guestEmail = row[4];
    if (!guestEmail) {
      return;
    }

    const exists = event.getGuestList().find(guest => {
      return guest.getEmail() === guestEmail;
    });
    if (!exists) {
      console.log('add guest:', guestEmail);
      event.addGuest(guestEmail);
    }
  });
  const eventValues = eventsSheet
    .getRange(5, 1, eventsSheet.getLastRow(), eventsSheet.getLastColumn())
    .getValues();

  // カレンダーの参加者にいるが、スプレッドシートにいない場合は削除する
  eventValues.forEach(row => {
    // row[0:Event 1:ID	2:タイトル	3:開始日時	4:終了日時	5:場所	6:説明	7:URL]
    const eventId = row[0];
    if (!eventId) {
      return;
    }
    const event = CalendarApp.getEventById(eventId);
    if (!event) {
      return;
    }

    const notExistGuests = event.getGuestList().filter(guest => {
      const foundGuest = guestValues.find(row => {
        return row[0] === eventId && row[4] === guest.getEmail();
      });
      if (foundGuest) {
        return false;
      }
      return true;
    });
    if (notExistGuests) {
      notExistGuests.forEach(guest => {
        console.log('remove guest:', guest.getEmail());
        event.removeGuest(guest.getEmail());
      });
    }
  });
}
