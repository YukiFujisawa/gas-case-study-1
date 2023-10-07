/* eslint-disable @typescript-eslint/no-unused-vars */

import { MeetingCalender } from './app/meeting-calender';

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
  const events = MeetingCalender.getEventValues();
  events
    .filter(event => {
      return !eventIdSet.has(event[0]);
    })
    .forEach(event => {
      eventsSheet.appendRow(event);
    });
}

/**
 * スプレッドシート(guestsシート)にしたがって、カレンダー予定の参加者を更新する
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
    const eventId = row[0];
    if (!eventId) {
      return;
    }
    const event = CalendarApp.getEventById(eventId);
    if (!event) {
      return;
    }
    // 0:Event ID	1:タイトル	2:開始時間	3:終了時間	4:メールアドレス	5:部署	6:名前
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
    // Event ID	タイトル	開始日時	終了日時	場所	説明	URL
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
