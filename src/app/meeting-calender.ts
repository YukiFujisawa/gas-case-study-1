/**
 * Copyright 2023 wywy LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
export class MeetingCalender {
  public static PREFIX = 'MTG';
  static getEventValues() {
    const calendarId = Session.getActiveUser().getEmail();
    const calendar = CalendarApp.getCalendarById(calendarId);
    const startTime = new Date();
    const entTime = new Date();
    entTime.setDate(startTime.getDate() + 30);
    const events = calendar.getEvents(startTime, entTime);
    return events
      .filter(event => {
        return event.getTitle().includes(MeetingCalender.PREFIX);
      })
      .map(event => {
        return [
          event.getId(),
          event.getTitle(),
          event.getStartTime(),
          event.getEndTime(),
          event.getLocation(),
          event.getDescription(),
          MeetingCalender.getCalenderEventLink(calendarId, event),
        ];
      });
  }
  /**
   * 予定へのリンクを取得します。
   * @param calendarId カレンダーID
   * @param event イベント
   * @returns 予定へのリンク
   */
  static getCalenderEventLink(
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
}
