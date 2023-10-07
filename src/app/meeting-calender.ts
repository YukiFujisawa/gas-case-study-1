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
