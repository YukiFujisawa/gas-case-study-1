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
/* eslint-disable @typescript-eslint/no-unused-vars */

import { MeetingCalender } from './app/meeting-calender';

/**
 * Googleカレンダーのイベントをスプレッドシート(events)に連携する
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
