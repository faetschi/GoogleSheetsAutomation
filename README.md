# Google Sheets Task Manager

A **task management system** implemented entirely in Google Sheets with **automated task scheduling, Today view, and a calendar view**. Designed for easy task tracking, colored task indicators, and per-person assignment without needing a separate database.

## Features

- **Tasks Sheet**: Define your recurring tasks with start date, frequency, color, and active status.
- **TaskOccurrences Sheet**: Auto-generated list of task occurrences for the upcoming months.
- **TodayTasks Sheet**: Displays tasks due today with checkboxes to mark completion and assign a person.

**Calendar Sheet**: Month-view calendar showing tasks on their due dates with colored font and borders.
**Automatic Updates**:
  - Changes in Tasks propagate to TaskOccurrences, TodayTasks, and Calendar.
  - Editing the TodayTasks sheet updates TaskOccurrences.
**Custom Person Colors**: Assign colors to persons in the Persons sheet.

## Sheets Overview


| Sheet Name         | Purpose                                           |
|-------------------|-------------------------------------------------|
| `Tasks`            | Define tasks (ID, name, start date, frequency, color, active) |
| `TaskOccurrences`  | Auto-generated occurrences (read-only)        |
| `TodayTasks`       | Tasks due today; mark done and assign person  |
| `Calendar`         | Month-view calendar with tasks and colors     |
| `Persons`          | List of persons with optional font colors     |

## Setup

1. Open the Google Sheets file.
2. Go to **Extensions → Apps Script** and paste the project code (the script in `Code.js`).

3. Run `createInstallableOnEditTrigger()` once to enable automatic updates and grant OAuth scopes.
4. Set up your **Tasks** and **Persons** sheets.

## Usage

- Add or edit tasks in `Tasks`.

- Assign people and mark tasks as done in `TodayTasks`.
- View tasks in a monthly overview in `Calendar`.

## Notes

* Task colors are used in both TodayTasks and Calendar views.
* The Person column is maintained on occurrences and TodayTasks; it is not defined inside `Tasks` by default.
* The system auto-generates future occurrences for the next 12 months by default (configurable in `generateTaskOccurrences`).

## License

Copyright (c) 2025 faetschi — Licensed under the MIT License. See `LICENSE`.