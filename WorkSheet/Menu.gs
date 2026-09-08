function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // ==========================================================
  // CALENDAR
  // ==========================================================

  ui.createMenu('Month Sheet')
    .addItem(
      'Create Current Month',
      'createCurrentMonthSheet'
    )
    .addToUi();


  // ==========================================================
  // SETUP
  // ==========================================================

  ui.createMenu('Setup')
    .addItem(
      'Create Lists Sheet',
      'createListsSheet'
    )
    .addItem(
      'Format Current Sheet',
      'formatCurrentSheet'
    )
    .addToUi();


  // ==========================================================
  // TASKS
  // ==========================================================

  ui.createMenu('Tasks')
    .addItem(
      'Fill task for today',
      'fillTaskForToday'
    )
    .addItem(
      'Fill task for selected date in the current month',
      'fillTaskForSelectedDate'
    )
    .addToUi();


  // ==========================================================
  // DSR
  // ==========================================================

  ui.createMenu('DSR')
    .addItem(
      'Generate DSR for today',
      'generateDSRForToday'
    )
    .addItem(
      'Generate DSR for selected date in the month',
      'generateDSRForSelectedDate'
    )
    .addToUi();
}