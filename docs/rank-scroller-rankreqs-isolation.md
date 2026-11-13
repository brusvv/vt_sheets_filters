# Rank Scroller: изоляция RankReqs

Дата аудита: 2026-11-13  
Базовый commit: `55cfce9e1ec493f168bcce9699c760b2215824a0`  
Ветка: `chore/organize-repository`

## Вывод

В исходном коде на базовом commit прямой путь от `resetFilters()` к изменению формул или физической видимости столбцов листа `RankReqs` не обнаружен.

Если активен `RankReqs`, `getTargetSheet()` намеренно выбирает fallback-лист `CONFIG.sheetName`, который сейчас равен `Intermediate`. Поэтому `resetFilters()` очищает `Intermediate!I1:J1`, применяет фильтр к `Intermediate` и передаёт имя `Intermediate` в Rank Scroller.

`rankScroller(false, ..., 'Intermediate', ...)` принудительно обрабатывает Intermediate, поскольку он задан как `targetSheetName`. Это объясняет `forceUpdate=true` при `forceAll=false`. Оно не означает полную синхронизацию всех листов.

Конфигурация Rank Scroller перечисляет только `Novice`, `Intermediate` и `Advanced`. Для Intermediate пороговый диапазон — N:AC. Сами операции видимости исполняются на объекте листа, полученном через `spreadsheet.getSheetByName(config.sheetName)`.

## Фактический путь вызова

```text
resetFilters()
  ├─ SpreadsheetApp.getActiveSpreadsheet()
  ├─ getTargetSheet(null, spreadsheet)
  │    ├─ active = RankReqs
  │    ├─ RankReqs не входит в RANK_TRACKED_SHEETS
  │    └─ target = CONFIG.sheetName = Intermediate
  ├─ clearContent() -> Intermediate!I1:J1
  ├─ applyFilters(false, Intermediate, resetVisibility=true, ...)
  │    ├─ фильтрует строки Intermediate
  │    ├─ showRows/hideRows только Intermediate
  │    └─ сохраняет состояние в _FilterState
  ├─ SpreadsheetApp.flush()
  └─ rankScroller(false, ..., targetSheetName='Intermediate', ...)
       ├─ Intermediate: isTargetSheet=true -> forceUpdate=true
       ├─ Novice/Advanced: пропускаются, если не dirty
       └─ Intermediate: hideColumns(N:AC), showColumns(рассчитанный фрагмент)
```

Ветка `handleEdit()` для `RankReqs` — отдельный путь: при правке B4:C70 она помечает три игровых листа dirty и запускает `runRankScroller(false)`. Rank Scroller после этого может синхронизировать все три игровых листа. Сам `RankReqs` в эту обработку не передаётся.

## Обнаруженные границы и защита

В базовом коде имелась слабая граница доверия: основной цикл Rank Scroller строился напрямую из ключей `CONFIG.rankScroller.sheets`, а низкоуровневые функции видимости принимали произвольный объект `Sheet`. Текущая конфигурация была корректна, но ошибочное добавление `RankReqs` в неё или прямой вызов visibility helper не имели дополнительного allow-list барьера.

Добавлена fail-closed изоляция:

- конфигурация Rank Scroller должна содержать ровно три имени из `RANK_TRACKED_SHEETS`; добавочный или отсутствующий лист блокирует запуск до обращения к листам;
- явная цель Rank Scroller проверяется по allow-list;
- обработчик листа проверяет имя до `getSheetByName()` и сверяет имя полученного объекта;
- обе низкоуровневые функции `hideColumns/showColumns` отказываются работать с листом вне allow-list;
- `getTargetSheet()` отклоняет явно переданное имя, не являющееся игровым листом;
- `applyFilters()` отклоняет переданный объект неигрового листа до чтения или записи диапазонов.

Фильтрационная логика, формулы и настройки CONFIG не менялись. Список игровых листов не расширялся.

## Карта изменений и операций

| Функция / модуль | Цель | Операция |
|---|---|---|
| `resetFilters()` | выбранный игровой лист или fallback Intermediate | `clearContent(I1:J1)` |
| `applyFilters()` / FilterVisibility | игровой лист | `showRows/hideRows` в строках данных |
| FilterState | `_FilterState` | `setValues`, `clearContent`; при необходимости создаёт и скрывает служебный лист `_FilterState` |
| `readRankScrollerSheetData_()` | один из трёх игровых листов | чтение Energy и порогов из M2:AB2 или M2:AC2 |
| `showRankScrollerColumns_()` | один из трёх игровых листов | скрывает N:AB или N:AC и показывает рассчитанный сегмент |
| `showRankScrollerTierColumns_()` | один из трёх игровых листов | скрывает N:AB или N:AC и показывает группы кратных 100 порогов |
| `duplicateScenario()` | игровой лист из `getTargetSheet()` | вставляет строку и копирует диапазоны; не вызывается из resetFilters |
| Visible Sort | игровой лист | при перестановке возвращает `TRANSFER_NOT_VERIFIED`; переноса C:L нет |

В просмотренном проектном коде не найдено вызовов `setColumnWidth()`, `setColumnWidths()`, `deleteColumn()`, `deleteColumns()`, `deleteDimension()`, `setFormula()` или `setFormulas()`, которые могли бы объяснить описанное изменение `RankReqs`. Записи `setValues/clearContent` используются для фильтров и служебного `_FilterState`; `resetFilters()` очищает I1:J1 только у выбранного игрового листа.

Исключение по типу действия — `duplicateScenario()`: оно использует `insertRowAfter()` и `copyTo()` на целевом игровом листе. Но этот путь не вызывается из `resetFilters()`.

## Отдельная зависимость пользовательской функции

`core/customFunctions.gs`, функция `energy()`, читает ранги через:

```js
SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  .getRange(2, 10, 1, reqs.length)
  .getValues()[0]
```

То есть `energy()` зависит от активной вкладки, а не получает лист-источник явным аргументом. Если формулы `RankReqs` используют `energy()`, эта зависимость заслуживает отдельной проверки. В рамках этой задачи функция и формулы не менялись: без выгрузки реальных формул нельзя безопасно выбрать корректный диапазон-источник или менять контракт пользовательской функции.

## Что говорят наблюдения

- `targetSheetName='Intermediate'` согласуется с fallback-логикой `getTargetSheet()`, если в момент запуска активной вкладкой был `RankReqs`.
- `forceUpdate=true` при `forceAll=false` ожидаемо, потому что Intermediate явно передан как цель.
- Изменение Energy 563 → 580 и расчёт текущей колонки S относятся к данным Intermediate.
- Эти наблюдения не доказывают, что Rank Scroller изменял `RankReqs`: объект `RankReqs` не является целью его операций.

## Почему причина #REF! пока не установлена

В репозитории нет содержимого формул рабочего листа `RankReqs`, снимка книги до/после сбоя или журнала истории версий Google Sheets. Поэтому нельзя установить, какие именно формулы были в H3, K3, N3, Q3, T3, H26, K26, N26, Q26 и T26, когда они стали `#REF!`, и изменились ли их ссылки до появления ошибки.

Из кода `resetFilters()` и Rank Scroller не следует операция удаления листов/столбцов или перезаписи формул на `RankReqs`. Скрытие столбца само по себе не переписывает формулу или её ссылку. Пересчёт пользовательской функции `energy()` при этом является отдельной потенциальной зависимостью, но доступных данных недостаточно, чтобы связать её с конкретными `#REF!`.

Для установления первопричины нужны история версий Google Sheets и сами формулы/их прежние значения. Формулы здесь намеренно не восстанавливались и не реконструировались.

## Регрессионные тесты

Добавлен `tests/rankScroller.isolation.test.cjs`. Тесты используют отдельные mock-листы `Novice`, `Intermediate`, `Advanced`, `RankReqs` и `Options`, регистрируют операции видимости, прямые записи и изменения формул.

Проверяются:

1. Полный запуск обращается только к трём разрешённым листам.
2. `RankReqs` как явная цель Rank Scroller приводит к отказу до обращения к листам.
3. Добавление `RankReqs` в CONFIG блокирует весь запуск до операций.
4. Engine отклоняет `RankReqs` до `getSheetByName()`.
5. Низкоуровневые visibility helpers не выполняют `hideColumns/showColumns` на `RankReqs`.
6. Обновление Intermediate не меняет формулы и состояние видимости mock-листа RankReqs.
7. `resetFilters()` при активном RankReqs маршрутизирует операцию на Intermediate и не пишет в формулы RankReqs.
8. Общий helper и `applyFilters()` отклоняют явно переданный RankReqs.

Это unit/integration-like mock-тесты. Они доказывают границы вызовов в текущем коде под заданной моделью объектов, но не являются экспериментом с настоящим Google Sheets API и не подтверждают поведение пересчёта формул Google Sheets.

## Не выполнялось

- Apps Script не запускался.
- Рабочая Google Sheets не открывалась и не изменялась.
- Формулы и значения RankReqs не восстанавливались.
- Тесты не обращались к Google Sheets API.
- Защита `TRANSFER_NOT_VERIFIED` не снималась; перенос C:L не реализовывался.
