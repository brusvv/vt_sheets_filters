# Visible Sort: аудит интеграции и готовность к переносу

Дата проверки: 13 ноября 2026  
Репозиторий: brusvv/vt_sheets_filters  
Ветка: chore/organize-repository  
Исходная ревизия: a7e2c14865e4cfea240b304f75db6af2c9cf6e39

## Итог

Перенос C:L не реализован. Статус TRANSFER_NOT_VERIFIED и запрет Range.sort() остаются активными. Проверка ограничена алгоритмом, Resolver, подготовкой видимости и интеграционными сценариями с mock-объектами Google Sheets.

## Подтверждённое исправление

До этой проверки applyFilters() при неуспешном Resolver всё равно переходил к buildFilterVisibility_(). Если ManualState содержал одинаковые ключи, FIFO consumeManualHiddenKey_() мог сопоставить ключ первой одинаковой записи, хотя физически скрыта была другая запись. Следствием могла стать смена скрытой физической позиции при последующем применении видимости.

Теперь при включённом AutoSort:
- валидный Resolver snapshot и совпадающие частоты ManualState позволяют продолжить фильтрацию после TRANSFER_NOT_VERIFIED; buildFilterVisibility_() получает resolvedManualHiddenRows и не использует FIFO;
- конфликт/несовпадение ManualState останавливает фильтрационную часть до изменения физической видимости;
- при resetVisibility=true и невалидном snapshot операция останавливается до showRows(), потому что иначе могли бы раскрываться строки, которые Resolver не смог классифицировать.

Обычная фильтрация при выключенном AutoSort не изменена.

Resolver дополнительно хранит исходные ManualState keys и метаданные в fingerprint. Чистая проверка snapshot теперь выявляет изменение захваченных полей объекта. Добавлена отдельная fresh-проверка, перечитывающая C:L, FilterState, ManualState и физическую видимость, а также сверяющая границу сценарных строк и Sheet ID/name. Она предназначена для обязательного optimistic-concurrency gate будущего commit-фазы; текущий сортировщик не вызывает её, потому что commit-фазы ещё нет.

## Покрытие тестами

### Unit tests

Существующий tests/visibleSort.test.cjs сохранён. Он покрывает сортировочную перестановку, стабильность, пустые/некорректные Score и Dots, сценарные блоки, pinned hidden slots, конфликты ManualState и отсутствие записей при TRANSFER_NOT_VERIFIED.

### Integration-like tests с mocks

Новый tests/visibleSort.integration.test.cjs вызывает реальные функции проекта внутри Node VM, в частности readSortVisibilityContext_(), sortScenarioBlocksDetailed_(), applyFilters() и runManualSortOperation_(). Mock Spreadsheet/Sheet/Range реализует чтение диапазонов, запись значений, физическое скрытие строк, FilterState и ManualState.

Проверяются:
- последовательность Resolver → Planner → Prepare FilterState → Filter Engine → Save FilterState;
- продолжение AutoSort-фильтрации при TRANSFER_NOT_VERIFIED с валидным snapshot;
- сохранение C:L, ManualState, FilterState и исходной физической скрытой позиции;
- resetVisibility: showRows выполняется только после Resolver, а ручная позиция скрывается повторно;
- конфликт ManualState: фильтрационная часть останавливается без изменения содержимого, FilterState, ManualState и физической видимости;
- ManualSort entry point: план строится, toast возвращается, но нет записи/показа/скрытия;
- AutoSort OFF: Resolver и Planner не вызываются;
- fresh-проверка отдельно обнаруживает изменение C, других данных C:L, FilterState, ManualState, физической видимости, границы и target sheet.

Эти тесты не являются тестами реального Google Sheets API.

### Реальная проверка Google Sheets

Не выполнялась. Не запускались Apps Script, clasp push/deploy или операции над рабочей таблицей. Не проверялись CopyPasteRequest, формулы, Rich Text, Notes, Data Validation, hyperlinks, пересчёт формул, merged ranges, row heights, группировки, conditional formatting или гонки с UI edits.

## Что ещё требуется до реализации переноса

1. Выполнить staging round-trip на отдельной копии рабочей книги для полного C:L: formulas и их относительные ссылки, Rich Text runs, number formats, форматирование, validation, notes и hyperlinks.
2. В commit-фазе повторно вызвать validateSortVisibilityContextFresh_() непосредственно перед первой записью и остановить commit при любом несовпадении fingerprint.
3. Учесть, что ScriptLock не блокирует UI edits. Между fresh-проверкой и записью остаётся неустранимое API-окно гонки; commit-план должен минимизировать это окно и проверять состояние после переноса.
4. Выполнить регрессионный прогон на копии Intermediate, включая скрытые строки внутри многорядных блоков, дубликаты, футер и инварианты вне C:L.

## Ограничения результата

Unit-тесты подтверждают чистую логику. Integration-like tests подтверждают порядок и поведение кода в mock-среде. Ни те, ни другие не доказывают совместимость CopyPasteRequest/Range.copyTo с реальной семантикой Google Sheets. TRANSFER_NOT_VERIFIED не снимается.
