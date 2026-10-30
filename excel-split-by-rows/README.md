# Разбивка таблицы на файлы по 300 строк (VBA)

Личный кабинет принимал загрузку не больше 300 строк за раз. Макрос делит первый лист книги на отдельные файлы `ДД.ММ.ГГГГ_1.xlsx`, `_2.xlsx`… в папке с исходником. Размер куска задаётся в настройках.

Как подключить: Excel → Alt+F11 → File → Import File → `SplitToRowsFiles.bas`, сохранить книгу как .xlsm.

**EN.** VBA macro that splits a large sheet into files of N rows for marketplace upload limits.
