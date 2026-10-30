# Автоматизация работы менеджера маркетплейсов

Инструменты, которые я сделала для себя, когда работала менеджером по ключевым партнёрам в книжном издательстве (Ozon, Wildberries, Яндекс Маркет, Мегамаркет). Каждый из них заменил ручную рутину, которая повторялась каждую неделю.

Писала с помощью ИИ (Claude, Cursor, ChatGPT): ставила задачу как ТЗ, проверяла результат на реальных файлах и дорабатывала до рабочей версии, которой пользовались каждый день.

> Все реквизиты, ФИО, номера и продажи в репозитории заменены вымышленными.

| Проект | Что делает | Стек | Экономия |
|---|---|---|---|
| [wb-order-planner](wb-order-planner/) | План отгрузки на региональные склады WB: ABC-анализ, прогноз спроса, учёт всплесков, роста и товаров в пути | Python, pandas, openpyxl | — |
| [ttn-generator](ttn-generator/) | Приложение для заполнения транспортных накладных для поставок Ozon | Python, tkinter, openpyxl, PyInstaller | 2–4 ч/нед |
| [ozon-orders-to-1c](ozon-orders-to-1c/) | Пакетное преобразование заказов из ЛК Ozon в формат загрузки в 1С | PowerShell, Excel COM | 2–4 ч/нед |
| [wb-report-summarizer](wb-report-summarizer/) | Итоги по еженедельным финансовым отчётам WB (20–50 тыс. строк) с учётом возвратов | Python, pandas, tkinter | 1–2 ч/нед |
| [flip-builder](flip-builder/) | Конвертер заказов для оптового партнёра в нужный формат | Python, openpyxl, tkinter | — |
| [excel-split-by-rows](excel-split-by-rows/) | Макрос: разбивает большую таблицу на файлы по 300 строк для загрузки в ЛК | VBA | — |

---

## Marketplace manager automation (EN)

Tools I built for my own daily work as a key account manager for marketplaces (Ozon, Wildberries) at a book publisher. Each one replaced a weekly manual routine.

Built with AI assistants (Claude, Cursor, ChatGPT): I wrote the specs, tested on real files and iterated until the tools were used in production every day.

- **wb-order-planner** — weekly replenishment plan for Wildberries regional warehouses: ABC analysis, demand forecast with spike detection and growth factor, stock-in-transit.
- **ttn-generator** — desktop app that fills in Russian consignment notes (TTN) for Ozon shipments. Saves 2–4 h/week.
- **ozon-orders-to-1c** — batch converter of Ozon orders into a 1C import format. Saves 2–4 h/week.
- **wb-report-summarizer** — totals across weekly WB financial reports (20–50k rows), returns handled. Saves 1–2 h/week.
- **flip-builder**, **excel-split-by-rows** — smaller converters (Python, VBA).

All company data has been replaced with fictitious values.
