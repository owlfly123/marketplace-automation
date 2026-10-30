# FLIP Builder

Конвертер заказов оптового партнёра (маркетплейс FLIP.KZ) в формат для загрузки: из исходного файла берутся нужные колонки, добавляется БИН контрагента, результат сохраняется в папку `__OUT` с датой в имени.

- Принимает XLSX и CSV (автоопределение разделителя и кодировки), можно выбрать сразу папку.
- Обработка идёт в отдельном потоке с прогресс-баром и логом.

```bash
pip install -r requirements.txt
python flip_builder.py
```

**EN.** Small batch converter (Python, openpyxl, tkinter) of wholesale partner orders into an upload format.
