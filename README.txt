IFC CRC Checker — полная сборка
================================

Функции
- Сверка XML↔IFC: имя (строго), CRC-32. Отчёт: ifc_crc_report.xlsx
- Сверка ИУЛ (PDF)↔IFC: имя, CRC-32, дата/время (mtime), размер, проверка имени PDF (мягко/строго). Отчёт: *_iul.xlsx
- Поддержка скан-PDF через OCR (PyMuPDF + pytesseract + установленный Tesseract-OCR)
- GUI с журналом, цветами, эмодзи, подтверждением перезаписи и опцией «Открыть по завершению»
- CLI с гибкими флагами и рекурсивным обходом папок

Установка зависимостей
    py -m pip install -r requirements.txt

Зависимости для разработки и тестов
    py -m pip install -r requirements.txt -r requirements-dev.txt

Дополнительно для OCR (сканы в PDF)
- Установите Tesseract-OCR с русским языком (rus).
- После установки проверьте в терминале: tesseract --version

Встраивание Tesseract в сборку
- Скачайте или установите Tesseract-OCR для Windows x64 (например, сборку от проекта UB Mannheim) и убедитесь, что в ней есть `tesseract.exe`, все `.dll` и папка `tessdata` с языковыми файлами `eng` и `rus`.
- Скопируйте всю папку установки Tesseract рядом с исходниками и переименуйте её в `tesseract` (рядом должны лежать `main_gui.py`, `build_exe.py` и т. д.). Репозиторий игнорирует эту директорию, поэтому она не попадёт в git.
- При запуске из исходников приложение автоматически найдёт `tesseract/tesseract.exe`. Для PyInstaller-сборки `build_exe.py` добавляет папку `tesseract` в дистрибутив, так что дополнительная установка не понадобится.
- Если требуется иной путь, можно задать его через переменную окружения `TESSDATA_PREFIX` или стандартные настройки pytesseract.

Запуск GUI
    py main_gui.py

Запуск CLI (примеры)
    # только XML↔IFC
    py main_cli.py --check-xml --xml "C:\meta.xml" --ifc-dir "C:\IFC" --recursive-ifc --force

    # только ИУЛ↔IFC (папка с PDF, рекурсивно, строгая проверка имени PDF)
    py main_cli.py --check-iul --ifc-dir "C:\IFC" --recursive-ifc --iul-dir "C:\IUL" --recursive-pdf --pdf-name-strict --force

Сборка .exe (Windows, PyInstaller)
    py -m pip install -r requirements.txt -r requirements-dev.txt
    py build_exe.py
Экзешники появятся в папке dist/.

Примечание по VS Code
- Если запускаете через «Play», убедитесь, что рабочая директория — это папка ifc_crc_checker_mod.
  Либо используйте запуск из терминала:
      cd "C:\...\ifc_crc_checker_mod"
      py main_gui.py
