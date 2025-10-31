
# VED Scraper

![Python](https://img.shields.io/badge/python-3.10+-blue)
![PyQt5](https://img.shields.io/badge/PyQt5-required-orange)
![Status](https://img.shields.io/badge/status-active-success)

Программа для сбора оценок студентов с [dec.srspu.ru](https://dec.srspu.ru) и замены номеров зачеток на ФИО.

Автор: [@Oleger12](https://t.me/Oleger12)

---

## 📌 Описание

VED Scraper позволяет:
- Собирать оценки студентов по дисциплинам с портала DEC SRSPU.
- Сохранять HTML-страницы студентов и общий CSV с оценками.
- Заменять номера зачеток на ФИО в CSV.
- Работать через локальные куки браузера или удалённый cookie-server.
- Использовать удобный GUI с тёмной и светлой темами.

---

## ⚡ Функционал

1. **Сбор оценок:**
   - Поддержка локальных куков из популярных браузеров.
   - Возможность запроса cookies с cookie-server.
   - Автоматическая генерация CSV с оценками.
   - Сохранение HTML каждой страницы студента для последующего анализа.

2. **Замена номеров зачеток на ФИО:**
   - Вручную или автоматически по CSV.
   - Замена производится в новом CSV (`all_students_with_names.csv`).

3. **GUI:**
   - Удобный интерфейс на PyQt5.
   - Логи выполнения.
   - Прогрессбар для отслеживания сбора данных.
   - Вкладки для сбора оценок, замены ФИО и информации о программе.
   - Светлая и тёмная темы.

---

## 🛠 Установка

1. Клонируем репозиторий:

```bash
git clone https://github.com/Oleger1000/SRSPU_Ved_Scrapper.git
cd SRSPU_Ved_Scrapper
````

2. Устанавливаем зависимости:

```bash
pip install -r requirements.txt
```

**requirements.txt:**

```
requests
beautifulsoup4
browser-cookie3
PyQt5
```

---

## 🚀 Использование

1. Запускаем GUI:

```bash
python pyqtgui.py
```

2. **Сбор оценок:**

   * Введите login и password от cookie-server (или используйте локальные куки браузера).
   * Нажмите **"Запустить сбор оценок"**.
   * Дождитесь завершения и сохранения CSV (`ved_results/all_students.csv`).

3. **Замена номеров зачеток на ФИО:**

   * Перейдите на вкладку **"Замена ФИО"**.
   * Введите соответствия `номер зачетки, ФИО` построчно или используйте автозаполнение.
   * Нажмите **"Заменить на ФИО в CSV"**.
   * Новый CSV будет сохранён как `all_students_with_names.csv`.

---

## 📂 Структура проекта

```
ved-scraper/
│
├─ gui.py                 # Главный GUI
├─ final_parser2.py       # Основной парсер и функции работы с DEC
├─ ved_results/           # Папка для сохранённых HTML и CSV
│   ├─ html/
│   ├─ all_students.csv
│   └─ all_students_with_names.csv
├─ assets/                # Иконки для GUI
├─ README.md
└─ requirements.txt
```

---

## 🌈 Темы

Программа поддерживает:

* Светлую тему (по умолчанию)
* Тёмную тему (переключается кнопкой в вкладке "О программе")

---

## 🔒 Безопасность

* Логины и пароли не сохраняются.
* Куки браузера подставляются только локально для доступа к DEC.
* Можно использовать удалённый cookie-server для безопасного получения cookies.

---

## 📝 Лицензия

MIT License © 2025 [Oleger12](https://t.me/Oleger12)

---

## 📞 Контакты

* Telegram: [@Oleger12](https://t.me/Oleger12)
* GitHub: [Oleger1000](https://github.com/Oleger1000)

```
