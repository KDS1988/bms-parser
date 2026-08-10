# BMS Event Watcher

Гибридная архитектура:

- **GitHub Actions** (`poll_bms.py`, по крону) — сам логинится в BMS, тянет
  мероприятия с услугами "Прямая трансляция"/"Студия", шлёт уведомления в
  Telegram, и передаёт данные в Apps Script по HTTP.
- **Google Apps Script** (`Code.gs`) — принимает данные через веб-эндпоинт
  (`doPost`) и пишет их на листы: «Новые мероприятия», «Тех блок», доска по
  месяцам. Также отдельно: вход в BMS (ручной, через Telegram-код), «График
  персонала» (ежедневный сбор + мгновенный пересчёт периода по правке ячейки).

```
BMS API  <──  GitHub Actions (опрос, Telegram)  ──HTTP──>  Apps Script (запись в таблицу)
                                                                  ↑
                                              вход в BMS, график персонала — тоже здесь
```

## Структура

- `Code.gs` — код Apps Script (запись в таблицу, вход в BMS, график персонала)
- `appsscript.json` — манифест Apps Script
- `poll_bms.py` — опрос BMS + Telegram, запускается в GitHub Actions
- `.github/workflows/bms-poll.yml` — расписание запуска `poll_bms.py`
- `.clasp.json.example` — шаблон конфига [clasp](https://github.com/google/clasp)

## Установка

### 1. Apps Script

```bash
npm install -g @google/clasp
clasp login
cp .clasp.json.example .clasp.json   # впиши свой Script ID (Project Settings в Apps Script)
clasp push
```

Секреты (не в коде, в Script Properties — см. `setupSecrets()` в `Code.gs`):
1. `clasp open`
2. Впиши свои значения вместо плейсхолдеров прямо в `setupSecrets()` в редакторе
3. Выполни `setupSecrets()` (▶ Run) один раз
4. Верни плейсхолдеры обратно, это НЕ коммить — секреты уже сохранены отдельно

Нужны: `TELEGRAM_BOT_TOKEN`, `TELEGRAM_CHAT_ID`, `BMS_PHONE`, и новый — **`WEBHOOK_SECRET`** (любая длинная случайная строка, придумай сам — например `openssl rand -hex 32` в терминале).

### 2. Деплой веб-эндпоинта (для приёма данных от GitHub)

В Apps Script: **Deploy → New deployment**
- Type: **Web app**
- Execute as: **Me**
- Who has access: **Anyone**
- Deploy → скопируй URL вида `https://script.google.com/macros/s/XXXX/exec`

Это и есть `APPS_SCRIPT_WEBHOOK_URL` для следующего шага. Проверка запросов — через `WEBHOOK_SECRET` в теле каждого запроса (сверяется в `doPost`), так что открытый доступ ("Anyone") не проблема — без правильного секрета запрос отклоняется.

**Важно:** при каждом `clasp push` деплой веб-приложения нужно переразвернуть заново (Deploy → Manage deployments → ✏️ → New version), иначе изменения в `doPost` не подхватятся автоматически на уже выданном URL.

### 3. Вход в BMS для самого Apps Script

Меню таблицы **BMS → 🔑 Войти в BMS** (нужен для «Графика персонала» и ручных тестов — сам приём данных через `doPost` в BMS не ходит, только пишет то, что прислал GitHub).

### 4. GitHub Secrets

В репозитории: **Settings → Secrets and variables → Actions → New repository secret**

| Secret | Что это |
|---|---|
| `BMS_REFRESH_TOKEN` | см. ниже, разовый бутстрап |
| `TELEGRAM_BOT_TOKEN` | тот же токен, что и в Apps Script |
| `TELEGRAM_CHAT_ID` | тот же chat_id |
| `APPS_SCRIPT_WEBHOOK_URL` | URL из шага 2 |
| `WEBHOOK_SECRET` | та же строка, что в `setupSecrets()` |
| `GH_PAT` | Personal Access Token с правом менять секреты этого репозитория (см. ниже) |

**`BMS_REFRESH_TOKEN` — разовый бутстрап.** GitHub Actions не может пройти
интерактивный вход по SMS/Telegram-коду сам, поэтому используем уже
залогиненную сессию Apps Script как отправную точку:
1. В Apps Script выполни в редакторе: `Logger.log(PropertiesService.getScriptProperties().getProperty('BMS_REFRESH_TOKEN'))`
2. Скопируй значение из логов (View → Logs)
3. Вставь как `BMS_REFRESH_TOKEN` в GitHub Secrets

Дальше GitHub Actions ведёт **свою собственную** сессию — токен ротируется
при каждом запуске и автоматически перезаписывается в секретах (см. workflow).
Если этот токен протухнет насовсем (например, репозиторий долго не запускал
workflow) — повтори бутстрап.

**`GH_PAT`** — Settings (аккаунта) → Developer settings → Personal access
tokens → Fine-grained tokens → New token → выбери этот репозиторий →
Permissions → **Secrets: Read and write**. Без этого шаг обновления
`BMS_REFRESH_TOKEN` в конце workflow не сработает (обычный `GITHUB_TOKEN`,
который есть в любом workflow по умолчанию, править секреты не может).

### 5. Триггеры Apps Script (ставятся вручную, clasp их не переносит)

⏰ Triggers → **+ Add Trigger**:

| Функция | Тип | Когда |
|---|---|---|
| `dailyStaffScheduleFetch` | Time-driven, Day timer | 9:00-10:00 |

`doPoll` **больше не нужно ставить на расписание** — опрос теперь делает
GitHub Actions. `doPoll` остался в коде как ручной способ проверить всё то
же самое из меню таблицы, без GitHub.

`onEdit` — простой триггер, Apps Script подхватывает по имени функции сам.

## Рабочий процесс обновления кода

```bash
git add . && git commit -m "что поменялось"
git push

clasp push                     # если менялся Code.gs
# + если менялся doPost — Deploy -> Manage deployments -> New version
```

`poll_bms.py` и workflow подхватываются GitHub Actions сами при пуше, ничего
доп. заливать не нужно.

## Меню в таблице (после push и перезагрузки страницы)

- 🔑 Войти в BMS
- ▶ Запустить проверку новых мероприятий (`doPoll` вручную, без GitHub)
- 🧪 Тест: заполнить "Тех блок" по текущим данным
- 🧪 Тест: записать на доску (event по умолчанию — поменять ID в коде)
- ▶ Записать на доску дату...
- 📋 Обновить график персонала сейчас
- 🔬 Дамп структуры выделенных ячеек (для отладки доски)
- 🧹 Сбросить реестр положений на доске
