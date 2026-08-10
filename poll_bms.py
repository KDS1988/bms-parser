"""
BMS poller для GitHub Actions.

Логинится в BMS собственным refresh_token (независимо от Apps Script),
тянет мероприятия с нужными услугами, отправляет их в Apps Script по HTTP
(там уже вся логика записи на доску/в "Тех блок"/реестры), и на основе
ответа шлёт уведомления в Telegram.

Секреты — только из переменных окружения (GitHub Secrets), в коде их нет.
"""
import os
import sys
import time
from datetime import date, timedelta

import requests

BMS_API = "https://bms-api.vsporte.ru/api/v1/bms"
TARGET_SERVICES = ["Прямая трансляция", "Студия"]
LOOKAHEAD_DAYS = 60


def refresh_bms_token(refresh_token_value: str) -> tuple[str, str]:
    r = requests.post(f"{BMS_API}/users/token/refresh", json={"refresh_token": refresh_token_value}, timeout=30)
    r.raise_for_status()
    data = r.json()
    return data["access_token"], data["refresh_token"]


def fetch_assignments(access_token: str, date_from: str, date_to: str) -> list[dict]:
    items: list[dict] = []
    page = 1
    while True:
        r = requests.get(
            f"{BMS_API}/projects/assignment_table/technical",
            headers={"Authorization": f"Bearer {access_token}"},
            params={
                "page": page, "size": 100,
                "order_by": "+date,start_time,id",
                "date__gte": date_from, "date__lte": date_to,
            },
            timeout=60,
        )
        r.raise_for_status()
        data = r.json()
        items.extend(data["items"])
        if page >= data.get("pages", 1):
            break
        page += 1
        time.sleep(0.2)
    return items


def group_by_event(items: list[dict]) -> list[dict]:
    groups: dict[int, dict] = {}
    order: list[int] = []
    for item in items:
        ev = item["event"]
        key = ev["id"]
        if key not in groups:
            groups[key] = {"event": ev, "items": []}
            order.append(key)
        groups[key]["items"].append(item)
    return [groups[k] for k in order]


def send_telegram(token: str, chat_id: str, text: str) -> None:
    resp = requests.post(
        f"https://api.telegram.org/bot{token}/sendMessage",
        json={"chat_id": chat_id, "text": text, "parse_mode": "HTML"},
        timeout=30,
    )
    if not resp.ok:
        print(f"Telegram error: {resp.status_code} {resp.text}", file=sys.stderr)


def format_event_message(event: dict) -> str:
    sport = (event.get("sport_type") or {}).get("name", "")
    league = (event.get("league") or {}).get("short_name", "")
    home = event.get("home_team")
    away = event.get("away_team")
    teams = f"{home['name']} — {away['name']}" if home and away else event.get("event_description", "")
    location = (event.get("location") or {}).get("name") or event.get("venue", "")

    lines = [
        "🆕 <b>Новое мероприятие в BMS</b>",
        f"{event.get('date', '')} {event.get('start_time', '')}".strip(),
        " / ".join(x for x in [sport, league] if x),
        teams,
    ]
    if location:
        lines.append(f"📍 {location}")
    return "\n".join(x for x in lines if x)


def format_change_message(event: dict, service_name: str, changes: list[str]) -> str:
    home = event.get("home_team")
    away = event.get("away_team")
    teams = f"{home['name']} — {away['name']}" if home and away else event.get("event_description", "")
    lines = [
        "🔄 <b>Изменение в мероприятии</b>",
        f"{event.get('date', '')} {event.get('start_time', '')} {teams}".strip(),
        service_name,
        *changes,
    ]
    return "\n".join(x for x in lines if x)


def main() -> None:
    refresh_val = os.environ["BMS_REFRESH_TOKEN"]
    telegram_token = os.environ["TELEGRAM_BOT_TOKEN"]
    telegram_chat_id = os.environ["TELEGRAM_CHAT_ID"]
    webhook_url = os.environ["APPS_SCRIPT_WEBHOOK_URL"]
    webhook_secret = os.environ["WEBHOOK_SECRET"]

    access_token, new_refresh = refresh_bms_token(refresh_val)

    # сохраняем новый refresh_token сразу — даже если что-то ниже упадёт,
    # токен всё равно не потеряется (иначе на следующий запуск придёт уже невалидный)
    os.makedirs("state", exist_ok=True)
    with open("state/refresh_token.txt", "w") as f:
        f.write(new_refresh)

    today = date.today()
    date_from = today.isoformat()
    date_to = (today + timedelta(days=LOOKAHEAD_DAYS)).isoformat()

    items = fetch_assignments(access_token, date_from, date_to)
    entries = group_by_event(items)

    payload_events = []
    for entry in entries:
        target_items = [it for it in entry["items"] if (it.get("service") or {}).get("name") in TARGET_SERVICES]
        if target_items:
            payload_events.append({"event": entry["event"], "items": target_items})

    if not payload_events:
        print("Нет мероприятий с нужными услугами в горизонте.")
        return

    resp = requests.post(
        webhook_url,
        json={"secret": webhook_secret, "events": payload_events},
        timeout=180,
    )
    resp.raise_for_status()
    result = resp.json()
    if not result.get("ok"):
        print("Webhook error:", result.get("error"), file=sys.stderr)
        sys.exit(1)

    events_by_id = {e["event"]["id"]: e["event"] for e in payload_events}
    for r in result.get("results", []):
        event = events_by_id.get(r["eventId"])
        if event is None:
            continue

        if r["isNewEvent"]:
            send_telegram(telegram_token, telegram_chat_id, format_event_message(event))

        for svc in r.get("services", []):
            if svc.get("error"):
                print(f"Ошибка записи на доску event_id={r['eventId']}: {svc['error']}", file=sys.stderr)
                continue
            if r["isNewEvent"]:
                continue  # уже уведомили про мероприятие целиком выше

            if svc.get("isNew"):
                home = event.get("home_team")
                teams = home["name"] if home else event.get("event_description", "")
                send_telegram(
                    telegram_token, telegram_chat_id,
                    f"➕ В мероприятии {event.get('date', '')} {teams} добавлена услуга \"{svc.get('serviceName', '')}\"",
                )
            elif svc.get("changes"):
                send_telegram(
                    telegram_token, telegram_chat_id,
                    format_change_message(event, svc.get("serviceName", ""), svc["changes"]),
                )

    print(f"Готово. Обработано мероприятий: {len(payload_events)}.")


if __name__ == "__main__":
    main()
