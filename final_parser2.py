import re
import time
import csv
import os
import asyncio
from openpyxl import Workbook
import browser_cookie3
from urllib.parse import urlparse, urljoin

import requests
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

# ---------------- CONFIG ----------------
BASE = "https://dec.srspu.ru"
LOGIN_URL = f"{BASE}/Account/Login.aspx?ReturnUrl=%2fVed%2f"
DISCIPLINES_URL = f"{BASE}/Ved/"

LOGIN = ""
PASSWORD = ""

OUT_DIR = "ved_results"
HTML_DIR = os.path.join(OUT_DIR, "html")
CSV_PATH = os.path.join(OUT_DIR, "all_students.xlsx")

HEADLESS = True
REQUESTS_SLEEP = 0.5
TIMEOUT_NAV = 30000
# ----------------------------------------

os.makedirs(HTML_DIR, exist_ok=True)
print(f"[+] output -> {OUT_DIR}")


# # ---------------- Playwright Login ----------------
# def playwright_login_and_get_cookies(login, password):
#     print("[*] Запускаем Playwright и логинимся...")
#     with sync_playwright() as p:
#         browser = p.firefox.launch(headless=HEADLESS)
#         context = browser.new_context()
#         page = context.new_page()

#         page.goto(LOGIN_URL)
#         print("    -> страница логина загружена")

#         page.eval_on_selector(
#             "input[name='ctl00$MainContent$ucLoginFormPage$tbUserName']",
#             f"el => el.value = '{login}'"
#         )
#         page.eval_on_selector(
#             "input[name='ctl00$MainContent$ucLoginFormPage$tbUserName']",
#             "el => el.dispatchEvent(new Event('change'))"
#         )

#         page.eval_on_selector(
#             "input[name='ctl00$MainContent$ucLoginFormPage$tbPassword']",
#             f"el => el.value = '{password}'"
#         )
#         page.eval_on_selector(
#             "input[name='ctl00$MainContent$ucLoginFormPage$tbPassword']",
#             "el => el.dispatchEvent(new Event('change'))"
#         )

#         try:
#             page.click("#ctl00_MainContent_ucLoginFormPage_btnLogin", timeout=10000)
#         except Exception:
#             page.eval_on_selector(
#                 "input[name='ctl00$MainContent$ucLoginFormPage$btnLogin']",
#                 "el => el.click()"
#             )

#         try:
#             page.wait_for_url("**/Ved/**", timeout=TIMEOUT_NAV)
#             print("    -> редирект на /Ved/ подтверждён")
#         except Exception:
#             print("    -> ожидание редиректа превысило таймаут. Продолжаем...")

#         cookies = context.cookies()
#         browser.close()
#         print(f"    -> получено куки: {[c['name'] for c in cookies]}")
#         return cookies

# ---------------- Browser Cookie Loader ----------------
def get_cookiejar_for_domain(domain: str):
    """Пытается получить cookiejar из популярных локальных браузеров."""
    browsers = [
        ("chrome", browser_cookie3.chrome),
        ("chromium", browser_cookie3.chromium),
        ("edge", browser_cookie3.edge),
        ("firefox", browser_cookie3.firefox),
        ("opera", browser_cookie3.opera),
    ]
    for name, fn in browsers:
        try:
            cj = fn(domain_name=domain)
            if cj and len(list(cj)) > 0:
                print(f"[+] Cookies найдены в {name}")
                return cj
            else:
                print(f"[-] {name}: куки не найдены")
        except Exception as e:
            print(f"[-] {name}: ошибка при чтении куков ({e})")
    print("[!] Куки не найдены ни в одном браузере.")
    return None


def inject_cookiejar_into_session(session, cj, domain: str):
    """Переносит куки из cookiejar (browser_cookie3) в requests.Session."""
    for cookie in cj:
        if domain in cookie.domain:
            session.cookies.set(cookie.name, cookie.value, domain=cookie.domain, path=cookie.path)
    print("[+] Куки из браузера подставлены в requests.Session()")


def fetch_cookies_from_cookie_server(api_url: str, api_key: str, login: str, password: str):
    """Запрашивает cookies у удалённого cookie-server."""
    url = api_url.rstrip("/") + "/get_cookies"
    headers = {"X-API-Key": api_key}
    payload = {"login": login, "password": password}
    try:
        r = requests.post(url, json=payload, headers=headers, timeout=60)
        r.raise_for_status()
        j = r.json()
        return j.get("cookies", [])
    except Exception as e:
        print(f"[-] Ошибка при запросе к cookie-server: {e}")
        return None


def transfer_cookies_from_playwright_format(session: requests.Session, cookies: list):
    """Подставляем Playwright куки в requests.Session."""
    for c in cookies:
        name = c.get("name")
        value = c.get("value")
        if not name or value is None:
            continue
        # Игнорируем domain, path, secure для простоты
        session.cookies.set(name, value, domain="dec.srspu.ru", path="/", secure=False)
    print("[+] Cookies подставлены в requests.Session()")


# ---------------- Requests Session ----------------
def transfer_cookies_to_requests(session, cookies):
    for c in cookies:
        session.cookies.set(
            c["name"],
            c["value"],
            domain=c.get("domain", urlparse(BASE).hostname),
            path=c.get("path", "/"),
        )
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:146.0) Gecko/20100101 Firefox/146.0",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3",
    })
    print("    -> cookies установлены в requests.Session()")


# ---------------- Парсинг ----------------
def get_first_available_discipline(html):
    soup = BeautifulSoup(html, "html.parser")
    a = soup.find("a", class_="dxeHyperlink_MaterialCompact", href=re.compile(r"Ved\.aspx\?id=\d+"))
    if a:
        href = a["href"]
        m = re.search(r"id=(\d+)", href)
        if m:
            did = m.group(1)
            name = a.get_text(strip=True)
            full_url = urljoin(DISCIPLINES_URL, href)
            return did, name, full_url
    return None, None, None


def extract_group_from_discipline_page(html):
    soup = BeautifulSoup(html, "html.parser")
    a = soup.find("a", id="ctl00_MainContent_ucVedBox_lblGroup")
    if a and "href" in a.attrs:
        href = a["href"]
        m = re.search(r"id=(\d+)", href)
        if m:
            gid = m.group(1)
            group_name = a.get_text(strip=True)
            return gid, group_name
    return None, None


def fetch_group_page(session, group_id):
    url = urljoin(BASE, f"/Dek/?mode=stud&f=group&id={group_id}")
    print(f"[*] GET {url}")
    r = session.get(url, timeout=30)
    r.raise_for_status()
    return r.text


def extract_student_ids_and_names(html):
    soup = BeautifulSoup(html, "html.parser")
    result = {}
    for a in soup.find_all("a", href=True):
        m = re.search(r"[?&]id=(\d+)", a["href"])
        if m:
            sid = m.group(1)
            txt = a.get_text(strip=True)
            if txt.isdigit() and int(sid) > 1000:
                result[sid] = txt
    print(f"    -> найдено студентов: {len(result)}")
    return result


def fetch_totalved_for_student(session, student_id):
    url = urljoin(BASE, f"/Ved/TotalVed.aspx?year=cur&sem=cur&id={student_id}")
    headers = {"Referer": DISCIPLINES_URL}
    print(f"    -> GET {url}")
    r = session.get(url, headers=headers, timeout=30)
    r.raise_for_status()
    return r.text


def parse_totalved_discipline_scores(html):
    soup = BeautifulSoup(html, "html.parser")
    rows = []
    for tr in soup.find_all("tr", class_="dxgvDataRow_MaterialCompact"):
        tds = tr.find_all("td")
        if len(tds) >= 5:
            a_tag = tds[1].find("a")
            discipline = a_tag.get_text(strip=True) if a_tag else tds[1].get_text(strip=True)
            score = tds[4].get_text(strip=True)
            rows.append((discipline, score))
    return rows


# ---------------- Main ----------------
def main():
    login = LOGIN or input("Login: ").strip()
    password = PASSWORD or __import__("getpass").getpass("Password: ")

    s = requests.Session()

    # 1) Пытаемся достать куки из браузера
    cj = get_cookiejar_for_domain("dec.srspu.ru")
    if cj:
        inject_cookiejar_into_session(s, cj, "dec.srspu.ru")
    else:
        # если локально не нашлось — запросим у удалённого сервиса Playwright
        API_URL = "http://89.169.12.12:63592"   # <- адрес твоего cookie-server
        API_KEY = "transfer_train_never_been_located"
        pw_cookies = fetch_cookies_from_cookie_server(API_URL, API_KEY, login, password)
        if pw_cookies:
            transfer_cookies_from_playwright_format(s, pw_cookies)
        else:
            print("[!] Не удалось получить cookies ни локально, ни с cookie-server.")
            return

    # 2) Если нет — логинимся через Pyppeteer
    # if not cj:
    #     cookies = playwright_login_and_get_cookies(login, password)
    #     transfer_cookies_to_requests(s, cookies)

    # 3) Работаем как раньше
    r = s.get(DISCIPLINES_URL)
    r.raise_for_status()
    discipline_id, discipline_name, discipline_url = get_first_available_discipline(r.text)
    if not discipline_id:
        print("[!] Нет доступных дисциплин.")
        return
    print(f"[*] Первая дисциплина: {discipline_name} (id={discipline_id})")

    r_disc = s.get(discipline_url)
    r_disc.raise_for_status()
    group_id, group_name = extract_group_from_discipline_page(r_disc.text)
    if not group_id:
        print("[!] Не удалось получить group_id.")
        return
    print(f"[*] Группа: {group_name} (id={group_id})")

    group_html = fetch_group_page(s, group_id)
    students = extract_student_ids_and_names(group_html)

    rows_for_csv = []
    for sid, student_name in students.items():
        try:
            ved_html = fetch_totalved_for_student(s, sid)
        except Exception as e:
            print(f"    !! Ошибка при загрузке {sid}: {e}")
            continue

        html_file = os.path.join(HTML_DIR, f"student_{sid}.html")
        with open(html_file, "w", encoding="utf-8") as f:
            f.write(ved_html)
        print(f"    -> сохранено {html_file}")

        discipline_scores = parse_totalved_discipline_scores(ved_html)
        if discipline_scores:
            for discipline, score in discipline_scores:
                rows_for_csv.append([sid, student_name, discipline, score])
        else:
            rows_for_csv.append([sid, student_name, "NO_DISCIPLINES", ""])

        time.sleep(REQUESTS_SLEEP)

    if rows_for_csv:
        xlsx_path = os.path.join(OUT_DIR, "all_students.xlsx")

        wb = Workbook()
        ws = wb.active
        ws.append(["student_id", "student_name", "discipline", "score"])

        for row in rows_for_csv:
            ws.append(row)

        wb.save(xlsx_path)
        print(f"[+] XLSX сохранён: {xlsx_path}")
        
    else:
        print("[!] Нечего сохранять в XLSX.")

    with open(CSV_PATH, "w", newline="", encoding="utf-8") as csvf:
        writer = csv.writer(csvf)
        writer.writerow(header)
        writer.writerows(rows_for_csv)

    print("[*] Готово.")


if __name__ == "__main__":
    main()
