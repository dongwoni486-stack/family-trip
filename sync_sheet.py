# -*- coding: utf-8 -*-
"""
sync_sheet.py
-----------------------------------------------------------------
구글 스프레드시트("오사카_교토_여행일정_공유용" - 새 버전)의 '오사카_교토' 시트를
읽어 index.html의 ITINERARY_AUTO_START~END 구간을 자동으로 다시 생성합니다.

시트 컬럼 구조 (A~H):
  A 일차 및 날짜 | B 시간 | C 제목 | D 장소(@스마트칩) | E 설명 | F 후보(맛집/대안) | G 체크리스트 | H 바우처URL

사용 전 준비물 (본인 구글 계정에서 직접 진행):
  1. Google Cloud Console에서 프로젝트 생성 (이미 있다면 재사용)
  2. "Google Sheets API" 활성화 (지도/장소 표시는 이미 있는
     Maps JavaScript API 브라우저 키를 그대로 쓰고, 이 서버 스크립트는
     별도의 "서버 전용 API 키"를 새로 만들어 쓰는 걸 권장합니다 —
     브라우저 파일 안에 절대 노출되지 않으므로 훨씬 안전합니다.)
  3. 서비스 계정 생성 → JSON 키 다운로드 → 이 파일과 같은 폴더에 credentials.json 으로 저장
  4. 서비스 계정 이메일(xxx@xxx.iam.gserviceaccount.com)을
     구글 스프레드시트 "공유"에 뷰어 권한으로 추가
  5. (선택) Geocoding API로 좌표를 자동 갱신하려면 GOOGLE_MAPS_API_KEY 환경변수에
     "서버 전용" 키를 설정하세요. 설정하지 않으면 기존 index.html의 좌표를 그대로 둡니다.
  6. ⚠️ 과금 방지: Google Cloud Console > API 및 서비스 > 할당량에서
     Geocoding API 하루 호출 한도를 낮게(예: 하루 300회) 설정해두세요.

설치:
    pip install gspread google-auth requests --break-system-packages

실행:
    python sync_sheet.py
"""

import json
import os
import re
import sys
import time

try:
    import gspread
    from google.oauth2.service_account import Credentials
except ImportError:
    print("필요한 패키지가 없습니다. 다음 명령으로 설치해주세요:")
    print("  pip install gspread google-auth requests --break-system-packages")
    sys.exit(1)

# ------------------------------------------------------------------
# 설정
# ------------------------------------------------------------------
# URL의 /d/ 와 /edit 사이 값입니다.
# 다음 여행에는 코드를 고치지 말고, GitHub 시크릿 SPREADSHEET_ID 값만 새 시트 ID로 바꾸세요.
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1yPOv4d9XSQIdyTKcc1T8A2PnTP3oExwDaZrMyZUtW0A")

# 시트 탭 이름 — 앞으로 만드는 모든 여행 시트에서 이 이름을 그대로 써주세요.
# (다음 여행에도 탭 이름만 똑같이 맞추면 이 파일은 손댈 필요가 없습니다.)
SCHEDULE_SHEET_NAME = os.environ.get("SCHEDULE_SHEET_NAME", "일정")
CREDENTIALS_PATH = "credentials.json"
OUTPUT_PATH = "index.html"
START_MARKER = "/* ITINERARY_AUTO_START */"
END_MARKER = "/* ITINERARY_AUTO_END */"
REF_START_MARKER = "/* REF_AUTO_START */"
REF_END_MARKER = "/* REF_AUTO_END */"
# 참고 목록(맛집/쇼핑/관광지) 탭 이름. 이것도 매 여행 시트에서 동일하게 맞춰주세요.
REF_SHEET_TABS = [
    ("맛집", "food"),
    ("쇼핑", "shop"),
    ("관광지", "sight"),
]
GOOGLE_MAPS_API_KEY = os.environ.get("GOOGLE_MAPS_API_KEY", "")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]


def open_spreadsheet():
    """
    인증 우선순위:
      1) GOOGLE_CREDENTIALS_JSON 환경변수 (GitHub Actions 시크릿) — 서비스 계정 키 JSON 전체 문자열
      2) credentials.json 파일 (로컬 실행용)
    """
    env_creds = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")
    if env_creds:
        try:
            info = json.loads(env_creds)
        except json.JSONDecodeError:
            sys.exit("GOOGLE_CREDENTIALS_JSON 값이 올바른 JSON이 아닙니다.")
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    elif os.path.exists(CREDENTIALS_PATH):
        creds = Credentials.from_service_account_file(CREDENTIALS_PATH, scopes=SCOPES)
    else:
        print(f"인증 정보가 없습니다. GOOGLE_CREDENTIALS_JSON 환경변수를 설정하거나")
        print(f"'{CREDENTIALS_PATH}' 파일을 준비해주세요.")
        sys.exit(1)

    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID)


def geocode(place_name):
    if not GOOGLE_MAPS_API_KEY or not place_name:
        return None
    import requests
    try:
        resp = requests.get(
            "https://maps.googleapis.com/maps/api/geocode/json",
            params={"address": place_name, "key": GOOGLE_MAPS_API_KEY},
            timeout=10,
        ).json()
        if resp.get("status") == "OK" and resp.get("results"):
            loc = resp["results"][0]["geometry"]["location"]
            return loc["lat"], loc["lng"]
    except Exception:
        pass
    return None


def get_directions(origin, destination, mode, transit_mode=None):
    """서버 쪽에서 딱 한 번 경로를 계산해 결과를 저장합니다.
    브라우저는 이 결과를 그리기만 하므로 사용자가 앱을 몇 번을 열어도
    Directions API가 추가로 호출되지 않습니다.
    transit_mode: mode가 'transit'일 때만 'train' 또는 'bus'로 세분화 가능."""
    if not GOOGLE_MAPS_API_KEY:
        return None
    import requests
    try:
        params = {
            "origin": f"{origin[0]},{origin[1]}",
            "destination": f"{destination[0]},{destination[1]}",
            "mode": mode,
            "key": GOOGLE_MAPS_API_KEY,
            "language": "ko",
            "region": "jp",
        }
        if mode == "transit" and transit_mode:
            params["transit_mode"] = transit_mode
        resp = requests.get(
            "https://maps.googleapis.com/maps/api/directions/json",
            params=params,
            timeout=10,
        ).json()
        if resp.get("status") == "OK" and resp.get("routes"):
            route = resp["routes"][0]
            leg = route["legs"][0]
            return {
                "distance": leg["distance"]["text"],
                "duration": leg["duration"]["text"],
                "poly": route["overview_polyline"]["points"],
            }
    except Exception:
        pass
    return None


def load_ref_sheet(sh, tab_name):
    """참고 목록 탭 하나를 읽어옵니다. 탭이 없으면 None을 반환하고 건너뜁니다."""
    try:
        ws = sh.worksheet(tab_name)
    except Exception:
        print(f"  - '{tab_name}' 탭을 찾지 못해 건너뜁니다 (탭 이름을 확인해주세요).")
        return None
    rows = ws.get_all_values()
    items = []
    last_region = ""
    for row in rows[1:]:
        row = row + [""] * (5 - len(row))
        region, cat, name, note = [str(x).strip() for x in row[:4]]
        region = region if region else last_region
        last_region = region
        if not name:
            continue
        items.append({"region": region, "cat": cat, "name": name, "note": note})
    return items


def geocode_ref_items(items):
    for it in items:
        query = f"{it['name']} {it.get('region', '')}".strip()
        latlng = geocode(query)
        if latlng is not None:
            it["lat"], it["lng"] = latlng
        time.sleep(0.2)


def parse_rows_to_days(rows):
    """A~H 8개 컬럼을 읽어 day 단위 리스트로 묶습니다.
    병합된 '일차' 셀은 gspread가 첫 행에만 값을 주고 나머지는 빈 문자열을
    반환하므로, 마지막으로 본 day_label을 이어받습니다."""
    days = []
    current_day = None
    last_day_label = ""

    for row in rows[1:]:  # 헤더 제외
        row = row + [""] * (8 - len(row))
        day_label, time_val, title, place, desc, cands, checklist, url = [
            str(x).strip() for x in row[:8]
        ]
        if not title and not place and not desc:
            continue

        day_label = day_label if day_label else last_day_label
        last_day_label = day_label

        if current_day is None or current_day["dayLabel"] != day_label:
            current_day = {"dayLabel": day_label, "checklist": [], "stops": []}
            days.append(current_day)

        if checklist:
            current_day["checklist"].extend(
                [item.strip() for item in checklist.split(",") if item.strip()]
            )

        stop = {"time": time_val, "title": title, "place": place, "desc": desc, "url": url}

        if cands:
            names = [n.strip() for n in re.findall(r"[-•]\s*([^\n]+)", cands) if n.strip()]
            if not names:
                names = [cands]
            stop["candidates"] = [{"name": n} for n in names]

        current_day["stops"].append(stop)

    for d in days:
        d["checklist"] = list(dict.fromkeys(d["checklist"]))  # 중복 제거, 순서 유지

    return days


def main():
    print("1. 구글 시트에 접속하는 중...")
    sh = open_spreadsheet()

    print("2. 일정 탭을 불러와 정리하는 중...")
    rows = sh.worksheet(SCHEDULE_SHEET_NAME).get_all_values()
    days = parse_rows_to_days(rows)

    ref_data = {"food": [], "shop": [], "sight": []}

    if GOOGLE_MAPS_API_KEY:
        print("3. Geocoding API로 일정 좌표를 갱신하는 중... (과호출 방지를 위해 0.2초 간격)")
        for d in days:
            for s in d["stops"]:
                target = s.get("place") or s.get("title")
                latlng = geocode(target)
                if latlng is not None:
                    s["lat"], s["lng"] = latlng
                time.sleep(0.2)

        print("4. 구간별 이동 경로(도보/기차/버스/택시)를 미리 계산하는 중...")
        # key: (구글 mode, transit_mode 또는 None)
        MODE_MAP = {
            "walking": ("walking", None),
            "train": ("transit", "train"),
            "bus": ("transit", "bus"),
            "driving": ("driving", None),
        }
        for d in days:
            legs = []
            stops = d["stops"]
            for i in range(len(stops) - 1):
                a, b = stops[i], stops[i + 1]
                leg_result = {}
                if all(k in a for k in ("lat", "lng")) and all(k in b for k in ("lat", "lng")):
                    origin = (a["lat"], a["lng"])
                    destination = (b["lat"], b["lng"])
                    for key, (gmode, tmode) in MODE_MAP.items():
                        r = get_directions(origin, destination, gmode, tmode)
                        if r:
                            leg_result[key] = r
                        time.sleep(0.2)
                legs.append(leg_result)
            d["legs"] = legs

        print("5. 맛집/쇼핑/관광지 참고 목록을 불러와 좌표를 계산하는 중... (시간이 걸릴 수 있습니다)")
        for tab_name, key in REF_SHEET_TABS:
            items = load_ref_sheet(sh, tab_name)
            if items is None:
                continue
            print(f"  - '{tab_name}' {len(items)}개 항목 좌표 계산 중...")
            geocode_ref_items(items)
            ref_data[key] = items
    else:
        print("3~5. GOOGLE_MAPS_API_KEY가 없어 좌표/경로/참고목록 계산은 건너뜁니다 (기존 값 유지).")

    print("6. index.html 파일을 갱신하는 중...")
    if not os.path.exists(OUTPUT_PATH):
        sys.exit(f"'{OUTPUT_PATH}' 파일이 없습니다. 먼저 받은 index.html을 이 폴더에 넣어주세요.")

    with open(OUTPUT_PATH, "r", encoding="utf-8") as f:
        html = f.read()

    if START_MARKER not in html or END_MARKER not in html:
        sys.exit("일정 마커를 찾을 수 없습니다. index.html이 원본 그대로인지 확인해주세요.")

    # JSON 직렬화 — 문자열 수동 조립보다 안전 (따옴표/이스케이프 문제 원천 차단)
    js_data = f"const itinerary = {json.dumps(days, ensure_ascii=False, indent=2)};"
    start_idx = html.index(START_MARKER) + len(START_MARKER)
    end_idx = html.index(END_MARKER)
    html = html[:start_idx] + "\n" + js_data + "\n" + html[end_idx:]

    if GOOGLE_MAPS_API_KEY and REF_START_MARKER in html and REF_END_MARKER in html:
        ref_js = f"const refData = {json.dumps(ref_data, ensure_ascii=False, indent=2)};"
        r_start = html.index(REF_START_MARKER) + len(REF_START_MARKER)
        r_end = html.index(REF_END_MARKER)
        html = html[:r_start] + "\n" + ref_js + "\n" + html[r_end:]

    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        f.write(html)

    total_stops = sum(len(d["stops"]) for d in days)
    total_ref = sum(len(v) for v in ref_data.values())
    print(f"완료: index.html 갱신됨 (일정 {len(days)}일 {total_stops}개, 참고목록 {total_ref}개)")


if __name__ == "__main__":
    main()
