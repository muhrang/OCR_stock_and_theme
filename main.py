# runner_onefile.py
# ✅ 단일 파일 (XING=Python 3.14-32 실행)
# ✅ OCR은 Python 3.12-64 subprocess로 mapping.json 생성(코드->테마)
# ✅ 화면: [거래대금상위] | [등락률상위] | [교집합+양봉]
# ✅ 추가: mapping.json 전체 종목을 rate>=5%만 "테마 패널"로 옆으로 출력

import os
import time
import json
import subprocess
import tempfile
import textwrap
import unicodedata

import pythoncom
import win32com.client
from dataclasses import dataclass


# =========================
# OCR용 Python(3.12) 지정
# =========================
PY312_EXE = r"C:\Users\User\AppData\Local\Programs\Python\Python312\python.exe"
USE_PY_LAUNCHER = False  # True면 ["py","-3.12"]


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MAPPING_PATH = os.path.join(BASE_DIR, "mapping.json")
DEFAULT_THEME = "미분류"


# =========================
# XING 설정
# =========================
@dataclass
class XingConfig:
    user_id: str = os.environ.get("XING_USER_ID", "")
    user_pw: str = os.environ.get("XING_USER_PW", "")
    cert_pw: str = os.environ.get("XING_CERT_PW", "")
    server: str = os.environ.get("XING_SERVER", "real")  # real/demo
    timeout_sec: int = 12

CFG = XingConfig()

RES_DIR = r"C:\xingAPI_Program(2025.06.07)\Res"
SERVER_ADDR = {"real": "hts.ebestsec.co.kr", "demo": "demo.ebestsec.co.kr"}

TOP_MONEY = 50
TOP_RATE = 30
REFRESH_SEC = 40
PRINT_MIN_RATE = 5.0

# t8407 배치 크기(보통 50~100 사이가 안전)
T8407_BATCH = 50


# =========================
# 공용 유틸
# =========================
def sstrip(x) -> str:
    return (x or "").strip()

def to_float_or_none(x):
    s = sstrip(x).replace(",", "")
    if s == "" or s == "-":
        return None
    try:
        return float(s)
    except Exception:
        return None

def to_int_or_none(x):
    s = sstrip(x).replace(",", "")
    if s == "" or s == "-":
        return None
    try:
        return int(float(s))
    except Exception:
        return None

def fmt_rate(rate):
    return "?%" if rate is None else f"{rate:.2f}%"

def clear_screen():
    os.system("cls")

def disp_width(s: str) -> int:
    w = 0
    for ch in s:
        ea = unicodedata.east_asian_width(ch)
        w += 2 if ea in ("F", "W") else 1
    return w

def ljust_disp(s: str, width: int) -> str:
    pad = width - disp_width(s)
    return s if pad <= 0 else s + (" " * pad)

def center_disp(s: str, width: int) -> str:
    w = disp_width(s)
    if w >= width:
        return s
    left = (width - w) // 2
    right = width - w - left
    return (" " * left) + s + (" " * right)

def sort_by_rate_desc(rows):
    def key(r):
        v = r.get("rate")
        return (-1e18 if v is None else v)
    return sorted(rows, key=key, reverse=True)

def apply_min_rate_filter(rows, min_rate):
    if min_rate is None:
        return rows or []
    out = []
    for r in rows or []:
        v = r.get("rate")
        if v is not None and v >= min_rate:
            out.append(r)
    return out

def build_panel_lines(title, rows, min_rate=None):
    body = []
    if rows:
        for r in rows:
            v = r.get("rate")
            if min_rate is not None and (v is None or v < min_rate):
                continue
            body.append((r["name"], fmt_rate(v)))
    if not body:
        body = [("(없음)", "")]

    name_w = max(disp_width(n) for n, _ in body)
    rate_w = max(disp_width(rt) for _, rt in body)
    width = max(disp_width(title), name_w + 2 + rate_w)

    lines = [center_disp(title, width), "=" * width]
    for name, rt in body:
        line = f"{ljust_disp(name, name_w)}  {rt.rjust(rate_w)}"
        lines.append(ljust_disp(line, width))
    return lines, width

def print_panels_side_by_side(panel_infos, gap=" | "):
    panels = [p for p, _ in panel_infos]
    widths = [w for _, w in panel_infos]
    max_len = max(len(p) for p in panels) if panels else 0

    for i in range(max_len):
        cells = []
        for col, p in enumerate(panels):
            cells.append(p[i] if i < len(p) else (" " * widths[col]))

        last = -1
        for j in range(len(cells) - 1, -1, -1):
            if cells[j].strip():
                last = j
                break
        if last == -1:
            continue
        print(gap.join(cells[:last + 1]))


# =========================
# mapping.json 로드 (코드->테마)
# =========================
def load_mapping_code_to_theme():
    """
    기대 포맷:
    {
      "themes": ["로봇",...],
      "map": {"005930":"반도체", ...}
    }
    """
    try:
        with open(MAPPING_PATH, "r", encoding="utf-8") as f:
            j = json.load(f)
        if isinstance(j, dict) and isinstance(j.get("map"), dict):
            mp = {}
            for k, v in j["map"].items():
                kk = str(k).strip()
                vv = str(v).strip() if v is not None else DEFAULT_THEME
                if kk.isdigit() and len(kk) == 6:
                    mp[kk] = vv or DEFAULT_THEME
            return mp
    except Exception:
        pass
    return {}

def group_rows_by_theme(rows, code_to_theme: dict):
    buckets = {}
    for r in rows or []:
        code = str(r.get("code","")).strip()
        theme = code_to_theme.get(code, DEFAULT_THEME)
        buckets.setdefault(theme, []).append(r)
    for t in buckets:
        buckets[t] = sort_by_rate_desc(buckets[t])
    return buckets


# =========================
# OCR(3.12) subprocess: mapping.json 생성 (코드->테마)
# =========================
def _get_py312_cmd():
    if USE_PY_LAUNCHER:
        return ["py", "-3.12"]
    if not os.path.exists(PY312_EXE):
        raise RuntimeError(f"PY312 경로가 잘못됨: {PY312_EXE}")
    return [PY312_EXE]

def run_ocr_with_py312_make_mapping():
    pycmd = _get_py312_cmd()

    # ❗f-string 금지(중괄호 안정)
    ocr_template = r"""
import cv2
import numpy as np
import easyocr
import re
import FinanceDataReader as fdr
import Levenshtein
import json
from tkinter import Tk, filedialog

MAPPING_PATH = r"__MAPPING_PATH__"
DEFAULT_THEME = "__DEFAULT_THEME__"
THEME_POOL = [
  "로봇","휴머노이드","의료로봇","물류로봇",
  "로봇부품","로봇감속기","공장자동화","스마트팩토리",
  "이재명","정치테마","대선","지역화폐","주택","부동산",
  "전기차","차량용반도체","자율주행",
  "전력반도체","전력",
  "항암","비만치료제","mRNA","백신","RNA치료","유전자치료","세포치료","CAR-T",
  "의료기기","진단키트",
  "헬스케어","의료AI","원격",
  "디스플레이","LCD","OLED","마이크로LED","플렉서블디스플레이",
  "VR","메타버스",
  "반도체","시스템반도체","메모리반도체","비메모리","파운드리",
  "AI","온디바이스AI","AI반도체","AI서버",
  "데이터센터","빅데이터","클라우드","양자컴퓨터",
  "보안","블록체인","IoT","스마트시티",
  "스마트폰","스마트폰부품","모바일부품","카메라","카메라모듈",
  "폴더블폰","힌지","터치패널","강화유리","OLED","스페이스","통신칩","안테나",
  "스마트폰배터리","충전기","모바일OS","안드로이드",
  "스페이스","통신장비","5G","6G",
  "2차전지","전고체","리튬","니켈","코발트","망간",
  "음극재","양극재","전해질","분리막","배터리장비","배터리재활용","ESS",
  "수소","수소연료전지","태양광","풍력","원전","SMR","풍력발전","태양광발전",
  "전력설비","스마트그리드","탄소중립","탄소포집","탄소배출권",
  "바이오","제약","바이오시밀러","마이크로바이옴","재생의료","줄기세포",
  "당뇨","치매","희귀질환","신약개발",
  "자동차","자동차부품","라이다","레이더","전기선박","드론",
  "우주","우주항공","항공우주","민간우주","우주산업","우주개발",
  "위성","소형위성","정찰위성",
  "로켓","재사용로켓","항공",
  "조선","조선기자재","LNG","LPG",
  "친환경선박","이중연료엔진","선박엔진",
  "해양플랜트","해저케이블","해저자원",
  "방산","국방","미사일",
  "건설","재건","철강","구리","희토류","철도","남북경협",
  "2차전지","스마트폰","자동차","우주",
  "기계","화학",
  "고순도소재","세라믹","나노소재","그래핀","탄소섬유","복합소재",
  "게임","블록체인","유통","콘텐츠","엔터","웹툰","광고","한한령","이커머스",
  "푸드","프랜차이즈","화장품","여행","호텔","카지노","음식",
  "금융","은행","증권","보험","핀테크","가상자산","결제","스테이블","스테이블코인",
  "총선","저출산","고령화","남북경협","재건","원자재","곡물","농업","스마트팜","기후변화","스마트홈","헷지","헷지주",
  "신규주","IPO"
]


print("📢 KRX 상장목록 로드 중...")
try:
    df = fdr.StockListing("KRX")[["Code","Name"]]
    df["Code"] = df["Code"].astype(str).str.zfill(6)
    krx_names = df["Name"].tolist()
    name_to_code = dict(zip(df["Name"], df["Code"]))
except Exception:
    krx_names = []
    name_to_code = {}

reader = easyocr.Reader(['ko','en'], gpu=False)

def h2j(text):
    CHO = ['ㄱ','ㄲ','ㄴ','ㄷ','ㄸ','ㄹ','ㅁ','ㅂ','ㅃ','ㅅ','ㅆ','ㅇ','ㅈ','ㅉ','ㅊ','ㅋ','ㅌ','ㅍ','ㅎ']
    JUNG = ['ㅏ','ㅐ','ㅑ','ㅒ','ㅓ','ㅔ','ㅕ','ㅖ','ㅗ','ㅘ','ㅙ','ㅚ','ㅛ','ㅜ','ㅝ','ㅞ','ㅟ','ㅠ','ㅡ','ㅢ','ㅣ']
    JONG = ['','ㄱ','ㄲ','ㄳ','ㄴ','ㄵ','ㄶ','ㄷ','ㄹ','ㄺ','ㄻ','ㄼ','ㄽ','ㄾ','ㄿ','ㅀ','ㅁ','ㅂ','ㅄ','ㅅ','ㅆ','ㅇ','ㅈ','ㅊ','ㅋ','ㅌ','ㅍ','ㅎ']
    res = ""
    for c in text:
        if '가' <= c <= '힣':
            code = ord(c) - ord('가')
            res += CHO[code//588] + JUNG[(code//28)%21] + JONG[code%28]
        else:
            res += c
    return res

def microscopic_correct_stock(n):
    n_clean = re.sub(r'[0-9]', '', n).upper().replace(" ", "")
    if not n_clean:
        return ""
    n_comp = h2j(n_clean)
    candidates = []
    for s in krx_names:
        s_comp = h2j(s)
        if abs(len(s) - len(n_clean)) <= 2:
            dist = Levenshtein.distance(n_comp, s_comp)
            sim = 1 - (dist / max(len(n_comp), len(s_comp)) if max(len(n_comp), len(s_comp)) > 0 else 1)
            if s.startswith(n_clean[:1]):
                sim += 0.2
            candidates.append((s, sim))
    candidates.sort(key=lambda x: x[1], reverse=True)
    return candidates[0][0] if candidates and candidates[0][1] >= 0.52 else n_clean

def correct_theme_from_pool(raw):
    clean = re.sub(r'[^가-힣A-Z0-9]', '', raw.upper())
    if len(clean) < 2:
        return None
    for t in THEME_POOL:
        if t.upper() == clean:
            return t
    cj = h2j(clean)
    best, best_sim = None, 0
    for t in THEME_POOL:
        tj = h2j(t.upper())
        sim = 1 - (Levenshtein.distance(cj, tj) / max(len(cj), len(tj)))
        if t.startswith(clean[:1]):
            sim += 0.2
        if sim > best_sim:
            best_sim, best = sim, t
    return best if best_sim >= 0.5 else None

def pick_image():
    root = Tk(); root.withdraw()
    path = filedialog.askopenfilename(
        title="테마 분류표 이미지 선택",
        filetypes=[("Image files","*.png;*.jpg;*.jpeg;*.bmp;*.webp"),("All files","*.*")]
    )
    root.destroy()
    return path

def main():
    img_path = pick_image()
    if not img_path:
        print("❌ 이미지 선택 취소")
        return 2

    img = cv2.imread(img_path)
    if img is None:
        print("❌ 이미지 로드 실패")
        return 3

    print("🎨 [STEP1] 테마 분석...")
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    mask = cv2.inRange(hsv, (15,70,120), (45,255,255))
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    theme_locations = []
    for c in contours:
        x,y,w,h = cv2.boundingRect(c)
        if w < 30 or h < 8:
            continue
        roi = cv2.resize(img[y:y+h, x:x+w], None, fx=3, fy=3)
        raw_theme = "".join(reader.readtext(roi, detail=0))
        fixed_theme = correct_theme_from_pool(raw_theme)
        if fixed_theme:
            theme_locations.append({"name": fixed_theme, "x": x})

    print("🔍 [STEP2] 종목 분석...")
    img_res = cv2.resize(cv2.convertScaleAbs(img, alpha=1.5), None, fx=3.5, fy=3.5, interpolation=cv2.INTER_LANCZOS4)
    gray = cv2.cvtColor(img_res, cv2.COLOR_BGR2GRAY)
    thresh = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 21, 10)

    mapping_map = {}
    themes_seen = set()

    col_w = img_res.shape[1] // 4
    for i in range(4):
        c_start, c_end = i * col_w, (i + 1) * col_w

        current_theme = DEFAULT_THEME
        for tl in theme_locations:
            if (i * (img.shape[1]//4)) <= tl["x"] < ((i+1) * (img.shape[1]//4)):
                current_theme = tl["name"]
                break

        themes_seen.add(current_theme)

        h_sum = np.sum(thresh[:, c_start:c_end], axis=1)
        line_limit = np.mean(h_sum) * 0.4

        rows = []
        in_line, start = False, 0
        for idx, val in enumerate(h_sum):
            if (not in_line) and val > line_limit:
                in_line, start = True, idx
            elif in_line and val < line_limit:
                if idx - start > 18:
                    rows.append((start, idx))
                in_line = False

        for r_start, r_end in rows:
            chip = img_res[max(0, r_start-3):min(img_res.shape[0], r_end+3), c_start:c_end]
            if chip.size == 0:
                continue
            name_text = "".join(reader.readtext(chip[:, :int(chip.shape[1]*0.72)], detail=0))
            fixed_name = microscopic_correct_stock(name_text)
            code = name_to_code.get(fixed_name, "")
            if code and code.isdigit() and len(code) == 6:
                mapping_map[code] = current_theme

    out = {
        "themes": sorted([t for t in themes_seen if t]),
        "map": mapping_map
    }
    if DEFAULT_THEME not in out["themes"]:
        out["themes"].append(DEFAULT_THEME)

    with open(MAPPING_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    print("✅ 완료: mapping.json 생성 (codes=%d)" % (len(mapping_map),))
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
"""

    ocr_script = (
        ocr_template
        .replace("__MAPPING_PATH__", MAPPING_PATH.replace("\\", "\\\\"))
        .replace("__DEFAULT_THEME__", DEFAULT_THEME)
    )

    with tempfile.NamedTemporaryFile("w", suffix=".py", delete=False, encoding="utf-8") as tf:
        tf.write(ocr_script)
        tmp_path = tf.name

    try:
        print("[OCR] 호출:", " ".join(pycmd))
        r = subprocess.run(pycmd + [tmp_path], check=False)
        if r.returncode != 0:
            raise RuntimeError(f"OCR subprocess 실패 returncode={r.returncode}")
        if not os.path.exists(MAPPING_PATH):
            raise RuntimeError("OCR은 성공했는데 mapping.json이 없음")
        print("[OCR] mapping.json 생성 완료")
    finally:
        try:
            os.remove(tmp_path)
        except Exception:
            pass


# =========================
# XING 이벤트/클래스
# =========================
class XASessionEvents:
    def OnLogin(self, code, msg):
        self.parent._login_code = code
        self.parent._login_msg = msg

class XAQueryEvents:
    def OnReceiveData(self, tr_code):
        self.parent._received = True
        self.parent._last_tr = tr_code

class XingAPI:
    def __init__(self):
        self._received = False
        self._last_tr = ""
        self._login_code = None
        self._login_msg = ""

        self.session = win32com.client.DispatchWithEvents("XA_Session.XASession", XASessionEvents)
        self.session.parent = self

        self.query = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvents)
        self.query.parent = self

    def _wait(self, timeout, tag="TR timeout"):
        st = time.time()
        while not self._received:
            pythoncom.PumpWaitingMessages()
            if time.time() - st > timeout:
                raise TimeoutError(tag)
            time.sleep(0.01)

    def _set_res(self, res_filename: str) -> str:
        path = os.path.join(RES_DIR, res_filename)
        if not os.path.exists(path):
            raise FileNotFoundError(f"res 파일 없음: {path}")
        self.query.ResFileName = path
        return path

    def _get_field_try(self, outb: str, i: int, names):
        for nm in names:
            try:
                v = self.query.GetFieldData(outb, nm, i)
                if sstrip(v) != "":
                    return v
            except Exception:
                pass
        return ""

    def login(self):
        addr = SERVER_ADDR[CFG.server]
        if not self.session.ConnectServer(addr, 20001):
            raise RuntimeError("서버 연결 실패")

        server_type = 0 if CFG.server == "real" else 1
        self.session.Login(CFG.user_id, CFG.user_pw, CFG.cert_pw, server_type, 0)

        st = time.time()
        while self._login_code is None:
            pythoncom.PumpWaitingMessages()
            if time.time() - st > CFG.timeout_sec:
                raise TimeoutError("로그인 응답 타임아웃")
            time.sleep(0.01)

        if self._login_code != "0000":
            raise RuntimeError(f"로그인 실패: {self._login_code} {self._login_msg}")

        print("[LOGIN] 성공")

    def t1463_top(self):
        self._received = False
        self._set_res("t1463.res")

        inb = "t1463InBlock"
        self.query.SetFieldData(inb, "gubun", 0, "0")
        self.query.SetFieldData(inb, "jnilgubun", 0, "0")
        self.query.SetFieldData(inb, "idx", 0, "")

        ret = self.query.Request(0)
        if ret < 0:
            raise RuntimeError(f"t1463 Request 실패 ret={ret}")
        self._wait(CFG.timeout_sec, "t1463 timeout")

        outb = "t1463OutBlock1"
        cnt = self.query.GetBlockCount(outb)

        rows = []
        for i in range(cnt):
            code = sstrip(self.query.GetFieldData(outb, "shcode", i))
            name = sstrip(self.query.GetFieldData(outb, "hname", i))
            rate = to_float_or_none(self.query.GetFieldData(outb, "diff", i))

            open_px = to_int_or_none(self._get_field_try(outb, i, ["open", "openprc", "open_price", "openPrc", "opnprc"]))
            close_px = to_int_or_none(self._get_field_try(outb, i, ["price", "close", "closeprc", "close_price", "last", "nowprc", "curprc"]))

            if code and code.isdigit() and name:
                rows.append({"code": code, "name": name, "rate": rate, "open": open_px, "close": close_px})

        return sort_by_rate_desc(rows[:TOP_MONEY])

    def t1441_top(self):
        self._received = False
        self._set_res("t1441.res")

        inb = "t1441InBlock"
        self.query.SetFieldData(inb, "gubun1", 0, "0")
        self.query.SetFieldData(inb, "gubun2", 0, "0")
        self.query.SetFieldData(inb, "gubun3", 0, "0")
        self.query.SetFieldData(inb, "idx", 0, "")

        ret = self.query.Request(0)
        if ret < 0:
            raise RuntimeError(f"t1441 Request 실패 ret={ret}")
        self._wait(CFG.timeout_sec, "t1441 timeout")

        outb = "t1441OutBlock1"
        cnt = self.query.GetBlockCount(outb)

        rows = []
        for i in range(cnt):
            code = sstrip(self.query.GetFieldData(outb, "shcode", i))
            name = sstrip(self.query.GetFieldData(outb, "hname", i))
            rate = to_float_or_none(self.query.GetFieldData(outb, "diff", i))

            open_px = to_int_or_none(self._get_field_try(outb, i, ["open", "openprc", "open_price", "openPrc", "opnprc"]))
            close_px = to_int_or_none(self._get_field_try(outb, i, ["price", "close", "closeprc", "close_price", "last", "nowprc", "curprc"]))

            if code and code.isdigit() and name:
                rows.append({"code": code, "name": name, "rate": rate, "open": open_px, "close": close_px})

        return sort_by_rate_desc(rows)[:TOP_RATE]

    def t8407_quotes(self, codes):
        """
        mapping.json에 있는 '전체 종목코드'를 현재가/등락률로 조회하기 위한 멀티 TR.
        ✅ res 파일 필드명이 환경마다 조금 다를 수 있어서 후보 필드 여러개로 시도함.
        """
        codes = [c for c in codes if isinstance(c, str) and c.isdigit() and len(c) == 6]
        if not codes:
            return []

        self._received = False
        self._set_res("t8407.res")

        inb = "t8407InBlock"
        # shcode: 종목코드 리스트를 ';'로 연결하는 경우가 많음
        self.query.SetFieldData(inb, "shcode", 0, ";".join(codes))

        ret = self.query.Request(0)
        if ret < 0:
            raise RuntimeError(f"t8407 Request 실패 ret={ret}")
        self._wait(CFG.timeout_sec, "t8407 timeout")

        outb = "t8407OutBlock1"
        cnt = self.query.GetBlockCount(outb)

        rows = []
        for i in range(cnt):
            code = sstrip(self._get_field_try(outb, i, ["shcode", "code"]))
            name = sstrip(self._get_field_try(outb, i, ["hname", "name"]))
            # 등락률 후보
            rate = to_float_or_none(self._get_field_try(outb, i, ["diff", "drate", "chgrate", "changeRate", "updnrate"]))
            if code and code.isdigit() and len(code) == 6:
                rows.append({"code": code, "name": name or code, "rate": rate})
        return rows


def is_bullish(row):
    o = row.get("open"); c = row.get("close")
    return (o is not None and c is not None and c > o)


# =========================
# main
# =========================
def main():
    pythoncom.CoInitialize()

    # 1) OCR(3.12) -> mapping.json 생성(코드->테마)
    run_ocr_with_py312_make_mapping()

    # 2) mapping 로드
    code_to_theme = load_mapping_code_to_theme()
    if not code_to_theme:
        print("[WARN] mapping.json 비어있음 -> 테마 패널 출력 불가(또는 전부 미분류)")
        # 그래도 XING 패널은 돌아가게는 둠

    # 3) XING 로그인
    x = XingAPI()
    x.login()

    # mapping 전체 코드 고정(루프마다 파일 다시 읽고 싶으면 여기 말고 루프 안에서 reload하면 됨)
    mapping_codes = sorted(code_to_theme.keys())

    while True:
        try:
            clear_screen()
            now = time.strftime("%Y-%m-%d %H:%M:%S")

            # 상단 3패널
            money_rows_raw = x.t1463_top()
            rate_rows_raw = x.t1441_top()

            money_rows = sort_by_rate_desc(apply_min_rate_filter(money_rows_raw, PRINT_MIN_RATE))
            rate_rows = sort_by_rate_desc(apply_min_rate_filter(rate_rows_raw, PRINT_MIN_RATE))

            # 교집합 + 양봉(상단 패널용)
            money_map = {r["code"]: r for r in money_rows_raw}
            rate_map = {r["code"]: r for r in rate_rows_raw}
            common = [c for c in money_map.keys() if c in rate_map]

            bull_rows = []
            for c in common:
                a = money_map[c]; b = rate_map[c]
                merged = {
                    "code": c,
                    "name": a.get("name") or b.get("name") or c,
                    "rate": b.get("rate") if b.get("rate") is not None else a.get("rate"),
                    "open": a.get("open") if a.get("open") is not None else b.get("open"),
                    "close": a.get("close") if a.get("close") is not None else b.get("close"),
                }
                if is_bullish(merged):
                    bull_rows.append(merged)

            bull_rows = sort_by_rate_desc(apply_min_rate_filter(bull_rows, PRINT_MIN_RATE))

            p_money = build_panel_lines("[거래대금상위]", money_rows, None)
            p_rate  = build_panel_lines("[등락률상위]",   rate_rows,  None)
            p_bull  = build_panel_lines("[교집합+양봉]",   bull_rows,  None)

            print_panels_side_by_side([p_money, p_rate, p_bull], gap=" | ")

            # =========================
            # ✅ 여기부터가 너가 원하는 "mapping 전체 종목 -> 테마 패널"
            # - mapping.json에 있는 전체 종목을 t8407로 조회
            # - rate >= 5%만 남기고
            # - 테마별 패널을 "거래대금/등락률/교집합"과 같은 방식으로 옆으로 출력
            # =========================
            if mapping_codes:
                all_rows = []
                for i in range(0, len(mapping_codes), T8407_BATCH):
                    batch = mapping_codes[i:i+T8407_BATCH]
                    all_rows.extend(x.t8407_quotes(batch))

                # rate>=5%만
                all_rows = sort_by_rate_desc(apply_min_rate_filter(all_rows, PRINT_MIN_RATE))

                if all_rows:
                    buckets = group_rows_by_theme(all_rows, code_to_theme)
                    theme_order = sorted(buckets.keys(), key=lambda t: (t == DEFAULT_THEME, -len(buckets[t]), t))

                    print("\n" + "=" * 90)
                    print(f"[테마별 분해] (mapping.json 전체 종목 기준 / {PRINT_MIN_RATE}% 이상만)")
                    print("=" * 90)

                    per_row = 3
                    for k in range(0, len(theme_order), per_row):
                        chunk = theme_order[k:k+per_row]
                        infos = []
                        for t in chunk:
                            title = f"[{t}] ({len(buckets[t])})"
                            infos.append(build_panel_lines(title, buckets[t], None))
                        print_panels_side_by_side(infos, gap=" | ")
                        print("")
                else:
                    print("\n" + "=" * 90)
                    print(f"[테마별 분해] (mapping.json 전체 종목 기준 / {PRINT_MIN_RATE}% 이상) -> (없음)")
                    print("=" * 90)

            print("\n[TIME]", now)

        except KeyboardInterrupt:
            break
        except Exception as e:
            print("\n[ERROR]", e)

        time.sleep(REFRESH_SEC)


if __name__ == "__main__":
    main()
