import cv2
import numpy as np
import easyocr
import re
import pandas as pd
import FinanceDataReader as fdr
import Levenshtein
import os
import json
from tkinter import Tk, filedialog

# ==========================================
# [공통] 상장사 및 OCR 초기화
# ==========================================
print("📢 상장사 목록 로드 중...")
try: krx_names = fdr.StockListing("KRX")["Name"].tolist()
except: krx_names = []

reader = easyocr.Reader(['ko', 'en'], gpu=False)

# [자모 분해 함수]
def h2j(text):
    CHO = ['ㄱ','ㄲ','ㄴ','ㄷ','ㄸ','ㄹ','ㅁ','ㅂ','ㅃ','ㅅ','ㅆ','ㅇ','ㅈ','ㅉ','ㅊ','ㅋ','ㅌ','ㅍ','ㅎ']
    JUNG = ['ㅏ','ㅐ','ㅑ','ㅒ','ㅓ','ㅔ','ㅕ','ㅖ','ㅗ','ㅘ','ㅙ','ㅚ','ㅛ','ㅜ','ㅝ','ㅞ','ㅟ','ㅠ','ㅡ','ㅢ','ㅣ']
    JONG = ['','ㄱ','ㄲ','ㄳ','ㄴ','ㄵ','ㄶ','ㄷ','ㄹ','ㄺ','ㄻ','ㄼ','ㄽ','ㄾ','ㄿ','ㅀ','ㅁ','ㅂ','ㅄ','ㅅ','ㅆ','ㅇ','ㅈ','ㅊ','ㅋ','ㅌ','ㅍ','ㅎ']
    res = ""
    for c in text:
        if '가' <= c <= '힣':
            code = ord(c) - ord('가')
            res += CHO[code//588] + JUNG[(code//28)%21] + JONG[code%28]
        else: res += c
    return res

# ==========================================
# [로직 1] 종목 보정 엔진 (Microscopic)
# ==========================================
def microscopic_correct_stock(n):
    n_clean = re.sub(r'[0-9]', '', n).upper().replace(" ", "")
    if not n_clean or n_clean in krx_names: return n_clean
    n_comp = h2j(n_clean)
    candidates = []
    for s in krx_names:
        s_comp = h2j(s)
        if abs(len(s) - len(n_clean)) <= 2:
            dist = Levenshtein.distance(n_comp, s_comp)
            sim = 1 - (dist / max(len(n_comp), len(s_comp)) if max(len(n_comp), len(s_comp)) > 0 else 1)
            if s.startswith(n_clean[0]): sim += 0.2
            candidates.append((s, sim))
    candidates.sort(key=lambda x: x[1], reverse=True)
    return candidates[0][0] if candidates and candidates[0][1] >= 0.52 else n_clean

# ==========================================
# [로직 2] 테마 보정 엔진 (Pool)
# ==========================================
THEME_POOL = ["로봇","반도체","바이오","자동차","2차전지","AI","우주항공","방산","신약개발","자율주행"] # 위에 주신 리스트 사용

def correct_theme_from_pool(raw):
    clean = re.sub(r'[^가-힣A-Z0-9]', '', raw.upper())
    if len(clean) < 2: return None
    if clean in THEME_POOL: return clean
    cj = h2j(clean)
    best, best_sim = None, 0
    for t in THEME_POOL:
        tj = h2j(t.upper())
        sim = 1 - (Levenshtein.distance(cj, tj) / max(len(cj), len(tj)))
        if t.startswith(clean[:1]): sim += 0.2
        if sim > best_sim: best_sim, best = sim, t
    return best if best_sim >= 0.5 else None

# ==========================================
# [통합 분석] 사진 한 장으로 두 로직 따로 돌리기
# ==========================================
def run_integrated_analysis():
    root = Tk(); root.withdraw(); img_path = filedialog.askopenfilename(); root.destroy()
    if not img_path: return
    img = cv2.imread(img_path)

    # --- 1. 테마 분석 (노란색 마스크 로직) ---
    print("🎨 [STEP 1] 테마 분석 중...")
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    mask = cv2.inRange(hsv, (15,70,120), (45,255,255))
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    theme_locations = []
    for c in contours:
        x,y,w,h = cv2.boundingRect(c)
        if w < 30 or h < 8: continue
        roi = cv2.resize(img[y:y+h, x:x+w], None, fx=3, fy=3)
        raw_theme = "".join(reader.readtext(roi, detail=0))
        fixed_theme = correct_theme_from_pool(raw_theme)
        if fixed_theme:
            theme_locations.append({'name': fixed_theme, 'x': x, 'y': y})

    # --- 2. 종목 분석 (기둥 및 행 분석 로직) ---
    print("🔍 [STEP 2] 종목 분석 중...")
    img_res = cv2.resize(cv2.convertScaleAbs(img, alpha=1.5), None, fx=3.5, fy=3.5, interpolation=cv2.INTER_LANCZOS4)
    thresh = cv2.adaptiveThreshold(cv2.cvtColor(img_res, cv2.COLOR_BGR2GRAY), 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 21, 10)
    
    mapping_result = {}
    col_w = img_res.shape[1] // 4
    
    for i in range(4): # 4개 기둥 분석
        c_start, c_end = i * col_w, (i + 1) * col_w
        # 해당 기둥에 속한 테마 찾기
        current_column_theme = "미분류"
        for tl in theme_locations:
            if (i * (img.shape[1]//4)) <= tl['x'] < ((i+1) * (img.shape[1]//4)):
                current_column_theme = tl['name']
                break

        h_sum = np.sum(thresh[:, c_start:c_end], axis=1)
        line_limit = np.mean(h_sum) * 0.4
        rows = []; in_line, start = False, 0
        for idx, val in enumerate(h_sum):
            if not in_line and val > line_limit: in_line, start = True, idx
            elif in_line and val < line_limit:
                if idx - start > 18: rows.append((start, idx))
                in_line = False

        for r_start, r_end in rows:
            chip = img_res[r_start-3:r_end+3, c_start:c_end]
            name_text = "".join(reader.readtext(chip[:, :int(chip.shape[1]*0.72)], detail=0))
            refined_stock = microscopic_correct_stock(name_text)
            if len(refined_stock) >= 2:
                mapping_result[refined_stock] = current_column_theme

    # --- 3. 최종 저장 ---
    with open("mapping.json", "w", encoding="utf-8") as f:
        json.dump(mapping_result, f, ensure_ascii=False, indent=2)
    print(f"🎯 완료! {len(mapping_result)}개 종목이 매핑되었습니다.")

if __name__ == "__main__":
    run_integrated_analysis()
