#!/usr/bin/env python3
"""
네이버 지면 캡처 자동화
- 4개 키워드 PC + 모바일 전체 페이지 고화질 캡처
- 구글 드라이브 폴더에 원본 저장 (동기화 방식)
- HTML 리포트 생성 (PC/모바일 가로 4열 레이아웃)
- 구글 시트에 날짜별 누적 삽입
"""

import warnings
warnings.filterwarnings('ignore')
import os
import re
import time
import requests
from datetime import datetime
from pathlib import Path

import subprocess
from playwright.sync_api import sync_playwright
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ── 설정 ─────────────────────────────────────────────
KEY_FILE      = os.environ.get('GOOGLE_KEY_FILE', '/Users/teamo2/Desktop/elly/구글시트계정.json')
SHEET_ID      = '1uDLJAB3p4aUXTzECF7gmxrC0LPWl_58vO3bsRRnn4zI'
SHEET_TAB             = '광고'
SHEET_TAB_CONTENT     = '콘텐츠 지면'
SHEET_TAB_COMPETITOR  = '경쟁사 분석'
SHEET_TAB_CONTENT_BIZ      = '콘텐츠 업체 현황'
SHEET_TAB_ANALYSIS         = '광고 분석 리포트'
SHEET_TAB_CONTENT_ANALYSIS = '콘텐츠 분석 리포트'

# 자사 / 주요 경쟁사
MY_COMPANY  = '카모아'
KEY_RIVAL   = '돌하루팡'
MY_URL_KEY  = 'camoa'      # URL에서 자사 판별용
RIVAL_URL_KEY = 'dolharu'  # URL에서 돌하루팡 판별용
DRIVE_FOLDER  = '13-2hvJcyXpLSILsgCRhduPr79jh0LS60'
LOCAL_DRIVE   = os.environ.get('LOCAL_DRIVE', '/tmp/naver_capture')

KEYWORDS = ['제주렌트카', '제주도렌트카', '제주렌터카', '제주도렌터카']

SLACK_TOKEN   = 'SLACK_BOT_TOKEN_HERE'   # 아래 안내 참고
SLACK_CHANNEL = 'C05C7D9F5K8'            # #마케팅팀_대화

SCOPES   = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

TODAY = datetime.now().strftime('%Y-%m-%d')
NOW   = datetime.now().strftime('%Y-%m-%d %H:%M')

# PC 캡처 설정
PC_CROP_HEIGHT = 2400
PC_VIEWPORT    = {'width': 1280, 'height': 1080}

# 모바일 캡처 설정
MOBILE_CROP_HEIGHT = 4000
MOBILE_VIEWPORT    = {'width': 390, 'height': 844}
MOBILE_UA          = 'Mozilla/5.0 (iPhone; CPU iPhone OS 16_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.0 Mobile/15E148 Safari/604.1'

COL_WIDTH = 1280
# ─────────────────────────────────────────────────────


def get_services():
    creds  = service_account.Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
    sheets = build('sheets', 'v4', credentials=creds)
    drive  = build('drive', 'v3', credentials=creds)
    return sheets, drive


def capture_pages(mode='pc', content=False):
    """
    4개 키워드 스크린샷 → 로컬 드라이브 저장
    content=False: 0 ~ crop_h (지면확보)
    content=True : crop_h ~ 끝 (콘텐츠 지면)
    """
    Path(LOCAL_DRIVE).mkdir(parents=True, exist_ok=True)
    paths = {}

    if mode == 'mobile':
        vp     = MOBILE_VIEWPORT
        ua     = MOBILE_UA
        scale  = 2
        crop_h = MOBILE_CROP_HEIGHT
        suffix = 'mobile'
    else:
        vp     = PC_VIEWPORT
        ua     = None
        scale  = 1
        crop_h = PC_CROP_HEIGHT
        suffix = 'pc'

    tag = ('content_' if content else '') + suffix

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        # 시크릿창 동일 조건: 쿠키/캐시/로그인 정보 없는 새 컨텍스트
        ctx_kwargs = dict(
            viewport=vp,
            device_scale_factor=scale,
            no_viewport=False,
            storage_state=None,   # 저장된 세션 없음
        )
        if ua:
            ctx_kwargs['user_agent'] = ua
        context = browser.new_context(**ctx_kwargs)
        page = context.new_page()
        page.set_extra_http_headers({'Accept-Language': 'ko-KR,ko;q=0.9'})

        for kw in KEYWORDS:
            print(f'  [{tag.upper()}] 캡처: {kw}')
            page.goto(f'https://search.naver.com/search.naver?query={kw}',
                      wait_until='networkidle', timeout=30000)
            page.wait_for_timeout(2000)

            # 페이지 전체 스크롤 → 이미지 lazy load 트리거
            page.evaluate('''() => {
                return new Promise(resolve => {
                    let y = 0;
                    const step = 600;
                    const timer = setInterval(() => {
                        window.scrollBy(0, step);
                        y += step;
                        if (y >= document.body.scrollHeight) {
                            window.scrollTo(0, 0);
                            clearInterval(timer);
                            resolve();
                        }
                    }, 100);
                });
            }''')
            page.wait_for_timeout(2000)

            fname = f'{TODAY}_{kw}_{tag}.png'
            fpath = f'{LOCAL_DRIVE}/{fname}'
            page.screenshot(path=fpath, full_page=True)

            from PIL import Image
            img = Image.open(fpath)
            if content:
                # 브랜드콘텐츠 sc_new 블럭 상단 Y 위치 탐색
                brand_y = page.evaluate('''() => {
                    const sels = [".fds-comps-footer-more-subject", ".sds-comps-footer-label", "[class*='footer-more-subject']", "[class*='footer-label']"];
                    for (const sel of sels) {
                        for (const el of document.querySelectorAll(sel)) {
                            if ((el.innerText || "").includes("브랜드")) {
                                const block = el.closest(".sc_new") || el.closest("._fe_view_root");
                                if (block) return block.getBoundingClientRect().top + window.scrollY;
                            }
                        }
                    }
                    return null;
                }''')
                start = int(brand_y * scale) if brand_y else min(crop_h, img.height)
                start = min(start, img.height)
                img_cropped = img.crop((0, start, img.width, img.height))
                print(f'    브랜드콘텐츠 시작: {start}px (원본 brand_y={brand_y})')
            else:
                img_cropped = img.crop((0, 0, img.width, min(crop_h, img.height)))
            img_cropped.save(fpath, optimize=False, compress_level=1)
            print(f'    해상도: {img_cropped.size}')
            paths[kw] = (fpath, fname)

        context.close()
        browser.close()
    return paths


def upload_to_drive(drive, fpath, fname):
    """파일을 Drive API로 직접 업로드 후 파일 ID 반환"""
    from googleapiclient.http import MediaFileUpload
    file_metadata = {'name': fname, 'parents': [DRIVE_FOLDER]}
    media = MediaFileUpload(fpath, mimetype='image/png')
    file = drive.files().create(
        body=file_metadata,
        media_body=media,
        fields='id'
    ).execute()
    print(f'  업로드 완료: {fname} (id={file["id"]})')
    return file['id']


def make_public_url(drive, file_id):
    """파일 공개 설정 후 URL 반환"""
    drive.permissions().create(
        fileId=file_id,
        body={'role': 'reader', 'type': 'anyone'}
    ).execute()
    return f'https://drive.google.com/uc?export=view&id={file_id}'


def get_sheet_id(sheets):
    meta = sheets.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
    for s in meta['sheets']:
        if s['properties']['title'] == SHEET_TAB:
            return s['properties']['sheetId']
    raise ValueError(f'탭 없음: {SHEET_TAB}')


def get_or_create_tab(sheets, tab_name):
    """탭이 없으면 새로 만들고 sheetId 반환"""
    meta = sheets.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
    for s in meta['sheets']:
        if s['properties']['title'] == tab_name:
            return s['properties']['sheetId']
    # 탭 신규 생성
    resp = sheets.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID,
        body={'requests': [{'addSheet': {'properties': {'title': tab_name}}}]}
    ).execute()
    return resp['replies'][0]['addSheet']['properties']['sheetId']


def get_next_row(sheets, tab=None):
    t = tab or SHEET_TAB
    result = sheets.spreadsheets().values().get(
        spreadsheetId=SHEET_ID,
        range=f"'{t}'!A:A"
    ).execute()
    return len(result.get('values', [])) + 2


def render_html_screenshot(kw_paths, mode, brand_data=None):
    """
    4개 키워드를 가로 4열 HTML로 렌더링 후 Playwright 스크린샷 → PNG
    brand_data: {kw: [{'section','author','brand','title'}, ...]} 전달 시
                이미지 아래에 콘텐츠 업체 분석 블록을 추가
    """
    import base64

    # 섹션 유형별 아이콘
    SECTION_ICON = {
        '브랜드콘텐츠': '🏷',
        '인플루언서':   '⭐',
        '인기글':       '🔥',
        '카페글':       '☕',
        '연관업체':     '🔗',
    }

    cols_html = ''
    for kw, fpath in kw_paths.items():
        with open(fpath, 'rb') as f:
            b64 = base64.b64encode(f.read()).decode()

        # 업체 분석 블록 (brand_data 있을 때만)
        brand_block = ''
        if brand_data and kw in brand_data:
            device_key = 'PC' if 'PC' in mode else '모바일'
            items = brand_data[kw].get(device_key, [])
            if items:
                rows_html = ''
                for it in items:
                    icon = SECTION_ICON.get(it.get('section', ''), '📄')
                    sec  = it.get('section', '')
                    biz  = it.get('brand') or it.get('author', '')
                    title = it.get('title', '')[:50]
                    rows_html += f'''
                    <tr>
                      <td class="tag">{icon} {sec}</td>
                      <td class="biz">{biz}</td>
                      <td class="ttl">{title}</td>
                    </tr>'''
                brand_block = f'''
                <div class="brand-box">
                  <div class="brand-hd">📊 콘텐츠 지면 업체 현황</div>
                  <table>
                    <thead><tr>
                      <th>섹션</th><th>업체명</th><th>콘텐츠 제목</th>
                    </tr></thead>
                    <tbody>{rows_html}</tbody>
                  </table>
                </div>'''

        cols_html += f'''
        <div class="col">
            <div class="kw-title">{kw}</div>
            <img src="data:image/png;base64,{b64}" />
            {brand_block}
        </div>'''

    mode_label = 'PC' if 'PC' in mode else '모바일'
    html = f'''<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: "Apple SD Gothic Neo", sans-serif; background: #f0f2f5; }}
  header {{ background: #222; color: #fff; padding: 20px 28px; }}
  header h1 {{ font-size: 22px; font-weight: 700; }}
  header span {{ font-size: 16px; color: #bbb; margin-left: 16px; }}
  .grid {{ display: flex; gap: 0; align-items: flex-start; }}
  .col {{ flex: 1; display: flex; flex-direction: column; border-right: 3px solid #ddd; }}
  .col:last-child {{ border-right: none; }}
  .kw-title {{ background: #03C75A; color: #fff; font-weight: 700; font-size: 16px;
               padding: 12px 16px; }}
  img {{ width: 100%; display: block; }}
  .brand-box {{ background: #fff; border-top: 3px solid #03C75A; padding: 14px 12px; }}
  .brand-hd {{ font-size: 13px; font-weight: 700; color: #333; margin-bottom: 8px; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
  thead tr {{ background: #f5f5f5; }}
  th, td {{ padding: 6px 8px; text-align: left; border-bottom: 1px solid #eee; }}
  th {{ font-weight: 700; color: #555; font-size: 11px; }}
  .tag {{ white-space: nowrap; color: #666; width: 90px; }}
  .biz {{ font-weight: 700; color: #03C75A; width: 110px; }}
  .ttl {{ color: #333; }}
</style>
</head>
<body>
<header>
  <h1>네이버 지면 캡처 · {mode_label}</h1><span>{NOW}</span>
</header>
<div class="grid">
{cols_html}
</div>
</body>
</html>'''

    html_path = f'{LOCAL_DRIVE}/{TODAY}_네이버지면_{mode}.html'
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html)

    out_path = f'{LOCAL_DRIVE}/{TODAY}_네이버지면_{mode}.png'
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(
            viewport={'width': 1600, 'height': 900},
            device_scale_factor=3          # 3배 고해상도 → 원본에 가까운 화질
        )
        page.goto(f'file://{html_path}', wait_until='networkidle')
        page.wait_for_timeout(1000)
        page.screenshot(path=out_path, full_page=True)
        browser.close()

    print(f'  → HTML 스크린샷: {Path(out_path).name}')
    return out_path, Path(out_path).name


def append_to_sheet(sheets, sheet_id, pc_urls, mobile_urls, pc_stitched_url, mob_stitched_url, tab=None):
    """
    레이아웃:
      행1 = 날짜
      행2 = 🖥PC 합성보기 | 키워드별 원본 링크
      행3 = 📱모바일 합성보기 | 키워드별 원본 링크
    """
    t = tab or SHEET_TAB
    start = get_next_row(sheets, tab=t)
    kw_list = list(KEYWORDS)
    values_data     = []
    resize_requests = []

    # ── 날짜 헤더 행 ─────────────────────────────────────
    date_row = start
    values_data.append({
        'range': f"'{t}'!A{date_row}",
        'values': [[f'📅 {NOW}']]
    })
    resize_requests.append({'updateDimensionProperties': {
        'range': {'sheetId': sheet_id, 'dimension': 'ROWS',
                  'startIndex': date_row - 1, 'endIndex': date_row},
        'properties': {'pixelSize': 32}, 'fields': 'pixelSize'
    }})

    # ── PC 행: A열=합성 클릭, B~E=키워드별 원본 ──────────
    pc_row  = date_row + 1
    pc_vals = [f'=HYPERLINK("{pc_stitched_url}","🖥 PC 전체보기")']
    for kw in kw_list:
        file_id  = pc_urls[kw].split('id=')[-1]
        view_url = f'https://drive.google.com/uc?export=view&id={file_id}'
        pc_vals.append(f'=HYPERLINK("{view_url}","📂 {kw}")')
    values_data.append({
        'range': f"'{t}'!A{pc_row}",
        'values': [pc_vals]
    })
    resize_requests.append({'updateDimensionProperties': {
        'range': {'sheetId': sheet_id, 'dimension': 'ROWS',
                  'startIndex': pc_row - 1, 'endIndex': pc_row},
        'properties': {'pixelSize': 28}, 'fields': 'pixelSize'
    }})

    # ── 모바일 행: A열=합성 클릭, B~E=키워드별 원본 ───────
    mob_row  = date_row + 2
    mob_vals = [f'=HYPERLINK("{mob_stitched_url}","📱 모바일 전체보기")']
    for kw in kw_list:
        file_id  = mobile_urls[kw].split('id=')[-1]
        view_url = f'https://drive.google.com/uc?export=view&id={file_id}'
        mob_vals.append(f'=HYPERLINK("{view_url}","📂 {kw}")')
    values_data.append({
        'range': f"'{t}'!A{mob_row}",
        'values': [mob_vals]
    })
    resize_requests.append({'updateDimensionProperties': {
        'range': {'sheetId': sheet_id, 'dimension': 'ROWS',
                  'startIndex': mob_row - 1, 'endIndex': mob_row},
        'properties': {'pixelSize': 28}, 'fields': 'pixelSize'
    }})

    # ── 열 너비: A=120, B~E=220 ───────────────────────────
    resize_requests.append({'updateDimensionProperties': {
        'range': {'sheetId': sheet_id, 'dimension': 'COLUMNS',
                  'startIndex': 0, 'endIndex': 1},
        'properties': {'pixelSize': 120}, 'fields': 'pixelSize'
    }})
    resize_requests.append({'updateDimensionProperties': {
        'range': {'sheetId': sheet_id, 'dimension': 'COLUMNS',
                  'startIndex': 1, 'endIndex': 5},
        'properties': {'pixelSize': 220}, 'fields': 'pixelSize'
    }})

    sheets.spreadsheets().values().batchUpdate(
        spreadsheetId=SHEET_ID,
        body={'valueInputOption': 'USER_ENTERED', 'data': values_data}
    ).execute()

    sheets.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID,
        body={'requests': resize_requests}
    ).execute()

    print(f'  → {date_row}~{mob_row}행 삽입 완료')


def scrape_competitor_ads():
    """4개 키워드 × PC + 모바일 파워링크 광고 텍스트 추출"""
    all_ads = {}

    # PC: 광고주 단위(li) 기준 — 사이트링크 중복 제거
    AD_JS_PC = r"""() => {
        const section = document.querySelector('.ad_section');
        if (!section) return [];
        const mainList = section.querySelector('ul, ol');
        if (!mainList) return [];
        const lis = Array.from(mainList.children).filter(el => el.tagName === 'LI');
        const results = [];
        lis.forEach((li, idx) => {
            const item    = li.querySelector('.title_url_area');
            if (!item) return;
            const titleEl = item.querySelector('.site, a.site');
            const urlEl   = item.querySelector('.url_area');
            const descEl  = li.querySelector('.desc_area, .desc, p');
            const urlRaw  = urlEl ? urlEl.innerText.trim() : '';
            const urlLines = urlRaw.split('\n').map(s => s.trim()).filter(Boolean);
            const domain  = urlLines[urlLines.length - 1] || '';
            results.push({
                rank: idx + 1,
                title: titleEl ? titleEl.innerText.trim() : '',
                desc: descEl ? descEl.innerText.trim() : '',
                url: domain
            });
        });
        return results;
    }"""

    # 모바일: 광고주 단위(li) 기준 — 사이트링크 중복 제거
    AD_JS_MOB = r"""() => {
        const section = document.querySelector('._pwl_video_container, .sc.ad_light_mode');
        if (!section) return [];
        const mainList = section.querySelector('ul, ol');
        if (!mainList) return [];
        const lis = Array.from(mainList.children).filter(el => el.tagName === 'LI');
        const results = [];
        lis.forEach((li, idx) => {
            const item   = li.querySelector('.tit_wrap');
            if (!item) return;
            const siteEl = item.querySelector('.site');
            const urlEl  = item.querySelector('.url');
            const linkEl = item.querySelector('a.txt_link');
            const descEl = li.querySelector('.desc_area, .desc, .dsc_area, .dsc');
            let adHeadline = '';
            if (linkEl) {
                const lines = linkEl.innerText.trim().split('\n').map(s => s.trim()).filter(Boolean);
                adHeadline = lines[lines.length - 1] || '';
            }
            const company = siteEl ? siteEl.innerText.trim() : '';
            if (!company) return;
            results.push({
                rank: idx + 1,
                title: company,
                desc: adHeadline + (descEl ? ' | ' + descEl.innerText.trim() : ''),
                url: urlEl ? urlEl.innerText.trim() : ''
            });
        });
        return results;
    }"""

    def extract_ads(page, kw, device):
        js = AD_JS_MOB if device == '모바일' else AD_JS_PC
        ads = page.evaluate(js)
        all_ads.setdefault(kw, {})[device] = ads
        print(f'    [{device}] {kw}: 광고 {len(ads)}개')

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)

        # ── PC ───────────────────────────────────────────
        ctx_pc = browser.new_context(viewport=PC_VIEWPORT, device_scale_factor=1, storage_state=None)
        page_pc = ctx_pc.new_page()
        page_pc.set_extra_http_headers({'Accept-Language': 'ko-KR,ko;q=0.9'})
        for kw in KEYWORDS:
            page_pc.goto(f'https://search.naver.com/search.naver?query={kw}',
                         wait_until='networkidle', timeout=30000)
            page_pc.wait_for_timeout(2000)
            extract_ads(page_pc, kw, 'PC')
        ctx_pc.close()

        # ── 모바일 ───────────────────────────────────────
        ctx_mob = browser.new_context(
            viewport=MOBILE_VIEWPORT, device_scale_factor=2,
            user_agent=MOBILE_UA, storage_state=None
        )
        page_mob = ctx_mob.new_page()
        page_mob.set_extra_http_headers({'Accept-Language': 'ko-KR,ko;q=0.9'})
        for kw in KEYWORDS:
            page_mob.goto(f'https://search.naver.com/search.naver?query={kw}',
                          wait_until='networkidle', timeout=30000)
            page_mob.wait_for_timeout(2000)
            extract_ads(page_mob, kw, '모바일')
        ctx_mob.close()

        browser.close()

    return all_ads


def write_competitor_tab(sheets, all_ads):
    """경쟁사 분석 탭에 날짜별 광고 데이터 누적 추가"""
    tab = SHEET_TAB_COMPETITOR
    sheet_id = get_or_create_tab(sheets, tab)

    # 헤더 없으면 1행에 추가
    existing = sheets.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range=f"'{tab}'!A1:G1"
    ).execute().get('values', [])

    if not existing or (existing[0] and existing[0][0] != '날짜'):
        sheets.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=f"'{tab}'!A1",
            valueInputOption='USER_ENTERED',
            body={'values': [['날짜', '기기', '키워드', '순위', '광고 제목', '광고 설명', '노출 URL']]}
        ).execute()
        # 헤더 배경색 (네이버 초록)
        sheets.spreadsheets().batchUpdate(
            spreadsheetId=SHEET_ID,
            body={'requests': [{'repeatCell': {
                'range': {'sheetId': sheet_id, 'startRowIndex': 0, 'endRowIndex': 1},
                'cell': {'userEnteredFormat': {
                    'backgroundColor': {'red': 0.01, 'green': 0.78, 'blue': 0.35},
                    'textFormat': {'bold': True,
                                   'foregroundColor': {'red': 1, 'green': 1, 'blue': 1}}
                }},
                'fields': 'userEnteredFormat(backgroundColor,textFormat)'
            }}]}
        ).execute()

    # 다음 빈 행
    col_a = sheets.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range=f"'{tab}'!A:A"
    ).execute().get('values', [])
    next_row = len(col_a) + 1

    # 데이터 행 생성: 키워드별로 묶고 사이에 구분선 삽입
    rows = []          # 실제 셀 값
    sep_rows = []      # 구분선으로 처리할 행 번호 (1-based)
    data_row_count = 0

    for i, kw in enumerate(KEYWORDS):
        # ── 키워드 구분선 ─────────────────────────────────
        sep_row_idx = next_row + len(rows)
        sep_rows.append(sep_row_idx)
        rows.append([f'▶ {kw}', '', '', '', '', '', ''])

        for device in ['PC', '모바일']:
            ads = all_ads.get(kw, {}).get(device, [])
            if not ads:
                rows.append([TODAY, device, kw, '-', '광고 없음', '', ''])
                data_row_count += 1
            else:
                for ad in ads:
                    rows.append([
                        TODAY,
                        device,
                        kw,
                        ad.get('rank', ''),
                        ad.get('title', ''),
                        ad.get('desc', ''),
                        ad.get('url', '')
                    ])
                    data_row_count += 1

    if rows:
        sheets.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=f"'{tab}'!A{next_row}",
            valueInputOption='USER_ENTERED',
            body={'values': rows}
        ).execute()

    # 구분선 행 서식: 진한 회색 배경 + 흰 굵은 글씨
    fmt_requests = []
    for r in sep_rows:
        fmt_requests.append({'repeatCell': {
            'range': {'sheetId': sheet_id,
                      'startRowIndex': r - 1, 'endRowIndex': r,
                      'startColumnIndex': 0, 'endColumnIndex': 7},
            'cell': {'userEnteredFormat': {
                'backgroundColor': {'red': 0.25, 'green': 0.25, 'blue': 0.25},
                'textFormat': {'bold': True,
                               'fontSize': 11,
                               'foregroundColor': {'red': 1, 'green': 1, 'blue': 1}}
            }},
            'fields': 'userEnteredFormat(backgroundColor,textFormat)'
        }})

    # 열 너비
    col_widths = [110, 80, 130, 50, 240, 340, 200]
    for i, w in enumerate(col_widths):
        fmt_requests.append({'updateDimensionProperties': {
            'range': {'sheetId': sheet_id, 'dimension': 'COLUMNS',
                      'startIndex': i, 'endIndex': i + 1},
            'properties': {'pixelSize': w}, 'fields': 'pixelSize'
        }})

    sheets.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID, body={'requests': fmt_requests}
    ).execute()

    print(f'  → 경쟁사 분석 {data_row_count}행 + 구분선 {len(sep_rows)}개 추가 완료')


def _is_company(ad, name, url_key):
    """광고 dict에서 업체명/URL로 회사 판별"""
    combined = (ad.get('title','') + ad.get('desc','') + ad.get('url','')).lower()
    return name.lower() in combined or url_key.lower() in combined


def analyze_ads_with_claude(all_ads):
    """Claude API로 경쟁사 대비 광고 T&D 분석 리포트 생성"""
    lines = []
    for kw in KEYWORDS:
        for device in ['PC', '모바일']:
            ads = all_ads.get(kw, {}).get(device, [])
            if not ads:
                continue
            lines.append(f'\n[{kw} · {device}]')
            for ad in ads:
                tag = ''
                if _is_company(ad, MY_COMPANY, MY_URL_KEY):
                    tag = '★카모아'
                elif _is_company(ad, KEY_RIVAL, RIVAL_URL_KEY):
                    tag = '▶돌하루팡'
                lines.append(
                    f"  {ad.get('rank','?')}위 {tag} | "
                    f"제목: {ad.get('title','')} | "
                    f"설명: {ad.get('desc','')[:80]} | "
                    f"URL: {ad.get('url','')}"
                )

    ad_text = '\n'.join(lines)

    prompt = f"""당신은 제주 렌트카 업계 마케팅 전문가입니다.
아래는 오늘({TODAY}) 네이버 파워링크 광고 현황입니다.

{ad_text}

다음 4가지를 분석해서 수요일 마케팅 회의용 보고서를 작성해주세요.

1. 📍 카모아 순위 현황
   - 키워드별 PC/모바일 노출 순위 요약
   - 돌하루팡 대비 순위 비교

2. 🔍 T&D 경쟁력 분석
   - 카모아 vs 돌하루팡 광고 제목/설명 문구 비교
   - 경쟁사 전체에서 자주 쓰이는 키워드/소구점 정리

3. ⚡ 경쟁사 주목 포인트
   - 돌하루팡 또는 다른 경쟁사 중 눈에 띄는 광고 전략

4. 💡 카모아 개선 제안
   - 현재 T&D에서 아쉬운 점
   - 구체적인 개선 제목/설명 예시 2~3가지

각 섹션 제목을 명확히 쓰고, 실무에서 바로 활용할 수 있도록 작성해주세요."""

    print('  Claude API 분석 중...')
    import anthropic
    client = anthropic.Anthropic()
    message = client.messages.create(
        model='claude-opus-4-6',
        max_tokens=4096,
        messages=[{'role': 'user', 'content': prompt}]
    )
    return message.content[0].text.strip()


def _rank_color(rank_str):
    """순위 문자열 → 배경색 dict 반환"""
    if rank_str == '미노출':
        return {'red': 0.88, 'green': 0.88, 'blue': 0.88}
    try:
        r = int(rank_str)
        if r <= 3:   return {'red': 0.72, 'green': 0.93, 'blue': 0.73}  # 연초록
        if r <= 6:   return {'red': 1.0,  'green': 0.95, 'blue': 0.60}  # 연노랑
        return           {'red': 1.0,  'green': 0.78, 'blue': 0.78}      # 연빨강
    except ValueError:
        return {'red': 1.0, 'green': 1.0, 'blue': 1.0}


def _clean_md(text):
    """마크다운 기호 제거 후 반환"""
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)   # **bold** → bold
    text = re.sub(r'^#{1,3}\s*', '', text)           # ## 헤더 prefix 제거
    return text.strip()


def _col_to_letter(n):
    """0-based 열 인덱스 → 시트 열 문자 (A, B, ..., Z, AA, ...)"""
    result = ''
    n += 1
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


# 고정 행 구조 (1-based)
# Row 1 : 날짜 헤더
# Row 2 : 열 헤더 (카모아순위 | 돌하루팡순위)
# Row 3~10 : 키워드×기기 순위 데이터 (4키워드 × 2기기 = 8행)
# Row 11 : 구분 여백
# Row 12 : AI 분석 텍스트 (셀 하나에 전체 저장, 텍스트 줄바꿈)
_KW_DEVICE_PAIRS = [(kw, dev) for kw in KEYWORDS for dev in ['PC', '모바일']]


def write_ad_analysis_tab(sheets, all_ads):
    """광고 분석 리포트 탭 — 날짜별 행으로 누적 (수직 쌓기)"""
    tab = SHEET_TAB_ANALYSIS
    sheet_id = get_or_create_tab(sheets, tab)

    # ── 현재 마지막 행 파악 (A열 기준) ──────────────────
    col_a = sheets.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range=f"'{tab}'!A:A"
    ).execute().get('values', [])

    is_first = len(col_a) == 0
    next_row_0 = len(col_a)   # 0-based 인덱스

    # ── 처음 실행: 타이틀 + 헤더 행 작성 ─────────────────
    if is_first:
        header_rows = [
            ['📊 광고 분석 리포트', '', '', '', ''],
            ['날짜', '키워드', '기기', f'{MY_COMPANY} 순위', f'{KEY_RIVAL} 순위'],
        ]
        sheets.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=f"'{tab}'!A1",
            valueInputOption='USER_ENTERED',
            body={'values': header_rows}
        ).execute()
        next_row_0 = 2   # 타이틀(0) + 헤더(1) 다음

        # 열 너비: 날짜 110 / 키워드 130 / 기기 70 / 카모아 90 / 경쟁사 90
        fmt_init = []
        for idx, w in enumerate([110, 130, 70, 90, 90]):
            fmt_init.append({'updateDimensionProperties': {
                'range': {'sheetId': sheet_id, 'dimension': 'COLUMNS',
                          'startIndex': idx, 'endIndex': idx + 1},
                'properties': {'pixelSize': w}, 'fields': 'pixelSize'
            }})
        sheets.spreadsheets().batchUpdate(
            spreadsheetId=SHEET_ID, body={'requests': fmt_init}
        ).execute()

    # ── 이번 날짜 데이터 행 구성 ────────────────────────
    data_rows = []
    rank_info = []   # (row_0based, my_rank, rival_rank)

    for i, (kw, device) in enumerate(_KW_DEVICE_PAIRS):
        ads = all_ads.get(kw, {}).get(device, [])
        my_rank    = next((str(a['rank']) for a in ads if _is_company(a, MY_COMPANY, MY_URL_KEY)), '미노출')
        rival_rank = next((str(a['rank']) for a in ads if _is_company(a, KEY_RIVAL, RIVAL_URL_KEY)), '미노출')
        data_rows.append([f'📅 {TODAY}', kw, device, my_rank, rival_rank])
        rank_info.append((next_row_0 + i, my_rank, rival_rank))

    # AI 분석 행
    analysis = analyze_ads_with_claude(all_ads)
    ai_row_0 = next_row_0 + len(_KW_DEVICE_PAIRS)
    data_rows.append(['', '🤖 AI 분석', '', _clean_md(analysis), ''])

    # 구분 빈 행
    data_rows.append(['', '', '', '', ''])

    # ── 시트에 쓰기 ─────────────────────────────────────
    sheets.spreadsheets().values().update(
        spreadsheetId=SHEET_ID,
        range=f"'{tab}'!A{next_row_0 + 1}",   # 1-based
        valueInputOption='USER_ENTERED',
        body={'values': data_rows}
    ).execute()

    # ── 서식 batchUpdate ─────────────────────────────────
    fmt = []

    GREEN_DARK  = {'red': 0.01, 'green': 0.78, 'blue': 0.35}
    GRAY_HEADER = {'red': 0.85, 'green': 0.85, 'blue': 0.85}
    GRAY_BG     = {'red': 0.97, 'green': 0.97, 'blue': 0.97}
    WHITE       = {'red': 1.0,  'green': 1.0,  'blue': 1.0}
    BLACK       = {'red': 0.1,  'green': 0.1,  'blue': 0.1}
    NUM_COLS    = 5

    def row_fmt(ri, bg, bold=False, font_size=10, fg=None, col_end=NUM_COLS):
        fmt.append({'repeatCell': {
            'range': {'sheetId': sheet_id,
                      'startRowIndex': ri, 'endRowIndex': ri + 1,
                      'startColumnIndex': 0, 'endColumnIndex': col_end},
            'cell': {'userEnteredFormat': {
                'backgroundColor': bg,
                'textFormat': {'bold': bold, 'fontSize': font_size,
                               'foregroundColor': fg or BLACK}
            }},
            'fields': 'userEnteredFormat(backgroundColor,textFormat)'
        }})

    # 처음 실행 시: 타이틀(row0) + 헤더(row1) 서식
    if is_first:
        row_fmt(0, GREEN_DARK, bold=True, font_size=12, fg=WHITE)
        row_fmt(1, GRAY_HEADER, bold=True)

    # 날짜 블록 첫 행 배경 (연초록)
    LIGHT_GREEN = {'red': 0.85, 'green': 0.96, 'blue': 0.87}
    for ri in range(next_row_0, next_row_0 + len(_KW_DEVICE_PAIRS)):
        row_fmt(ri, LIGHT_GREEN if ri == next_row_0 else WHITE)

    # 순위 셀 색상 (D·E열 = index 3·4)
    for row_0, my_rank, rival_rank in rank_info:
        for col_idx, rank_val in [(3, my_rank), (4, rival_rank)]:
            fmt.append({'repeatCell': {
                'range': {'sheetId': sheet_id,
                          'startRowIndex': row_0, 'endRowIndex': row_0 + 1,
                          'startColumnIndex': col_idx, 'endColumnIndex': col_idx + 1},
                'cell': {'userEnteredFormat': {
                    'backgroundColor': _rank_color(rank_val),
                    'textFormat': {'bold': True, 'fontSize': 10},
                    'horizontalAlignment': 'CENTER',
                }},
                'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
            }})

    # AI 분석 행 (D열에 텍스트, wrap)
    fmt.append({'repeatCell': {
        'range': {'sheetId': sheet_id,
                  'startRowIndex': ai_row_0, 'endRowIndex': ai_row_0 + 1,
                  'startColumnIndex': 0, 'endColumnIndex': NUM_COLS},
        'cell': {'userEnteredFormat': {
            'wrapStrategy': 'WRAP',
            'backgroundColor': GRAY_BG,
            'textFormat': {'fontSize': 9},
        }},
        'fields': 'userEnteredFormat(wrapStrategy,backgroundColor,textFormat)'
    }})
    # AI 분석 행 높이
    fmt.append({'updateDimensionProperties': {
        'range': {'sheetId': sheet_id, 'dimension': 'ROWS',
                  'startIndex': ai_row_0, 'endIndex': ai_row_0 + 1},
        'properties': {'pixelSize': 200}, 'fields': 'pixelSize'
    }})

    sheets.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID, body={'requests': fmt}
    ).execute()

    print(f'  → 광고 분석 리포트 작성 완료 (행 {next_row_0 + 1}~{ai_row_0 + 1})')


def run_capture_flow(sheets, drive, mode_pc, mode_mobile, tab_name, label, brand_data=None):
    """캡처 → URL 생성 → HTML 스크린샷 → 시트 업데이트 공통 플로우"""
    pc_kw      = {kw: fpath for kw, (fpath, fname) in mode_pc.items()}
    mobile_kw  = {kw: fpath for kw, (fpath, fname) in mode_mobile.items()}

    pc_urls, mobile_urls = {}, {}
    for kw, (fpath, fname) in mode_pc.items():
        pc_urls[kw] = make_public_url(drive, upload_to_drive(drive, fpath, fname))
    for kw, (fpath, fname) in mode_mobile.items():
        mobile_urls[kw] = make_public_url(drive, upload_to_drive(drive, fpath, fname))

    print(f'  HTML 스크린샷 생성 ({label})...')
    pc_path,  pc_fname  = render_html_screenshot(pc_kw,    f'{label}_PC',    brand_data)
    mob_path, mob_fname = render_html_screenshot(mobile_kw, f'{label}_모바일', brand_data)

    def make_pub(fpath, fname):
        fid = upload_to_drive(drive, fpath, fname)
        drive.permissions().create(fileId=fid, body={'role': 'reader', 'type': 'anyone'}).execute()
        return f'https://drive.google.com/uc?export=view&id={fid}'

    pc_url  = make_pub(pc_path, pc_fname)
    mob_url = make_pub(mob_path, mob_fname)
    print(f'  PC: {pc_url}')
    print(f'  모바일: {mob_url}')

    sheet_id = get_or_create_tab(sheets, tab_name)
    append_to_sheet(sheets, sheet_id, pc_urls, mobile_urls, pc_url, mob_url, tab=tab_name)


def scrape_content_brands():
    """
    4개 키워드 × PC + 모바일: 콘텐츠 지면 전체 섹션 분석
    - 브랜드콘텐츠: 작성자 = 업체명 (명시적)
    - 인플루언서/인기글/카페글: 실제 글 URL에 방문해 본문 전체에서 업체명 탐지
    """

    RENTAL_BRANDS = [
        '카모아', '돌하루팡', '쏘카', 'SK렌터카', 'SK렌트카', '제주속으로',
        '제주패스', '제주엔젤', 'JAR렌트카', 'JAR렌트', '하나렌트카', '제주닷컴',
        '제주로렌트카', '제주원렌트카', '제주빌리고', '빌리고', '제주썬렌트카',
        '허브렌트카', '롯데렌터카', '롯데렌트카', '딜카', '그린카', '제주렌트카본사',
    ]
    DATE_RE = re.compile(
        r'(\d+주 전|\d+일 전|\d+시간 전|\d+분 전|\d+개월 전'
        r'|\d{4}\.\d{2}\.\d{2}\.|\d{4}\.\d{2}\.)'
    )

    def detect_brands(text):
        found = []
        for b in RENTAL_BRANDS:
            if b in text and b not in found:
                found.append(b)
        for m in re.findall(r'[\w가-힣]{2,6}렌[트터]카', text):
            if m not in found and m not in ['제주렌트카', '제주도렌트카', '제주렌터카', '제주도렌터카']:
                found.append(m)
        return ', '.join(found)

    # JS: 브랜드콘텐츠 Y 이후 섹션별 아이템 + URL 추출
    SECTIONS_JS = r"""() => {
        let brandY = 999999;
        const bSels = ['.fds-comps-footer-more-subject', '.sds-comps-footer-label',
                       '[class*=footer-more-subject]'];
        for (const sel of bSels) {
            for (const el of document.querySelectorAll(sel)) {
                if ((el.innerText || '').includes('브랜드')) {
                    const b = el.closest('.sc_new') || el.closest('._fe_view_root');
                    if (b) { brandY = b.getBoundingClientRect().top + window.scrollY; break; }
                }
            }
            if (brandY < 999999) break;
        }

        const results = [];
        document.querySelectorAll('.sc_new, ._fe_view_root').forEach(sec => {
            const y = sec.getBoundingClientRect().top + window.scrollY;
            if (y < brandY - 50) return;

            const h = sec.querySelector('h2, h3');
            const secText = sec.innerText.trim();
            const heading = h ? h.innerText.trim() : secText.split('\n')[0].trim();
            if (!heading || heading.length > 60) return;
            if (heading.includes('함께 많이') || heading.includes('함께 찾는')) return;

            // 섹션 유형
            let stype = heading.substring(0, 15);
            if (heading.includes('브랜드 콘텐츠')) stype = '브랜드콘텐츠';
            else if (heading.includes('인플루언서'))  stype = '인플루언서';
            else if (heading.includes('인기글'))       stype = '인기글';
            else if (heading.includes('카페'))         stype = '카페글';
            else if (heading.includes('함께 보면'))    stype = '연관업체';

            // 텍스트 기반 아이템 파싱 (날짜 패턴 기준)
            const lines = secText.split('\n').map(l => l.trim()).filter(Boolean);
            const dateRe = /(\d+주 전|\d+일 전|\d+시간 전|\d+분 전|\d+개월 전|\d{4}\.\d{2})/;
            const skipWords = ['관련','브랜드 콘텐츠','인기 카페글','인기글','함께','더보기','팬하기','Keep','저장'];

            const items = [];
            for (let i = 0; i < lines.length; i++) {
                if (!dateRe.test(lines[i])) continue;
                const author = i > 0 ? lines[i-1] : '';
                const title  = i+1 < lines.length ? lines[i+1] : '';
                if (!author || author.length > 35 || skipWords.some(w => author.includes(w))) continue;
                if (!title) continue;

                // 해당 아이템의 링크 찾기: 제목 텍스트를 포함하는 a 태그
                let href = '';
                const titleShort = title.substring(0, 15);
                for (const a of sec.querySelectorAll('a[href]')) {
                    if ((a.innerText || '').includes(titleShort) || (a.title || '').includes(titleShort)) {
                        href = a.href;
                        break;
                    }
                }
                // fallback: 섹션 내 모든 링크에서 블로그/카페/인플루언서 URL 순서대로
                if (!href) {
                    const links = Array.from(sec.querySelectorAll(
                        'a[href*="blog.naver"], a[href*="cafe.naver"], a[href*="post.naver"], a[href*="in.naver"]'
                    ));
                    if (links.length > items.length) href = links[items.length].href;
                }

                // snippet: 제목 다음 줄 (날짜 아니고 skipWords 아닌 경우)
                let snippet = '';
                if (i+2 < lines.length && !dateRe.test(lines[i+2])
                        && !skipWords.some(w => lines[i+2].includes(w))
                        && lines[i+2].length < 120) {
                    snippet = lines[i+2];
                }
                items.push({ author, title: title.substring(0, 80), date: lines[i], href, snippet });
            }

            if (items.length > 0) results.push({ heading, stype, items });
        });
        return results;
    }"""

    def visit_and_detect(pg, href, snippet=''):
        """실제 글 URL에 방문해 본문 전체에서 업체명 탐지"""
        if not href or not href.startswith('http'):
            return detect_brands(snippet)

        # in.naver.com → blog.naver.com 으로 변환 시도
        visit_url = href
        m = re.match(r'https?://in\.naver\.com/([^/]+)/contents/[^/]+/(\d+)', href)
        if m:
            visit_url = f'https://blog.naver.com/{m.group(1)}/{m.group(2)}'

        try:
            pg.goto(visit_url, wait_until='networkidle', timeout=20000)
            pg.wait_for_timeout(2500)
            # 스크롤해서 lazy 콘텐츠 로드
            pg.evaluate('() => window.scrollBy(0, 600)')
            pg.wait_for_timeout(800)

            body_text = pg.evaluate(r'''() => {
                // 1) 네이버 블로그 iframe (se-main-container, postViewArea)
                for (const f of document.querySelectorAll('iframe')) {
                    try {
                        const doc = f.contentDocument
                                 || (f.contentWindow && f.contentWindow.document);
                        if (doc && doc.body) {
                            const t = doc.body.innerText || '';
                            if (t.length > 200) return t;
                        }
                    } catch(e) {}
                }
                // 2) 본문 innerText
                const bt = document.body ? (document.body.innerText || '') : '';
                if (bt.length > 100) return bt;
                // 3) meta description / og:description fallback
                const m = document.querySelector(
                    'meta[name="description"], meta[property="og:description"]');
                return m ? (m.getAttribute('content') || '') : '';
            }''')

            found = detect_brands(body_text[:5000])
            if found:
                return found
        except Exception:
            pass

        # URL 방문 실패 시 스니펫으로 fallback
        return detect_brands(snippet)

    all_content = {}

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)

        for device, vp, scale, ua in [
            ('PC',   PC_VIEWPORT,     1, None),
            ('모바일', MOBILE_VIEWPORT, 2, MOBILE_UA),
        ]:
            # 검색 결과 페이지용 컨텍스트
            ctx_kw = dict(viewport=vp, device_scale_factor=scale, storage_state=None)
            if ua:
                ctx_kw['user_agent'] = ua
            ctx = browser.new_context(**ctx_kw)
            pg_search = ctx.new_page()
            pg_search.set_extra_http_headers({'Accept-Language': 'ko-KR,ko;q=0.9'})

            # 개별 글 방문용 (PC UA 고정 - 모바일 블로그도 PC로 더 잘 읽힘)
            ctx_visit = browser.new_context(
                viewport=PC_VIEWPORT, device_scale_factor=1, storage_state=None
            )
            pg_visit = ctx_visit.new_page()
            pg_visit.set_extra_http_headers({'Accept-Language': 'ko-KR,ko;q=0.9'})

            for kw in KEYWORDS:
                pg_search.goto(f'https://search.naver.com/search.naver?query={kw}',
                               wait_until='networkidle', timeout=30000)
                pg_search.wait_for_timeout(2000)
                pg_search.evaluate(r"""() => new Promise(resolve => {
                    let y=0;const t=setInterval(()=>{window.scrollBy(0,600);y+=600;
                    if(y>=document.body.scrollHeight){window.scrollTo(0,0);clearInterval(t);resolve();}},100);
                })""")
                pg_search.wait_for_timeout(2000)

                sections = pg_search.evaluate(SECTIONS_JS)
                items = []
                for sec in sections:
                    for it in sec['items']:
                        if sec['stype'] == '브랜드콘텐츠':
                            brand = it['author']   # 명시적 업체명
                        else:
                            # 실제 글에 방문해서 본문에서 탐지 (스니펫을 fallback으로)
                            fallback_text = it.get('snippet', '') + ' ' + it.get('title', '')
                            brand = visit_and_detect(pg_visit, it['href'], fallback_text)
                        items.append({
                            'section': sec['stype'],
                            'author':  it['author'],
                            'title':   it['title'],
                            'date':    it['date'],
                            'brand':   brand,
                        })

                all_content.setdefault(kw, {})[device] = items
                sec_names = list({it['section'] for it in items})
                print(f'    [콘텐츠/{device}] {kw}: {len(items)}개 ({", ".join(sec_names)})')

            ctx.close()
            ctx_visit.close()

        browser.close()

    return all_content


def write_content_brands_tab(sheets, all_content):
    """콘텐츠 업체 현황 탭 – 섹션별 어느 업체 콘텐츠인지 누적 기록"""
    tab = SHEET_TAB_CONTENT_BIZ
    sheet_id = get_or_create_tab(sheets, tab)

    # 헤더
    existing = sheets.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range=f"'{tab}'!A1:H1"
    ).execute().get('values', [])

    HEADERS = ['날짜', '기기', '키워드', '섹션 유형', '작성자/계정', '연관 업체', '콘텐츠 제목', '게시일']
    if not existing or (existing[0] and existing[0][0] != '날짜'):
        sheets.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=f"'{tab}'!A1",
            valueInputOption='USER_ENTERED',
            body={'values': [HEADERS]}
        ).execute()
        sheets.spreadsheets().batchUpdate(
            spreadsheetId=SHEET_ID,
            body={'requests': [{'repeatCell': {
                'range': {'sheetId': sheet_id, 'startRowIndex': 0, 'endRowIndex': 1},
                'cell': {'userEnteredFormat': {
                    'backgroundColor': {'red': 0.13, 'green': 0.52, 'blue': 0.96},
                    'textFormat': {'bold': True,
                                   'foregroundColor': {'red': 1, 'green': 1, 'blue': 1}}
                }},
                'fields': 'userEnteredFormat(backgroundColor,textFormat)'
            }}]}
        ).execute()

    col_a = sheets.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range=f"'{tab}'!A:A"
    ).execute().get('values', [])
    next_row = len(col_a) + 1

    rows = []
    sep_rows = []

    for kw in KEYWORDS:
        sep_rows.append(next_row + len(rows))
        rows.append([f'▶ {kw}', '', '', '', '', '', '', ''])

        for device in ['PC', '모바일']:
            items = all_content.get(kw, {}).get(device, [])
            if not items:
                rows.append([TODAY, device, kw, '-', '-', '데이터 없음', '', ''])
            else:
                for it in items:
                    rows.append([
                        TODAY,
                        device,
                        kw,
                        it.get('section', ''),
                        it.get('author', ''),
                        it.get('brand', ''),
                        it.get('title', ''),
                        it.get('date', ''),
                    ])

    if rows:
        sheets.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=f"'{tab}'!A{next_row}",
            valueInputOption='USER_ENTERED',
            body={'values': rows}
        ).execute()

    # 서식: 구분선 + 열 너비
    fmt_requests = []
    for r in sep_rows:
        fmt_requests.append({'repeatCell': {
            'range': {'sheetId': sheet_id,
                      'startRowIndex': r - 1, 'endRowIndex': r,
                      'startColumnIndex': 0, 'endColumnIndex': 8},
            'cell': {'userEnteredFormat': {
                'backgroundColor': {'red': 0.25, 'green': 0.25, 'blue': 0.25},
                'textFormat': {'bold': True, 'fontSize': 11,
                               'foregroundColor': {'red': 1, 'green': 1, 'blue': 1}}
            }},
            'fields': 'userEnteredFormat(backgroundColor,textFormat)'
        }})

    for i, w in enumerate([110, 80, 120, 100, 150, 160, 300, 110]):
        fmt_requests.append({'updateDimensionProperties': {
            'range': {'sheetId': sheet_id, 'dimension': 'COLUMNS',
                      'startIndex': i, 'endIndex': i + 1},
            'properties': {'pixelSize': w}, 'fields': 'pixelSize'
        }})

    sheets.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID, body={'requests': fmt_requests}
    ).execute()

    data_count = sum(1 for r in rows if r[1] and r[1] in ['PC', '모바일'])
    print(f'  → 콘텐츠 업체 현황 {data_count}행 + 구분선 {len(sep_rows)}개 추가 완료')


def analyze_content_with_claude(all_content):
    """Claude API로 콘텐츠 지면 분석 리포트 생성"""
    lines = []
    for kw in KEYWORDS:
        for device in ['PC', '모바일']:
            items = all_content.get(kw, {}).get(device, [])
            if not items:
                continue
            lines.append(f'\n[{kw} · {device}]')
            for it in items:
                lines.append(
                    f"  섹션={it.get('section','')} | "
                    f"작성자={it.get('author','')} | "
                    f"연관업체={it.get('brand','없음')} | "
                    f"제목={it.get('title','')} | "
                    f"게시일={it.get('date','')}"
                )

    content_text = '\n'.join(lines)

    prompt = f"""당신은 제주 렌트카 업계 마케팅 전문가입니다.
아래는 오늘({TODAY}) 네이버 콘텐츠 지면(브랜드콘텐츠/인플루언서/인기글/카페글) 노출 현황입니다.

{content_text}

다음 4가지를 분석해서 실무 보고서를 작성해주세요.

1. 📍 카모아 vs 돌하루팡 콘텐츠 노출 현황
   - 키워드별 PC/모바일 각각 몇 건 노출됐는지
   - 어느 섹션(브랜드콘텐츠/인플루언서 등)에서 노출됐는지

2. ⚠️ 카모아 공백 구간
   - 카모아가 전혀 안 잡히는 키워드·기기·섹션
   - 경쟁사만 노출되는 위험 구간

3. 🔍 경쟁사 콘텐츠 전략
   - 돌하루팡 또는 다른 경쟁사가 많이 노출되는 섹션/키워드
   - 상위 노출 콘텐츠 제목에서 공통으로 쓰이는 문구나 소구점

4. 💡 카모아 콘텐츠 개선 제안
   - 강화해야 할 섹션 유형 (브랜드콘텐츠/인플루언서 등)
   - 공략할 키워드와 콘텐츠 방향 제안

각 섹션 제목을 명확히 쓰고 실무에서 바로 활용 가능하게 작성해주세요."""

    print('  Claude API 콘텐츠 분석 중...')
    import anthropic
    client = anthropic.Anthropic()
    message = client.messages.create(
        model='claude-opus-4-6',
        max_tokens=4096,
        messages=[{'role': 'user', 'content': prompt}]
    )
    return message.content[0].text.strip()


def write_content_analysis_tab(sheets, all_content):
    """콘텐츠 분석 리포트 탭 — 날짜별 가로 누적 (3열씩)"""
    tab = SHEET_TAB_CONTENT_ANALYSIS
    sheet_id = get_or_create_tab(sheets, tab)

    # ── 다음 시작 열 계산 ────────────────────────────────
    row2 = sheets.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range=f"'{tab}'!2:2"
    ).execute().get('values', [[]])[0]

    existing_cols = len(row2)
    is_first = existing_cols <= 2
    start_col = max(2, existing_cols)

    # ── 처음 실행: A·B열 고정 레이블 ────────────────────
    if is_first:
        label_rows = [
            ['📊 콘텐츠 분석 리포트', ''],
            ['키워드', '기기'],
        ]
        for kw, dev in _KW_DEVICE_PAIRS:
            label_rows.append([kw, dev])
        label_rows.append(['', ''])
        label_rows.append(['📝 AI 분석', ''])

        sheets.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=f"'{tab}'!A1",
            valueInputOption='USER_ENTERED',
            body={'values': label_rows}
        ).execute()

    # ── 이번 날짜 데이터 (2열: 카모아 노출수 | 돌하루팡 노출수) ──
    col_data = []

    # Row 1: 날짜
    col_data.append([f'📅 {TODAY}', ''])

    # Row 2: 열 헤더
    col_data.append([f'{MY_COMPANY} 노출', f'{KEY_RIVAL} 노출'])

    # Rows 3-10: 키워드×기기별 노출 집계
    count_info = []
    for i, (kw, device) in enumerate(_KW_DEVICE_PAIRS):
        items = all_content.get(kw, {}).get(device, [])
        my_count    = sum(1 for it in items if MY_COMPANY in it.get('brand', ''))
        rival_count = sum(1 for it in items if KEY_RIVAL  in it.get('brand', ''))
        col_data.append([str(my_count), str(rival_count)])
        count_info.append((i + 2, my_count, rival_count))

    # Row 11: 여백
    col_data.append(['', ''])

    # Row 12: AI 분석
    analysis = analyze_content_with_claude(all_content)
    col_data.append([_clean_md(analysis), ''])

    # ── 시트에 쓰기 ─────────────────────────────────────
    col_letter = _col_to_letter(start_col)
    sheets.spreadsheets().values().update(
        spreadsheetId=SHEET_ID,
        range=f"'{tab}'!{col_letter}1",
        valueInputOption='USER_ENTERED',
        body={'values': col_data}
    ).execute()

    # ── 서식 ────────────────────────────────────────────
    fmt = []
    TEAL_DARK   = {'red': 0.00, 'green': 0.59, 'blue': 0.53}
    GRAY_HEADER = {'red': 0.85, 'green': 0.85, 'blue': 0.85}
    GRAY_BG     = {'red': 0.97, 'green': 0.97, 'blue': 0.97}
    WHITE       = {'red': 1.0,  'green': 1.0,  'blue': 1.0}
    BLACK       = {'red': 0.1,  'green': 0.1,  'blue': 0.1}
    GREEN_CELL  = {'red': 0.72, 'green': 0.93, 'blue': 0.73}
    RED_CELL    = {'red': 1.0,  'green': 0.78, 'blue': 0.78}

    cs, ce = start_col, start_col + 2

    def block_fmt(ri, bg, bold=False, font_size=10, fg=None):
        fmt.append({'repeatCell': {
            'range': {'sheetId': sheet_id,
                      'startRowIndex': ri, 'endRowIndex': ri + 1,
                      'startColumnIndex': cs, 'endColumnIndex': ce},
            'cell': {'userEnteredFormat': {
                'backgroundColor': bg,
                'textFormat': {'bold': bold, 'fontSize': font_size,
                               'foregroundColor': fg or BLACK}
            }},
            'fields': 'userEnteredFormat(backgroundColor,textFormat)'
        }})

    block_fmt(0, TEAL_DARK, bold=True, font_size=12, fg=WHITE)
    block_fmt(1, GRAY_HEADER, bold=True)

    # 카모아 노출수: 1이상=초록, 0=빨강 / 돌하루팡: 반대
    for row_offset, my_cnt, rival_cnt in count_info:
        for c, cnt, invert in [(cs, my_cnt, False), (cs + 1, rival_cnt, True)]:
            if invert:
                bg = RED_CELL if cnt >= 1 else GREEN_CELL
            else:
                bg = GREEN_CELL if cnt >= 1 else RED_CELL
            fmt.append({'repeatCell': {
                'range': {'sheetId': sheet_id,
                          'startRowIndex': row_offset, 'endRowIndex': row_offset + 1,
                          'startColumnIndex': c, 'endColumnIndex': c + 1},
                'cell': {'userEnteredFormat': {
                    'backgroundColor': bg,
                    'textFormat': {'bold': True, 'fontSize': 10}
                }},
                'fields': 'userEnteredFormat(backgroundColor,textFormat)'
            }})

    # Row 12: AI 분석 텍스트 줄바꿈
    fmt.append({'repeatCell': {
        'range': {'sheetId': sheet_id,
                  'startRowIndex': 11, 'endRowIndex': 12,
                  'startColumnIndex': cs, 'endColumnIndex': ce},
        'cell': {'userEnteredFormat': {
            'wrapStrategy': 'WRAP',
            'backgroundColor': GRAY_BG,
            'textFormat': {'fontSize': 9}
        }},
        'fields': 'userEnteredFormat(wrapStrategy,backgroundColor,textFormat)'
    }})

    if is_first:
        for i, w in enumerate([180, 70]):
            fmt.append({'updateDimensionProperties': {
                'range': {'sheetId': sheet_id, 'dimension': 'COLUMNS',
                          'startIndex': i, 'endIndex': i + 1},
                'properties': {'pixelSize': w}, 'fields': 'pixelSize'
            }})
        fmt.append({'updateDimensionProperties': {
            'range': {'sheetId': sheet_id, 'dimension': 'ROWS',
                      'startIndex': 11, 'endIndex': 12},
            'properties': {'pixelSize': 220}, 'fields': 'pixelSize'
        }})

    for i, w in enumerate([90, 90], start=cs):
        fmt.append({'updateDimensionProperties': {
            'range': {'sheetId': sheet_id, 'dimension': 'COLUMNS',
                      'startIndex': i, 'endIndex': i + 1},
            'properties': {'pixelSize': w}, 'fields': 'pixelSize'
        }})

    sheets.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID, body={'requests': fmt}
    ).execute()

    print(f'  → 콘텐츠 분석 리포트 작성 완료 (열 {_col_to_letter(cs)}~{_col_to_letter(ce - 1)})')


def main():
    print(f'=== 네이버 지면 캡처 [{NOW}] ===')
    sheets, drive = get_services()

    # ── 지면확보 (플레이스 위) ────────────────────────────
    print('[1/8] 지면확보 PC 캡처...')
    pc_paths = capture_pages(mode='pc', content=False)
    print('[2/8] 지면확보 모바일 캡처...')
    mob_paths = capture_pages(mode='mobile', content=False)
    print('[3/8] 지면확보 시트 업데이트...')
    run_capture_flow(sheets, drive, pc_paths, mob_paths, SHEET_TAB, 'PC')

    # ── 콘텐츠 지면 (플레이스 아래) ──────────────────────
    print('[4/10] 콘텐츠 지면 PC 캡처...')
    pc_content  = capture_pages(mode='pc',     content=True)
    print('[5/10] 콘텐츠 지면 모바일 캡처...')
    mob_content = capture_pages(mode='mobile', content=True)
    print('[6/10] 콘텐츠 지면 시트 업데이트...')
    run_capture_flow(sheets, drive, pc_content, mob_content, SHEET_TAB_CONTENT, '콘텐츠')

    # ── 콘텐츠 업체 현황 + 분석 리포트 ──────────────────
    print('[7/10] 콘텐츠 업체 현황 수집...')
    all_content = scrape_content_brands()
    write_content_brands_tab(sheets, all_content)
    print('[8/10] 콘텐츠 분석 리포트 생성...')
    write_content_analysis_tab(sheets, all_content)

    # ── 경쟁사 분석 (파워링크) ───────────────────────────
    print('[9/10] 경쟁사 광고 분석...')
    competitor_ads = scrape_competitor_ads()
    write_competitor_tab(sheets, competitor_ads)

    # ── 광고 분석 리포트 (카모아 vs 경쟁사) ──────────────
    print('[10/10] 광고 분석 리포트 생성...')
    write_ad_analysis_tab(sheets, competitor_ads)

    sheet_url = f'https://docs.google.com/spreadsheets/d/{SHEET_ID}'
    print(f'=== 완료! ===')
    print(f'시트: {sheet_url}')


if __name__ == '__main__':
    main()
