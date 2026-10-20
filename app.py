import streamlit as st
import pandas as pd
import datetime
import html
import streamlit.components.v1 as components
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

import gspread
from google.oauth2.service_account import Credentials

# 1. 화면 기본 설정
st.set_page_config(page_title="중점추진과제 WBS", layout="wide")

# ==========================================
# 디자인 토큰 (정수기 업체 → 물/잉크 톤)
#  ink #0F2A3D / water #0E7490 / mist #EEF3F6 / line #D5E0E7 / today #E5484D
# ==========================================
C_INK = "#0F2A3D"
C_WATER = "#0E7490"
C_PLAN = "#B8C7D2"        # 상세 계획 막대
C_ACT = "#0E7490"         # 상세 실적 막대
C_CAT_PLAN = "#8DA2B2"    # 대분류 계획 막대
C_CAT_ACT = "#0B4F63"     # 대분류 실적 막대

ICON_USER = "<svg width='14' height='14' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2'/><circle cx='12' cy='7' r='4'/></svg>"
ICON_CAL = "<svg width='14' height='14' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><rect x='3' y='4' width='18' height='18' rx='2'/><line x1='16' y1='2' x2='16' y2='6'/><line x1='8' y1='2' x2='8' y2='6'/><line x1='3' y1='10' x2='21' y2='10'/></svg>"
ICON_TASK = "<svg width='14' height='14' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><polyline points='9 11 12 14 22 4'/><path d='M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11'/></svg>"

st.markdown("""
<style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/variable/pretendardvariable-dynamic-subset.min.css');

    .stApp, .stApp button, .stApp input, .stApp textarea, .stApp h1, .stApp h2, .stApp h3 {
        font-family: 'Pretendard Variable', Pretendard, -apple-system, 'Noto Sans KR', sans-serif;
    }
    .stApp { background-color: #F3F6F8; }
    .block-container {
        padding-top: 3.5rem !important;
        padding-bottom: 1.5rem !important;
        min-width: 1200px !important;
    }

    /* ---------- 버튼 ---------- */
    div[data-testid="stButton"] button,
    div[data-testid="stDownloadButton"] button,
    div[data-testid="stLinkButton"] a {
        white-space: nowrap !important;
        word-break: keep-all !important;
        border-radius: 10px !important;
        min-height: 34px !important;
        height: 34px !important;
        padding: 2px 10px !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        gap: 4px !important;
        border: 1px solid #D5E0E7 !important;
        background: #FFFFFF !important;
        color: #0F2A3D !important;
        font-weight: 600 !important;
        box-shadow: none !important;
        transition: background .15s, border-color .15s;
    }
    div[data-testid="stButton"] button:hover,
    div[data-testid="stDownloadButton"] button:hover,
    div[data-testid="stLinkButton"] a:hover {
        background: #EEF3F6 !important;
        border-color: #0E7490 !important;
        color: #0E7490 !important;
    }
    div[data-testid="stButton"] button p,
    div[data-testid="stDownloadButton"] button p,
    div[data-testid="stLinkButton"] a p {
        font-size: 13px !important;
        line-height: 1.2 !important;
        margin: 0 !important;
    }
    div[data-testid="stButton"] button[kind="primary"],
    div[data-testid="stButton"] button[data-testid="stBaseButton-primary"] {
        background: #0E7490 !important;
        border-color: #0E7490 !important;
        color: #FFFFFF !important;
    }
    div[data-testid="stButton"] button[kind="primary"]:hover,
    div[data-testid="stButton"] button[data-testid="stBaseButton-primary"]:hover {
        background: #0B5C73 !important;
        border-color: #0B5C73 !important;
        color: #FFFFFF !important;
    }
    div[data-testid="stButton"] button[kind="primary"] *,
    div[data-testid="stButton"] button[data-testid="stBaseButton-primary"] * { color: #FFFFFF !important; }

    /* ---------- 대시보드 ---------- */
    .dash-title { margin: 0; font-size: 30px; font-weight: 800; color: #0F2A3D; letter-spacing: -0.5px; line-height: 1.2; }
    .dash-sub { margin: 4px 0 0 0; font-size: 14px; color: #5B7083; }
    .kpi-row { display: flex; gap: 12px; margin: 18px 0 22px 0; }
    .kpi { flex: 1; background: #FFFFFF; border: 1px solid #D5E0E7; border-radius: 14px; padding: 14px 18px; }
    .kpi-label { font-size: 12.5px; color: #5B7083; margin-bottom: 4px; }
    .kpi-value { font-size: 26px; font-weight: 800; color: #0F2A3D; letter-spacing: -0.5px; }
    .kpi-value small { font-size: 14px; font-weight: 600; color: #5B7083; margin-left: 3px; }

    div[data-testid="stVerticalBlockBorderWrapper"]:has(.pc-title) {
        background: #FFFFFF;
        border: 1px solid #D5E0E7 !important;
        border-radius: 16px !important;
        box-shadow: 0 1px 2px rgba(15,42,61,.05);
    }
    .kpi { flex: 0 0 240px !important; }

    /* 카드 클릭: 숨김 버튼('go_')을 JS가 대신 눌러 줌 (오버레이를 쓰지 않으므로 다른 버튼을 가리지 않음) */
    div[class*="st-key-go_"] { position: absolute !important; width: 1px !important; height: 1px !important; overflow: hidden !important; opacity: 0 !important; pointer-events: none !important; margin: 0 !important; }
    .wbs-card { cursor: pointer; transition: border-color .15s, box-shadow .15s; }
    .wbs-card:hover { border-color: #0E7490 !important; box-shadow: 0 4px 16px rgba(14,116,144,.14); }

    .pc-title { font-size: 18px; font-weight: 800; color: #0F2A3D; line-height: 1.3; letter-spacing: -0.3px; word-break: keep-all; }
    .pc-meta { display: flex; gap: 14px; margin: 6px 0 12px 0; font-size: 12.5px; color: #5B7083; }
    .pc-meta span { display: inline-flex; align-items: center; gap: 4px; }
    .pc-prog { height: 6px; background: #E3EBF0; border-radius: 3px; overflow: hidden; margin-top: 10px; }
    .pc-prog-bar { height: 100%; background: rgba(14,116,144,.35); border-radius: 3px; }
    .pc-prog-txt { font-size: 12px; color: #5B7083; margin: 5px 0 12px 0; }
    .pc-range { display: flex; align-items: center; gap: 4px; font-size: 12px; color: #5B7083; margin-bottom: 5px; }

    /* ---------- 상세 페이지 ---------- */
    .page-title { margin: 0; padding: 0; font-size: 32px; font-weight: 800; color: #0F2A3D; letter-spacing: -0.5px; line-height: 1.25; }
    .page-sub { font-size: 14px; color: #5B7083; margin: 4px 0 12px 0; }
    div[data-testid="stHorizontalBlock"]:has(.gantt-head-marker) {
        position: sticky; top: 3.75rem; z-index: 100;
        background: #F3F6F8; padding: 6px 0 4px 0;
    }
    .header-cell {
        height: 64px; display: flex; align-items: center; justify-content: center;
        font-weight: 700; font-size: 13px; color: #34495A;
        background-color: #E6EDF2; border-radius: 6px;
    }
    .cat-cell-bg { background-color: #E6EDF2; }
    .item-cell p, .item-cell-left p { margin-bottom: 0px !important; }
    .item-cell { min-height: 45px; display: flex; align-items: center; justify-content: center; font-size: 14px; color: #1F3547; }
    .item-cell-left { min-height: 45px; display: flex; align-items: center; justify-content: flex-start; padding-left: 15px; font-size: 14px; color: #1F3547; }
    .bar-container { position: relative; width: 100%; height: 45px; background-color: #fff; border: 1px solid #D5E0E7; border-top: none; box-sizing: border-box; overflow: visible !important; }
    .bar-est { position: absolute; top: 9px; height: 12px; background-color: #B8C7D2; border-radius: 4px; z-index: 1; }
    .bar-act { position: absolute; top: 24px; height: 12px; background-color: #0E7490; border-radius: 4px; z-index: 2; }
    hr.row-line { margin: 0; border: none; border-bottom: 1px solid #E3EBF0; }
    .legend { display: flex; gap: 18px; align-items: center; font-size: 12.5px; color: #5B7083; margin: 10px 0 4px 0; }
    .legend i { display: inline-block; width: 22px; height: 8px; border-radius: 4px; margin-right: 6px; vertical-align: middle; }
    .legend b { display: inline-block; width: 2px; height: 12px; background: #E5484D; margin-right: 6px; vertical-align: middle; }
</style>
""", unsafe_allow_html=True)

components.html("""
<script>
    const parentDoc = window.parent.document;
    parentDoc.body.classList.add('notranslate');
    parentDoc.body.setAttribute('translate', 'no');

    if (!window.parent.__wbsCardBinder) {
        window.parent.__wbsCardBinder = setInterval(function () {
            parentDoc.querySelectorAll('div[class*="st-key-go_"]').forEach(function (holder) {
                const card = holder.closest('div[data-testid="stVerticalBlockBorderWrapper"]');
                if (!card || card.dataset.wbsBound === '1') return;
                card.dataset.wbsBound = '1';
                card.classList.add('wbs-card');
                card.addEventListener('click', function (e) {
                    if (e.target.closest('button, a, input, textarea, [role="dialog"]')) return;
                    const btn = holder.querySelector('button');
                    if (btn) btn.click();
                });
            });
        }, 400);
    }
</script>
""", height=0)

# ==========================================
# 클라우드용 구글 시트 연동 엔진
# ==========================================
SHEET_URL = "https://docs.google.com/spreadsheets/d/1FBvFeUVOy0v2dSJOdZwnrx7aEggBNXkYaAlOGuPx54o/edit"

@st.cache_resource
def get_gspread_client():
    creds_dict = dict(st.secrets["gcp_service_account"])
    scopes = [
        "https://www.googleapis.com/auth/cloud-platform",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(credentials)

def load_data():
    try:
        gc = get_gspread_client()
        sh = gc.open_by_url(SHEET_URL)
        all_worksheets = sh.worksheets()
    except Exception as e:
        st.error(f"구글 시트 연결 실패: {e}")
        return {"projects": {}, "p_members": {}, "categories": {}}

    projects = {}
    p_members = {}
    categories = {}

    for ws in all_worksheets:
        try:
            all_values = ws.get_all_values()
        except:
            continue

        if len(all_values) <= 1:
            continue

        rows = all_values[1:]
        for row in rows:
            row = row + [""] * (12 - len(row))
            p_name, p_mem, cat, sub, man, dep, doc, est_s, est_e, act_s, act_e, memo = row[:12]

            if not p_name:
                p_name = ws.title

            if p_name not in projects:
                projects[p_name] = []
                categories[p_name] = []
                p_members[p_name] = p_mem

            if cat and cat not in categories[p_name]:
                categories[p_name].append(cat)

            if sub:
                projects[p_name].append({
                    "대분류": cat, "소분류": sub, "담당자": man, "부서": dep, "문서": doc,
                    "예상시작일": est_s, "예상종료일": est_e,
                    "실제시작일": act_s if act_s else None,
                    "실제종료일": act_e if act_e else None,
                    "메모": memo
                })

    return {"projects": projects, "p_members": p_members, "categories": categories}

def save_data(projects, p_members, categories):
    try:
        gc = get_gspread_client()
        sh = gc.open_by_url(SHEET_URL)
        existing_ws_titles = [ws.title for ws in sh.worksheets()]

        for p_name in projects:
            safe_p_name = p_name[:31]
            export_data = [["프로젝트명", "참여인원", "대분류", "세부업무명", "담당자", "관련부서", "관련문서", "계획시작일", "계획종료일", "실제시작일", "실제종료일", "비고(메모)"]]
            p_mem = p_members.get(p_name, "")
            p_cats = categories.get(p_name, [])
            p_tasks = projects.get(p_name, [])

            if not p_cats and not p_tasks:
                export_data.append([p_name, p_mem, "", "", "", "", "", "", "", "", "", ""])
            else:
                for cat in p_cats:
                    cat_tasks = [t for t in p_tasks if t['대분류'] == cat]
                    if not cat_tasks:
                        export_data.append([p_name, p_mem, cat, "", "", "", "", "", "", "", "", ""])
                    else:
                        for t in cat_tasks:
                            export_data.append([
                                p_name, p_mem, cat, t.get('소분류',''), t.get('담당자',''), t.get('부서',''), t.get('문서',''),
                                t.get('예상시작일',''), t.get('예상종료일',''),
                                t.get('실제시작일','') or '', t.get('실제종료일','') or '', t.get('메모','')
                            ])

            if safe_p_name in existing_ws_titles:
                ws = sh.worksheet(safe_p_name)
            else:
                ws = sh.add_worksheet(title=safe_p_name, rows=max(100, len(export_data)+10), cols=12)
                existing_ws_titles.append(safe_p_name)

            ws.clear()
            if export_data:
                ws.update(values=export_data, range_name="A1")

    except Exception as e:
        st.error(f"구글 시트 저장 실패: {e}")

if 'initialized' not in st.session_state:
    saved_data = load_data()
    st.session_state.projects = saved_data.get("projects", {})
    st.session_state.p_members = saved_data.get("p_members", {})
    st.session_state.categories = saved_data.get("categories", {})

    for proj_name, tasks in st.session_state.projects.items():
        if proj_name not in st.session_state.categories:
            cats = list(dict.fromkeys([d['대분류'] for d in tasks if '대분류' in d]))
            st.session_state.categories[proj_name] = cats

    st.session_state.current_page = 'Dashboard'
    st.session_state.close_dialog = False
    st.session_state.initialized = True

keys_to_init = ['p1_proj', 'p1_mem', 'p1_cat', 'p2_sub', 'p2_man', 'p2_dep', 'p2_doc', 'in_new_cat', 'msg_p1', 'msg_p2']
for k in keys_to_init:
    if k not in st.session_state: st.session_state[k] = ""
if 'p2_start' not in st.session_state: st.session_state.p2_start = datetime.date.today()
if 'p2_end' not in st.session_state: st.session_state.p2_end = datetime.date.today()

# ==========================================
# 간트 헤더 + 배경 격자
#  - 헤더(연/월/주)와 아래 막대 칸이 '같은 계산식'으로 폭을 나누므로 항상 정확히 맞습니다.
#  - 주(W) 칸은 실제 날짜 기준 (1~7일 / 8~14일 / 15~21일 / 22일~말일)
# ==========================================
def get_gantt_assets(min_date, max_date):
    min_date = pd.Timestamp(min_date).normalize().replace(day=1)
    max_date = pd.Timestamp(max_date).normalize() + pd.offsets.MonthEnd(0)
    total_days = max(1, (max_date - min_date).days + 1)

    months_data = []
    years_data = {}
    cur = min_date
    while cur <= max_date:
        nxt = cur + pd.offsets.MonthBegin(1)
        d_count = (min(nxt, max_date + pd.Timedelta(days=1)) - cur).days
        months_data.append((cur.year, cur.month, d_count))
        years_data[cur.year] = years_data.get(cur.year, 0) + d_count
        cur = nxt

    def week_widths(d_count):
        ws = [7, 7, 7, max(1, d_count - 21)]
        return [w / d_count * 100 for w in ws]

    year_html = "<div class='notranslate' style='display:flex; width:100%; height:22px; text-align:center; font-size:12px; font-weight:700; color:#0F2A3D; border-bottom:1px solid #D5E0E7; background-color:#E6EDF2;'>"
    for y, d_count in years_data.items():
        w_pct = d_count / total_days * 100
        year_html += f"<div style='width:{w_pct}%; border-right:1px solid #D5E0E7; box-sizing:border-box; line-height:22px; overflow:hidden; white-space:nowrap;'>{y}년</div>"
    year_html += "</div>"

    month_html = "<div class='notranslate' style='display:flex; width:100%; height:22px; text-align:center; font-size:12px; font-weight:700; color:#34495A; border-bottom:1px solid #D5E0E7; background-color:#F3F6F8;'>"
    for y, m, d_count in months_data:
        w_pct = d_count / total_days * 100
        month_html += f"<div style='width:{w_pct}%; border-right:1px solid #D5E0E7; box-sizing:border-box; line-height:22px; overflow:hidden; white-space:nowrap;'>{m}월</div>"
    month_html += "</div>"

    week_html = "<div class='notranslate' style='display:flex; width:100%; height:20px; text-align:center; font-size:10px; color:#6B7F90; background-color:#FFFFFF;'>"
    bg_grid_html = "<div style='position:absolute; top:0; left:0; width:100%; height:100%; display:flex; pointer-events:none; z-index:0;'>"

    for y, m, d_count in months_data:
        w_pct = d_count / total_days * 100
        widths = week_widths(d_count)
        week_html += f"<div style='width:{w_pct}%; display:flex; border-right:1px solid #D5E0E7; box-sizing:border-box;'>"
        bg_grid_html += f"<div style='width:{w_pct}%; display:flex; border-right:1px solid #D5E0E7; box-sizing:border-box;'>"
        for w in range(4):
            br = "border-right:1px dashed #E3EBF0;" if w < 3 else ""
            week_html += f"<div style='width:{widths[w]}%; {br} box-sizing:border-box; line-height:20px;'>{w+1}W</div>"
            bg_grid_html += f"<div style='width:{widths[w]}%; {br} box-sizing:border-box; height:100%;'></div>"
        week_html += "</div>"
        bg_grid_html += "</div>"

    week_html += "</div>"

    today = pd.Timestamp(datetime.date.today())
    if min_date <= today <= max_date:
        t_left = ((today - min_date).days + 0.5) / total_days * 100
        bg_grid_html += f"<div style='position:absolute; top:0; bottom:0; left:{t_left}%; width:2px; margin-left:-1px; background:#E5484D; opacity:.55;'></div>"
    bg_grid_html += "</div>"

    timeline_html = f"<div style='width:100%; box-sizing:border-box; display:flex; flex-direction:column; background-color:#fff; border:1px solid #D5E0E7; border-radius:6px 6px 0 0; overflow:hidden;'>{year_html}{month_html}{week_html}</div>"

    return timeline_html, bg_grid_html, min_date, max_date, total_days

@st.dialog("프로젝트 영구 삭제")
def delete_popup(p_name):
    st.write(f"**[{p_name}]** 프로젝트를 삭제하시겠습니까?")
    step1 = st.checkbox("네, 삭제를 진행하겠습니다.")
    if step1:
        st.error("정말 삭제하시겠습니까? 관련된 모든 세부 업무와 일정이 영구히 삭제됩니다.")
        c1, c2 = st.columns(2)
        if c1.button("예 (영구 삭제)", type="primary", use_container_width=True, icon=":material/delete:"):
            try:
                gc = get_gspread_client()
                sh = gc.open_by_url(SHEET_URL)
                if len(sh.worksheets()) > 1:
                    ws = sh.worksheet(p_name[:31])
                    sh.del_worksheet(ws)
                else:
                    ws = sh.worksheet(p_name[:31])
                    ws.clear()
            except: pass

            if p_name in st.session_state.projects: del st.session_state.projects[p_name]
            if p_name in st.session_state.p_members: del st.session_state.p_members[p_name]
            if p_name in st.session_state.categories: del st.session_state.categories[p_name]
            save_data(st.session_state.projects, st.session_state.p_members, st.session_state.categories)
            st.success("프로젝트가 삭제되었습니다.")
            st.rerun()
        if c2.button("아니오 (취소)", use_container_width=True):
            st.rerun()

def _save_project_data():
    p_name = st.session_state.p1_proj.strip()
    p_mem = st.session_state.p1_mem.strip()
    cat = st.session_state.p1_cat.strip()

    if not p_name:
        st.session_state.msg_p1 = "⚠️ 프로젝트명을 꼭 입력해주세요!"
        return False

    if p_name not in st.session_state.projects:
        st.session_state.projects[p_name] = []
        st.session_state.categories[p_name] = []

    st.session_state.p_members[p_name] = p_mem

    if cat:
        if cat not in st.session_state.categories[p_name]:
            st.session_state.categories[p_name].append(cat)
            st.session_state.msg_p1 = f"대분류 '{cat}' 추가 완료!"
        else:
            st.session_state.msg_p1 = "⚠️ 이미 존재하는 대분류입니다."
    else:
        st.session_state.msg_p1 = f"프로젝트 '{p_name}' 생성 완료!"

    save_data(st.session_state.projects, st.session_state.p_members, st.session_state.categories)
    st.session_state.p1_cat = ""
    return True

def cb_save_p1_continuous(): _save_project_data()

def cb_save_p1_and_close():
    if _save_project_data():
        st.session_state.msg_p1 = ""
        st.session_state.p1_proj = ""
        st.session_state.p1_mem = ""
        st.session_state.p1_cat = ""
        st.session_state.close_dialog = True

def cb_close_p1():
    st.session_state.msg_p1 = ""
    st.session_state.p1_proj = ""
    st.session_state.p1_mem = ""
    st.session_state.p1_cat = ""
    st.session_state.close_dialog = True

@st.dialog("프로젝트 만들기", width="normal")
def popup_step1():
    if st.session_state.close_dialog:
        st.session_state.close_dialog = False
        st.rerun()

    st.write("프로젝트를 생성하고 굵직한 **대분류(뼈대)**를 먼저 만들어 둡니다.")
    c1, c2 = st.columns([6, 4])
    p_name = c1.text_input("프로젝트명", key="p1_proj", placeholder="예: 얼음정수기 V2 개발")
    c2.text_input("참여 인원", key="p1_mem", placeholder="예: PM 홍길동, 아무개")

    st.markdown("<hr style='margin: 10px 0;'>", unsafe_allow_html=True)
    st.text_input("추가할 대분류 이름", key="p1_cat", placeholder="예: 1. 기획 / 2. 부품개발 (※ 비워두면 프로젝트만 생성됩니다)")

    c_btn1, c_btn2, c_btn3 = st.columns(3)
    c_btn1.button("연속 추가", type="primary", use_container_width=True, on_click=cb_save_p1_continuous, icon=":material/add:")
    c_btn2.button("저장 후 닫기", type="primary", use_container_width=True, on_click=cb_save_p1_and_close, icon=":material/check:")
    c_btn3.button("닫기", use_container_width=True, on_click=cb_close_p1)

    if p_name and p_name in st.session_state.categories and st.session_state.categories[p_name]:
        if st.session_state.msg_p1:
            if "⚠️" in st.session_state.msg_p1: st.warning(st.session_state.msg_p1)
            else: st.success(st.session_state.msg_p1)

        st.write(f"**[{p_name}]에 생성된 대분류 목록:**")
        for c in st.session_state.categories[p_name]: st.markdown(f"- {c}")
    elif st.session_state.msg_p1:
        if "⚠️" in st.session_state.msg_p1: st.warning(st.session_state.msg_p1)
        else: st.success(st.session_state.msg_p1)

def update_project_info(old_name, new_name, new_mem):
    new_name = new_name.strip()
    new_mem = new_mem.strip()
    if not new_name: return False, "프로젝트명을 입력해주세요."

    if old_name == new_name:
        if st.session_state.p_members.get(old_name) != new_mem:
            st.session_state.p_members[old_name] = new_mem
            save_data(st.session_state.projects, st.session_state.p_members, st.session_state.categories)
        return True, ""

    if new_name in st.session_state.projects: return False, "이미 존재하는 프로젝트명입니다."

    try:
        gc = get_gspread_client()
        sh = gc.open_by_url(SHEET_URL)
        ws = sh.worksheet(old_name[:31])
        ws.update_title(new_name[:31])
    except Exception as e: pass

    st.session_state.projects[new_name] = st.session_state.projects.pop(old_name)
    st.session_state.categories[new_name] = st.session_state.categories.pop(old_name)
    st.session_state.p_members.pop(old_name, None)
    st.session_state.p_members[new_name] = new_mem

    if st.session_state.get('current_page') == old_name:
        st.session_state.current_page = new_name

    save_data(st.session_state.projects, st.session_state.p_members, st.session_state.categories)
    return True, ""

def save_p2(p_sel, cat_sel):
    sub = st.session_state.p2_sub.strip()
    man = st.session_state.p2_man.strip()
    dep = st.session_state.p2_dep.strip()
    doc = st.session_state.p2_doc.strip()
    s_date = st.session_state.p2_start
    e_date = st.session_state.p2_end

    if p_sel not in st.session_state.projects: st.session_state.projects[p_sel] = []
    if p_sel not in st.session_state.categories: st.session_state.categories[p_sel] = []

    st.session_state.projects[p_sel].append({
        "대분류": cat_sel, "소분류": sub, "담당자": man, "부서": dep, "문서": doc,
        "예상시작일": str(s_date), "예상종료일": str(e_date), "실제시작일": None, "실제종료일": None,
        "메모": ""
    })

    if cat_sel not in st.session_state.categories[p_sel]:
        st.session_state.categories[p_sel].append(cat_sel)

    save_data(st.session_state.projects, st.session_state.p_members, st.session_state.categories)

    if sub: st.session_state.msg_p2 = f"[{cat_sel}] 하위에 '{sub}' 추가 완료!"
    else: st.session_state.msg_p2 = f"[{cat_sel}] 대분류 업무 추가 완료!"

    st.session_state.p2_sub = ""
    st.session_state.p2_man = ""
    st.session_state.p2_dep = ""
    st.session_state.p2_doc = ""

def cb_save_continuous(p_sel, cat_sel):
    if cat_sel.strip(): save_p2(p_sel, cat_sel.strip())
    else: st.session_state.msg_p2 = "⚠️ 대분류를 선택하거나 직접 입력해주세요!"

def cb_save_and_close(p_sel, cat_sel):
    if cat_sel.strip(): save_p2(p_sel, cat_sel.strip())
    st.session_state.close_dialog = True

def cb_just_close():
    st.session_state.msg_p2 = ""
    for k in ['p2_sub', 'p2_man', 'p2_dep', 'p2_doc', 'in_new_cat']:
        if k in st.session_state: st.session_state[k] = ""
    st.session_state.close_dialog = True

@st.dialog("세부업무 추가 및 프로젝트 수정", width="large")
def popup_step2(p_sel):
    if st.session_state.close_dialog:
        st.session_state.close_dialog = False
        st.rerun()

    with st.expander("프로젝트 기본 정보 수정 (이름/참여인원)", icon=":material/settings:"):
        st.markdown("<div style='font-size:13px; color:#555; margin-bottom:10px;'>내용 변경 후 아래 버튼을 누르면 구글 시트 탭 이름까지 즉시 변경됩니다.</div>", unsafe_allow_html=True)
        c_e1, c_e2 = st.columns([6, 4])
        new_p_name = c_e1.text_input("프로젝트명 수정", value=p_sel, key=f"edit_pname_{p_sel}")
        new_p_mem = c_e2.text_input("참여인원 수정", value=st.session_state.p_members.get(p_sel, ""), key=f"edit_pmem_{p_sel}")

        if st.button("프로젝트 정보 변경 적용", use_container_width=True):
            success, msg = update_project_info(p_sel, new_p_name, new_p_mem)
            if success:
                st.session_state.close_dialog = True
                st.rerun()
            else: st.error(msg)

    st.write(f"**[{p_sel}]** 프로젝트에 세부 업무를 추가합니다.")

    cat_list = st.session_state.categories.get(p_sel, [])
    options_c = cat_list + ["직접 입력 (새 대분류 추가)"]
    c_sel = st.selectbox("대분류 선택", options_c)

    if c_sel == "직접 입력 (새 대분류 추가)":
        final_cat = st.text_input("새 대분류 직접 입력", placeholder="여기에 새 대분류를 타이핑하세요", label_visibility="collapsed", key="in_new_cat")
    else: final_cat = c_sel

    st.markdown("<hr style='margin: 10px 0;'>", unsafe_allow_html=True)
    st.text_input("구분 (세부 업무명)", key="p2_sub", placeholder="예: 도면 설계 (※ 비워두면 대분류 이름만 저장됩니다)")

    c3, c4, c5 = st.columns(3)
    c3.text_input("담당자", key="p2_man")
    c4.text_input("관련부서", key="p2_dep")
    c5.text_input("관련문서", key="p2_doc")

    c6, c7 = st.columns(2)
    c6.date_input("계획 시작일", key="p2_start")
    c7.date_input("계획 종료일", key="p2_end")

    st.write("")
    c_btn1, c_btn2, c_btn3 = st.columns(3)
    c_btn1.button("연속 저장", type="primary", use_container_width=True, on_click=cb_save_continuous, args=(p_sel, final_cat), icon=":material/add:")
    c_btn2.button("저장 후 닫기", type="primary", use_container_width=True, on_click=cb_save_and_close, args=(p_sel, final_cat), icon=":material/check:")
    c_btn3.button("닫기", use_container_width=True, on_click=cb_just_close)

    if st.session_state.msg_p2:
        if "⚠️" in st.session_state.msg_p2: st.warning(st.session_state.msg_p2)
        else:
            st.success(st.session_state.msg_p2)
            added_tasks = [t['소분류'] for t in st.session_state.projects.get(p_sel, []) if t['대분류'] == final_cat.strip()]
            if added_tasks:
                st.write(f"**[{final_cat.strip()}]에 최근 추가된 업무:**")
                for t in added_tasks[-5:]:
                    disp_t = t if t else "- (대분류만 저장됨)"
                    st.markdown(f"↳ {disp_t}")

@st.dialog("업무 내용 수정", width="large")
def edit_task_popup(p_name, default_cat=None):
    raw_data = st.session_state.projects.get(p_name, [])
    cat_list = st.session_state.categories.get(p_name, [])

    if not cat_list:
        st.warning("수정할 대분류가 없습니다.")
        return

    st.write("수정할 업무를 선택 후 내용을 덮어쓰거나, **대분류 순서를 위아래로 자유롭게 이동**할 수 있습니다.")
    c1, c2 = st.columns(2)

    sel_key = f"edit_sel_{p_name}"
    if sel_key not in st.session_state or st.session_state[sel_key] not in cat_list:
        st.session_state[sel_key] = default_cat if default_cat in cat_list else cat_list[0]

    current_sel = st.session_state[sel_key]

    with c1:
        ui_top = st.container()
        ui_bottom = st.container()

        with ui_bottom:
            cc1, cc2 = st.columns(2)
            up_clicked = cc1.button("위로 이동", use_container_width=True, icon=":material/arrow_upward:")
            down_clicked = cc2.button("아래로 이동", use_container_width=True, icon=":material/arrow_downward:")

            if up_clicked:
                idx = st.session_state.categories[p_name].index(current_sel)
                if idx > 0:
                    st.session_state.categories[p_name][idx], st.session_state.categories[p_name][idx-1] = st.session_state.categories[p_name][idx-1], st.session_state.categories[p_name][idx]
                    save_data(st.session_state.projects, st.session_state.p_members, st.session_state.categories)

            if down_clicked:
                idx = st.session_state.categories[p_name].index(current_sel)
                if idx < len(st.session_state.categories[p_name]) - 1:
                    st.session_state.categories[p_name][idx], st.session_state.categories[p_name][idx+1] = st.session_state.categories[p_name][idx+1], st.session_state.categories[p_name][idx]
                    save_data(st.session_state.projects, st.session_state.p_members, st.session_state.categories)

        with ui_top:
            updated_cats = st.session_state.categories[p_name]
            cat_sel = st.selectbox("1. 대분류 선택 (이동할 대분류)", updated_cats, key=sel_key)

            styled_cats = []
            for c in updated_cats:
                if c == cat_sel: styled_cats.append(f"<span style='color:#0E7490; font-weight:bold; border-bottom:2px solid #0E7490;'>{c}</span>")
                else: styled_cats.append(c)
            order_str = " ➔ ".join(styled_cats)

            st.markdown(f"""
            <div style='font-size:14px; background:#EEF3F6; padding:10px; border-radius:8px; margin-top:5px; margin-bottom:10px; border:1px solid #D5E0E7;'>
                <b>현재 순서</b><br>{order_str}
            </div>
            """, unsafe_allow_html=True)

    task_options_with_idx = [(i, d['소분류'] if d['소분류'] else "- (세부업무명 없음)") for i, d in enumerate(raw_data) if d['대분류'] == cat_sel]

    with c2:
        if not task_options_with_idx:
            st.info("이 대분류에는 아직 세부 업무가 없습니다. 대분류 이름만 수정할 수 있습니다.")
            with st.form(key=f"edit_empty_cat_{cat_sel}"):
                new_cat_only = st.text_input("빈 대분류 폴더명 수정", value=cat_sel)
                if st.form_submit_button("이름 덮어쓰기", type="primary"):
                    if new_cat_only != cat_sel:
                        if p_name not in st.session_state.categories: st.session_state.categories[p_name] = []
                        idx = st.session_state.categories[p_name].index(cat_sel)
                        st.session_state.categories[p_name][idx] = new_cat_only
                        st.session_state[sel_key] = new_cat_only
                        save_data(st.session_state.projects, st.session_state.p_members, st.session_state.categories)
                        st.success("대분류 이름이 변경되었습니다!")
                        st.rerun()
            return

        format_func = lambda x: f"[{x[0]}] {x[1]}"
        selected_tuple = st.selectbox("2. 세부 업무 선택", task_options_with_idx, format_func=format_func)

    target_idx = selected_tuple[0]
    target_data = raw_data[target_idx]

    st.markdown("<hr style='margin: 10px 0;'>", unsafe_allow_html=True)

    with st.form(key=f"edit_form_{target_idx}"):
        f1, f2 = st.columns(2)
        new_cat = f1.text_input("대분류 덮어쓰기 (해당 그룹 일괄 변경)", value=cat_sel)
        new_sub = f2.text_input("세부업무 덮어쓰기", value=target_data['소분류'])

        f3, f4, f5 = st.columns(3)
        new_man = f3.text_input("담당자 수정", value=target_data.get('담당자', ''))
        new_dep = f4.text_input("관련부서 수정", value=target_data.get('부서', ''))
        new_doc = f5.text_input("관련문서 수정", value=target_data.get('문서', ''))

        f6, f7 = st.columns(2)
        try:
            est_s = pd.to_datetime(target_data['예상시작일']).date()
            est_e = pd.to_datetime(target_data['예상종료일']).date()
        except:
            est_s = datetime.date.today()
            est_e = datetime.date.today()

        new_est_s = f6.date_input("계획 시작일 수정", value=est_s)
        new_est_e = f7.date_input("계획 종료일 수정", value=est_e)

        if st.form_submit_button("수정 내용 완벽히 덮어쓰기", type="primary", use_container_width=True):
            if new_cat != cat_sel:
                if p_name not in st.session_state.categories: st.session_state.categories[p_name] = []
                if cat_sel in st.session_state.categories[p_name]:
                    cat_idx = st.session_state.categories[p_name].index(cat_sel)
                    st.session_state.categories[p_name][cat_idx] = new_cat
                for d in st.session_state.projects[p_name]:
                    if d['대분류'] == cat_sel: d['대분류'] = new_cat
                st.session_state[sel_key] = new_cat

            target_data['소분류'] = new_sub
            target_data['담당자'] = new_man
            target_data['부서'] = new_dep
            target_data['문서'] = new_doc
            target_data['예상시작일'] = str(new_est_s)
            target_data['예상종료일'] = str(new_est_e)

            save_data(st.session_state.projects, st.session_state.p_members, st.session_state.categories)
            st.success("업무 내용이 깔끔하게 덮어쓰기 되었습니다!")
            st.rerun()

def cb_update_progress(p_name, target_idx, sk, ek, memo_val):
    st.session_state.projects[p_name][target_idx]['실제시작일'] = str(st.session_state[sk])
    st.session_state.projects[p_name][target_idx]['실제종료일'] = str(st.session_state[ek])
    st.session_state.projects[p_name][target_idx]['메모'] = memo_val
    del st.session_state[sk]
    del st.session_state[ek]
    save_data(st.session_state.projects, st.session_state.p_members, st.session_state.categories)
    st.session_state.close_dialog = True

@st.dialog("진행 현황 입력")
def update_progress_popup(p_name, target_idx):
    if st.session_state.close_dialog:
        st.session_state.close_dialog = False
        st.rerun()

    target_data = st.session_state.projects[p_name][target_idx]
    disp_sub = target_data['소분류'] if target_data['소분류'] else "-"
    st.write(f"**[{target_data['대분류']}] {disp_sub}**")
    st.info(f"당초 계획 일정: {target_data['예상시작일']} ~ {target_data['예상종료일']}")

    memo_val = st.text_input("진행상태 비고/메모", value=target_data.get('메모', ''), placeholder="이슈나 특이사항을 자유롭게 적어주세요")
    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

    today = datetime.date.today()
    def parse_date(d_str):
        if d_str: return pd.to_datetime(d_str).date()
        return today

    sk = f"act_s_{target_idx}"
    ek = f"act_e_{target_idx}"

    if sk not in st.session_state: st.session_state[sk] = parse_date(target_data.get('실제시작일'))
    if ek not in st.session_state: st.session_state[ek] = parse_date(target_data.get('실제종료일'))

    c6, c7 = st.columns(2)
    c6.date_input("실제 시작일정", key=sk)
    c7.date_input("현재(완료) 일정", key=ek)

    st.write("")
    st.button("저장 및 반영", type="primary", use_container_width=True, on_click=cb_update_progress, args=(p_name, target_idx, sk, ek, memo_val), icon=":material/check:")

LEGEND_HTML = (
    "<div class='legend'>"
    "<span><i style='background:#B8C7D2;'></i>계획 일정</span>"
    "<span><i style='background:#0E7490;'></i>실적(진행) 일정</span>"
    "<span><b></b>오늘</span>"
    "</div>"
)

@st.dialog("프로젝트 진행 요약", width="large")
def summary_popup(p_name):
    st.markdown(f"### [{p_name}] 대분류 일정 요약")
    st.markdown(LEGEND_HTML, unsafe_allow_html=True)

    raw_data = st.session_state.projects.get(p_name, [])
    if not raw_data:
        st.info("아직 등록된 업무가 없습니다.")
        return

    min_date = pd.to_datetime(min([d['예상시작일'] for d in raw_data]))
    max_date = pd.to_datetime(max([d['예상종료일'] for d in raw_data]))
    for d in raw_data:
        if d.get('실제시작일'): min_date = min(min_date, pd.to_datetime(d['실제시작일']))
        if d.get('실제종료일'): max_date = max(max_date, pd.to_datetime(d['실제종료일']))

    timeline_html, bg_grid_html, full_min_date, full_max_date, total_days = get_gantt_assets(min_date, max_date)

    # 헤더와 막대 칸이 같은 열 비율을 쓰도록 한 줄에 배치
    col_h1, col_h2 = st.columns([2.5, 7.5], vertical_alignment="bottom")
    col_h1.markdown("<div class='header-cell notranslate'>대분류</div>", unsafe_allow_html=True)
    col_h2.markdown(f"<div style='width:100%;'>{timeline_html}</div>", unsafe_allow_html=True)

    categories = st.session_state.categories.get(p_name, [])
    for cat in categories:
        cat_items = [d for d in raw_data if d['대분류'] == cat]
        if not cat_items: continue

        c_est_starts = [pd.to_datetime(d['예상시작일']) for d in cat_items if d.get('예상시작일')]
        c_est_ends = [pd.to_datetime(d['예상종료일']) for d in cat_items if d.get('예상종료일')]
        c_act_starts = [pd.to_datetime(d['실제시작일']) for d in cat_items if d.get('실제시작일')]
        c_act_ends = [pd.to_datetime(d['실제종료일']) for d in cat_items if d.get('실제종료일')]

        c_est_s = min(c_est_starts) if c_est_starts else None
        c_est_e = max(c_est_ends) if c_est_ends else None
        c_act_s = min(c_act_starts) if c_act_starts else None
        c_act_e = max(c_act_ends) if c_act_ends else None

        col_cat, col_chart = st.columns([2.5, 7.5], vertical_alignment="center")

        with col_cat:
            st.markdown(f"<div class='notranslate' style='padding: 0 10px; background-color: #E6EDF2; font-weight: 700; color: #0F2A3D; font-size: 14px; border-radius: 6px; height: 50px; display: flex; align-items: center; justify-content: center; text-align: center; line-height: 1.2; word-break: keep-all;' title='{html.escape(cat)}'>{html.escape(cat)}</div>", unsafe_allow_html=True)

        with col_chart:
            cat_bar_html = f"<div class='notranslate' style='position:relative; width:100%; height:50px; background-color:#fff; border:1px solid #D5E0E7; box-sizing:border-box; border-radius:6px; overflow:hidden;'>{bg_grid_html}"

            if c_est_s and c_est_e:
                e_l = max(0.0, min(100.0, (c_est_s - full_min_date).days / total_days * 100))
                e_r = max(0.0, min(100.0, ((c_est_e - full_min_date).days + 1) / total_days * 100))
                e_w = max(0.5, e_r - e_l)
                cat_bar_html += f"<div style='position:absolute; top:12px; height:10px; background-color:{C_CAT_PLAN}; border-radius:4px; left:{e_l}%; width:{e_w}%; z-index:1;' title='계획: {c_est_s.date()} ~ {c_est_e.date()}'></div>"

            if c_act_s and c_act_e:
                a_l = max(0.0, min(100.0, (c_act_s - full_min_date).days / total_days * 100))
                a_r = max(0.0, min(100.0, ((c_act_e - full_min_date).days + 1) / total_days * 100))
                a_w = max(0.5, a_r - a_l)
                cat_bar_html += f"<div style='position:absolute; top:28px; height:10px; background-color:{C_CAT_ACT}; border-radius:4px; left:{a_l}%; width:{a_w}%; z-index:2;' title='실적: {c_act_s.date()} ~ {c_act_e.date()}'></div>"

            cat_bar_html += "</div>"
            st.markdown(cat_bar_html, unsafe_allow_html=True)

        st.markdown("<div style='height: 5px;'></div>", unsafe_allow_html=True)

    st.write("")
    if st.button("닫기", use_container_width=True):
        st.rerun()

# ==========================================
# 3. 메인 화면 (대시보드 ↔ 상세페이지)
# ==========================================
if st.session_state.current_page == 'Dashboard':

    c_head1, c_head2, c_head3, c_head4 = st.columns([5.0, 1.6, 1.6, 2.0], vertical_alignment="center")

    with c_head1:
        st.markdown(
            "<div class='dash-title notranslate'>중점추진과제 WBS</div>"
            "<div class='dash-sub notranslate'>프로젝트별 계획 대비 실적을 한눈에 관리합니다</div>",
            unsafe_allow_html=True)

    with c_head2:
        def make_integrated_excel():
            output = BytesIO()
            wb = Workbook()
            if "Sheet" in wb.sheetnames:
                wb.remove(wb["Sheet"])

            if not st.session_state.projects:
                ws = wb.create_sheet(title="프로젝트 없음")
                ws.append(["등록된 프로젝트가 없습니다."])
            else:
                for p_name, project_data in st.session_state.projects.items():
                    safe_p_name = p_name[:31]
                    ws = wb.create_sheet(title=safe_p_name)
                    headers = ["프로젝트명", "참여인원", "대분류", "세부업무명", "담당자", "관련부서", "관련문서", "계획시작", "계획종료", "실제시작", "실제종료", "비고(메모)"]
                    ws.append(headers)

                    header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                    for col in range(1, 13):
                        ws.cell(row=1, column=col).fill = header_fill
                        ws.cell(row=1, column=col).font = Font(bold=True)

                    p_mem = st.session_state.p_members.get(p_name, '')
                    for row in project_data:
                        a_start = str(row['실제시작일']) if row['실제시작일'] else ""
                        a_end = str(row['실제종료일']) if row['실제종료일'] else ""
                        memo = row.get('메모', '')
                        ws.append([p_name, p_mem, row['대분류'], row['소분류'], row['담당자'], row['부서'], row['문서'], str(row['예상시작일']), str(row['예상종료일']), a_start, a_end, memo])
            wb.save(output)
            return output.getvalue()

        st.download_button(
            label="엑셀 저장",
            data=make_integrated_excel(),
            file_name=f"WBS_전체일정_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            icon=":material/download:"
        )

    with c_head3:
        st.link_button("구글시트", url=SHEET_URL, use_container_width=True, icon=":material/table_view:")

    with c_head4:
        if st.button("새 프로젝트", type="primary", use_container_width=True, icon=":material/add:"):
            popup_step1()

    project_names = list(st.session_state.projects.keys())

    # ----- 요약 카드(KPI) -----
    st.markdown(
        "<div class='kpi-row notranslate'>"
        f"<div class='kpi'><div class='kpi-label'>진행 중인 프로젝트</div><div class='kpi-value'>{len(project_names)}<small>개</small></div></div>"
        "</div>",
        unsafe_allow_html=True)

    if not project_names:
        st.info("아직 생성된 프로젝트가 없습니다. 오른쪽 위 '새 프로젝트' 버튼으로 시작하세요.")
    else:
        cols = st.columns(3)

        for idx, p in enumerate(project_names):
            col_target = cols[idx % 3]

            with col_target:
                with st.container(border=True):
                    raw_data = st.session_state.projects[p]
                    p_mem = st.session_state.p_members.get(p, '')
                    tot = len(raw_data)
                    done_cnt = sum(1 for d in raw_data if d.get('실제시작일') and d.get('실제종료일'))
                    pct = int(done_cnt / tot * 100) if tot else 0

                    st.markdown(
                        f"<div class='pc-title notranslate'>{html.escape(p)}</div>"
                        f"<div class='pc-meta notranslate'><span>{ICON_USER}{html.escape(p_mem) if p_mem else '-'}</span><span>{ICON_TASK}업무 {tot}건</span></div>",
                        unsafe_allow_html=True)

                    if tot > 0:
                        min_date = pd.to_datetime(min([d['예상시작일'] for d in raw_data]))
                        max_date = pd.to_datetime(max([d['예상종료일'] for d in raw_data]))
                        for d in raw_data:
                            if d.get('실제시작일'): min_date = min(min_date, pd.to_datetime(d['실제시작일']))
                            if d.get('실제종료일'): max_date = max(max_date, pd.to_datetime(d['실제종료일']))

                        timeline_html, bg_grid_html, full_min_date, full_max_date, total_days = get_gantt_assets(min_date, max_date)

                        p_est_start = min([pd.to_datetime(d['예상시작일']) for d in raw_data])
                        p_est_end = max([pd.to_datetime(d['예상종료일']) for d in raw_data])

                        e_left = max(0.0, min(100.0, (p_est_start - full_min_date).days / total_days * 100))
                        e_right = max(0.0, min(100.0, ((p_est_end - full_min_date).days + 1) / total_days * 100))
                        e_width = max(0.5, e_right - e_left)

                        chart_html = (
                            f"<div class='pc-range notranslate'>{ICON_CAL}{full_min_date.strftime('%y.%m.%d')} ~ {full_max_date.strftime('%y.%m.%d')}</div>"
                            f"<div class='notranslate' style='position:relative; width:100%; height:26px; background-color:#fff; border-radius:6px; border:1px solid #D5E0E7; box-sizing:border-box; overflow:hidden;'>"
                            f"{bg_grid_html}"
                            f"<div style='position:absolute; top:5px; height:7px; background-color:{C_PLAN}; border-radius:3px; left:{e_left}%; width:{e_width}%; z-index:1;'></div>"
                        )

                        actual_starts = [pd.to_datetime(d['실제시작일']) for d in raw_data if d.get('실제시작일')]
                        actual_ends = [pd.to_datetime(d['실제종료일']) for d in raw_data if d.get('실제종료일')]

                        if actual_starts and actual_ends:
                            p_act_start = min(actual_starts)
                            p_act_end = max(actual_ends)

                            a_left = max(0.0, min(100.0, (p_act_start - full_min_date).days / total_days * 100))
                            a_right = max(0.0, min(100.0, ((p_act_end - full_min_date).days + 1) / total_days * 100))
                            a_width = max(0.5, a_right - a_left)

                            chart_html += f"<div style='position:absolute; top:14px; height:7px; background-color:{C_ACT}; border-radius:3px; left:{a_left}%; width:{a_width}%; z-index:2;'></div>"

                        chart_html += "</div><div style='height:14px;'></div>"

                        st.markdown(chart_html, unsafe_allow_html=True)
                    else:
                        st.markdown("<div style='font-size:12.5px; color:#8A9BA8; margin:14px 0 16px 0;'>등록된 일정이 없습니다. '업무추가'로 시작하세요.</div>", unsafe_allow_html=True)

                    # 카드 전체를 덮는 투명 버튼 (CSS로 카드 크기만큼 늘림) → 카드 안쪽을 누르면 입장
                    if st.button("열기", key=f"go_{p}"):
                        st.session_state.current_page = p
                        st.rerun()

                    b2, b3, b4 = st.columns([1.5, 1.1, 1.1])
                    with b2:
                        if st.button("업무추가", key=f"add_{p}", use_container_width=True, icon=":material/add:"):
                            popup_step2(p)
                    with b3:
                        if st.button("요약", key=f"sum_{p}", use_container_width=True, icon=":material/bar_chart:"):
                            summary_popup(p)
                    with b4:
                        if st.button("삭제", key=f"del_{p}", use_container_width=True, icon=":material/delete:"):
                            delete_popup(p)

else:
    # ----------------------------------------
    # 개별 프로젝트 상세 페이지
    # ----------------------------------------
    p_name = st.session_state.current_page

    col_t1, col_t2, col_t3 = st.columns([7.0, 1.5, 1.8], vertical_alignment="center")

    with col_t1:
        st.markdown(
            f"<div class='page-title notranslate'>{html.escape(p_name)} WBS</div>"
            f"<div class='page-sub notranslate'><b>담당자/참여인원</b> &nbsp;{html.escape(st.session_state.p_members.get(p_name, ''))}</div>",
            unsafe_allow_html=True)

    with col_t2:
        if st.button("대시보드", use_container_width=True, icon=":material/arrow_back:"):
            st.session_state.current_page = 'Dashboard'
            st.rerun()
    with col_t3:
        if st.button("전체 업무 수정", use_container_width=True, icon=":material/edit:"):
            if p_name in st.session_state.categories and st.session_state.categories[p_name]:
                if f"edit_sel_{p_name}" not in st.session_state:
                    st.session_state[f"edit_sel_{p_name}"] = st.session_state.categories[p_name][0]
            edit_task_popup(p_name)

    raw_data = st.session_state.projects.get(p_name, [])

    if len(raw_data) > 0:
        min_date = pd.to_datetime(min([d['예상시작일'] for d in raw_data]))
        max_date = pd.to_datetime(max([d['예상종료일'] for d in raw_data]))

        for d in raw_data:
            if d.get('실제시작일'): min_date = min(min_date, pd.to_datetime(d['실제시작일']))
            if d.get('실제종료일'): max_date = max(max_date, pd.to_datetime(d['실제종료일']))

        timeline_html, bg_grid_html, full_min_date, full_max_date, total_days = get_gantt_assets(min_date, max_date)

        st.markdown(LEGEND_HTML, unsafe_allow_html=True)

        COL_RATIO = [1.5, 1, 1, 1, 1, 5.5]

        # 헤더: 아래 행들과 '완전히 같은 열 비율'을 사용 (스크롤 박스를 없애 폭 어긋남 제거)
        h1, h2, h3, h4, h5, h6 = st.columns(COL_RATIO, vertical_alignment="bottom")

        h1.markdown("<div class='header-cell notranslate gantt-head-marker'>구분</div>", unsafe_allow_html=True)
        h2.markdown("<div class='header-cell notranslate'>담당자</div>", unsafe_allow_html=True)
        h3.markdown("<div class='header-cell notranslate'>관련부서</div>", unsafe_allow_html=True)
        h4.markdown("<div class='header-cell notranslate'>관련문서</div>", unsafe_allow_html=True)
        h5.markdown("<div class='header-cell notranslate'>진행 현황</div>", unsafe_allow_html=True)
        h6.markdown(f"<div style='width:100%;'>{timeline_html}</div>", unsafe_allow_html=True)

        categories = st.session_state.categories.get(p_name, [])
        for cat in categories:
            cat_items_with_idx = [(i, d) for i, d in enumerate(raw_data) if d['대분류'] == cat]
            if not cat_items_with_idx:
                continue

            st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

            cat_c1, cat_c2, cat_c3, cat_c4, cat_c5, cat_c6 = st.columns(COL_RATIO, vertical_alignment="center")

            with cat_c1:
                st.markdown(f"""<div style='position: relative; width: 100%; height: 45px;'><div class='notranslate' style='position: absolute; top: 0; left: 0; width: 285%; padding: 0 15px; background-color: #E6EDF2; font-weight: 700; color: #0F2A3D; font-size: 15px; border-radius: 6px; height: 45px; display: flex; align-items: center; white-space: nowrap; z-index: 10; box-sizing: border-box;'>{html.escape(cat)}</div></div>""", unsafe_allow_html=True)

            with cat_c5:
                if st.button("수정", key=f"edit_cat_{p_name}_{cat}", use_container_width=True, icon=":material/edit:"):
                    st.session_state[f"edit_sel_{p_name}"] = cat
                    edit_task_popup(p_name)

            with cat_c6:
                c_est_s, c_est_e = None, None
                c_act_s, c_act_e = None, None

                cat_est_starts = [pd.to_datetime(d['예상시작일']) for i, d in cat_items_with_idx if d.get('예상시작일')]
                cat_est_ends = [pd.to_datetime(d['예상종료일']) for i, d in cat_items_with_idx if d.get('예상종료일')]
                cat_act_starts = [pd.to_datetime(d['실제시작일']) for i, d in cat_items_with_idx if d.get('실제시작일')]
                cat_act_ends = [pd.to_datetime(d['실제종료일']) for i, d in cat_items_with_idx if d.get('실제종료일')]

                if cat_est_starts and cat_est_ends:
                    c_est_s = min(cat_est_starts)
                    c_est_e = max(cat_est_ends)
                if cat_act_starts and cat_act_ends:
                    c_act_s = min(cat_act_starts)
                    c_act_e = max(cat_act_ends)

                cat_bar_html = f"<div class='bar-container notranslate' style='background-color: #F7FAFB; overflow: hidden !important;'>{bg_grid_html}"

                if c_est_s and c_est_e:
                    e_l = max(0.0, min(100.0, (c_est_s - full_min_date).days / total_days * 100))
                    e_r = max(0.0, min(100.0, ((c_est_e - full_min_date).days + 1) / total_days * 100))
                    e_w = max(0.5, e_r - e_l)
                    cat_bar_html += f"<div style='position:absolute; top:12px; height:8px; background-color:{C_CAT_PLAN}; border-radius:4px; left:{e_l}%; width:{e_w}%; z-index:1;' title='대분류 전체 계획: {c_est_s.date()} ~ {c_est_e.date()}'></div>"

                if c_act_s and c_act_e:
                    a_l = max(0.0, min(100.0, (c_act_s - full_min_date).days / total_days * 100))
                    a_r = max(0.0, min(100.0, ((c_act_e - full_min_date).days + 1) / total_days * 100))
                    a_w = max(0.5, a_r - a_l)
                    cat_bar_html += f"<div style='position:absolute; top:25px; height:8px; background-color:{C_CAT_ACT}; border-radius:4px; left:{a_l}%; width:{a_w}%; z-index:2;' title='대분류 전체 실적: {c_act_s.date()} ~ {c_act_e.date()}'></div>"

                cat_bar_html += "</div>"
                st.markdown(cat_bar_html, unsafe_allow_html=True)

            for actual_idx, row in cat_items_with_idx:
                man = html.escape(row['담당자']) if row['담당자'] else "-"
                dep = html.escape(row['부서']) if row['부서'] else "-"
                doc = html.escape(row['문서']) if row['문서'] else "-"

                disp_sub = html.escape(row['소분류']) if row['소분류'] else "-"

                c1, c2, c3, c4, c5, c6 = st.columns(COL_RATIO, vertical_alignment="center")

                c1.markdown(f"<div class='item-cell-left notranslate'>&nbsp;&nbsp;&nbsp;&nbsp;↳ {disp_sub}</div>", unsafe_allow_html=True)
                c2.markdown(f"<div class='item-cell notranslate'>{man}</div>", unsafe_allow_html=True)
                c3.markdown(f"<div class='item-cell notranslate'>{dep}</div>", unsafe_allow_html=True)
                c4.markdown(f"<div class='item-cell notranslate'>{doc}</div>", unsafe_allow_html=True)

                with c5:
                    if st.button("진행입력", key=f"upd_btn_{p_name}_{actual_idx}", use_container_width=True, icon=":material/edit_calendar:"):
                        update_progress_popup(p_name, actual_idx)

                est_start = pd.to_datetime(row['예상시작일'])
                est_end = pd.to_datetime(row['예상종료일'])

                e_left = max(0.0, min(100.0, (est_start - full_min_date).days / total_days * 100))
                e_right = max(0.0, min(100.0, ((est_end - full_min_date).days + 1) / total_days * 100))
                e_width = max(0.5, e_right - e_left)

                bar_html = f"<div class='bar-container notranslate'>{bg_grid_html}<div class='bar-est' style='left: {e_left}%; width: {e_width}%;' title='계획: {est_start.date()} ~ {est_end.date()}'></div>"

                memo_left = e_left

                if row.get('실제시작일') and row.get('실제종료일'):
                    act_start = pd.to_datetime(row['실제시작일'])
                    act_end = pd.to_datetime(row['실제종료일'])

                    a_left = max(0.0, min(100.0, (act_start - full_min_date).days / total_days * 100))
                    a_right = max(0.0, min(100.0, ((act_end - full_min_date).days + 1) / total_days * 100))
                    a_width = max(0.5, a_right - a_left)

                    memo_left = a_left
                    bar_html += f"<div class='bar-act' style='left: {a_left}%; width: {a_width}%;' title='실적: {act_start.date()} ~ {act_end.date()}'></div>"

                memo_val = (row.get('메모', '') or '').strip()
                if memo_val:
                    bar_html += f"<div style='position:absolute; top:23px; left:calc({memo_left}% + 5px); font-size:12px; font-weight:bold; color:#000; z-index:10; white-space:nowrap; text-shadow: -1px -1px 0 #fff, 1px -1px 0 #fff, -1px 1px 0 #fff, 1px 1px 0 #fff;'>{html.escape(memo_val)}</div>"

                bar_html += "</div>"
                c6.markdown(bar_html, unsafe_allow_html=True)

                st.markdown("<hr class='row-line'>", unsafe_allow_html=True)

        st.divider()
    else:
        st.info("아직 등록된 업무가 없습니다. 대시보드로 돌아가 '업무추가'로 세부 업무를 등록하세요.")
