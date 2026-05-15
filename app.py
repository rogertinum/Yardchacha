# 공용차량 예약 및 주행거리 정산 시스템
import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime, date, timedelta
from io import BytesIO
import calendar as cal_module
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="공용차량 관리",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ────────────────────────────────────────────────────────────────
# 상수
# ────────────────────────────────────────────────────────────────
DEPARTMENTS = [
    "생산기술팀", "공법기술", "의장기술", "운반종합기술",
    "시공기술", "족장기술", "DFX그룹", "도장기술",
]
ADMIN_PASSWORD = "1111"
DB_PATH = "vehicle_management.db"
VEHICLE_NAME = "EV3"
VEHICLE_NUMBER = "05하 7211"

# ────────────────────────────────────────────────────────────────
# 쿠키 매니저 (옵션 의존성)
# ────────────────────────────────────────────────────────────────
COOKIES_ENABLED = False
_cm = None
try:
    import extra_streamlit_components as stx

    @st.cache_resource
    def _get_cm():
        return stx.CookieManager()

    _cm = _get_cm()
    COOKIES_ENABLED = True
except Exception:
    pass


def cookie_get(key, default=None):
    if not COOKIES_ENABLED or _cm is None:
        return default
    try:
        v = _cm.get(key)
        return v if v is not None else default
    except Exception:
        return default


def cookie_set(key, value):
    if not COOKIES_ENABLED or _cm is None:
        return
    try:
        _cm.set(key, str(value), expires_at=datetime(2028, 1, 1))
    except Exception:
        pass


# ────────────────────────────────────────────────────────────────
# 데이터베이스
# ────────────────────────────────────────────────────────────────
def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            phone       TEXT PRIMARY KEY,
            employee_id TEXT NOT NULL,
            department  TEXT NOT NULL,
            name        TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS reservations (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            user_phone  TEXT NOT NULL,
            department  TEXT NOT NULL,
            name        TEXT NOT NULL,
            phone       TEXT NOT NULL,
            res_date    TEXT NOT NULL,
            res_time    TEXT NOT NULL,
            destination TEXT NOT NULL,
            created_at  TEXT DEFAULT (datetime('now','localtime'))
        );
        CREATE TABLE IF NOT EXISTS driving_logs (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            user_phone       TEXT NOT NULL,
            department       TEXT NOT NULL,
            name             TEXT NOT NULL,
            phone            TEXT NOT NULL,
            drive_date       TEXT NOT NULL,
            odometer_start   REAL,
            odometer_end     REAL,
            companions       TEXT,
            destination      TEXT,
            charging         INTEGER DEFAULT 0,
            charging_amount  REAL    DEFAULT 0,
            parking_location TEXT,
            status           TEXT    DEFAULT 'pre',
            created_at       TEXT DEFAULT (datetime('now','localtime'))
        );
        """)


def _query(sql, params=(), one=False):
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(sql, params).fetchall()
        if one:
            return dict(rows[0]) if rows else None
        return [dict(r) for r in rows]


def _exec(sql, params=()):
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(sql, params)
        conn.commit()


# ── 사용자 ──
def get_user(phone):
    return _query("SELECT * FROM users WHERE phone=?", (phone,), one=True)


def register_user(phone, emp_id, dept, name):
    _exec(
        "INSERT OR REPLACE INTO users VALUES (?,?,?,?)",
        (phone, emp_id, dept, name),
    )


def auth_user(phone, emp_id):
    return _query(
        "SELECT 1 FROM users WHERE phone=? AND employee_id=?",
        (phone, emp_id),
        one=True,
    ) is not None


# ── 예약 ──
def get_all_reservations():
    return _query("SELECT * FROM reservations ORDER BY res_date, res_time")


def get_reservations_by_date(d):
    return _query(
        "SELECT * FROM reservations WHERE res_date=? ORDER BY res_time", (d,)
    )


def add_reservation(user_phone, dept, name, phone, res_date, res_time, dest):
    _exec(
        "INSERT INTO reservations "
        "(user_phone,department,name,phone,res_date,res_time,destination) "
        "VALUES (?,?,?,?,?,?,?)",
        (user_phone, dept, name, phone, res_date, res_time, dest),
    )


def delete_reservation(rid):
    _exec("DELETE FROM reservations WHERE id=?", (rid,))


# ── 운행 기록 ──
def get_pre_drives(user_phone):
    return _query(
        "SELECT * FROM driving_logs WHERE user_phone=? AND status='pre' ORDER BY drive_date DESC",
        (user_phone,),
    )


def add_pre_drive(user_phone, dept, name, phone, drive_date, odo_start, companions, dest):
    _exec(
        "INSERT INTO driving_logs "
        "(user_phone,department,name,phone,drive_date,odometer_start,companions,destination) "
        "VALUES (?,?,?,?,?,?,?,?)",
        (user_phone, dept, name, phone, drive_date, odo_start, companions, dest),
    )


def complete_drive(lid, odo_end, charging, charge_amt, parking, companions, drive_date):
    _exec(
        """UPDATE driving_logs
           SET odometer_end=?,charging=?,charging_amount=?,parking_location=?,
               companions=?,drive_date=?,status='complete'
           WHERE id=?""",
        (odo_end, int(charging), charge_amt, parking, companions, drive_date, lid),
    )


def get_logs_by_period(start, end):
    return _query(
        "SELECT * FROM driving_logs WHERE drive_date BETWEEN ? AND ? AND status='complete' "
        "ORDER BY drive_date, created_at",
        (start, end),
    )


# ────────────────────────────────────────────────────────────────
# 엑셀 내보내기
# ────────────────────────────────────────────────────────────────
def make_excel(logs, start_date, end_date):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "운행기록부"
    ws.sheet_view.showGridLines = False

    thin   = Side(style="thin")
    medium = Side(style="medium")
    B_ALL  = Border(left=thin, right=thin, top=thin, bottom=thin)
    B_MED  = Border(left=medium, right=medium, top=medium, bottom=medium)
    CA     = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LA     = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    BF     = Font(bold=True, name="맑은 고딕", size=10)
    NF     = Font(name="맑은 고딕", size=10)
    H_FILL = PatternFill("solid", fgColor="BDD7EE")
    G_FILL = PatternFill("solid", fgColor="D9D9D9")

    def mc(r1, c1, r2, c2, val="", font=BF, align=CA, border=B_ALL, fill=None):
        """merge + set"""
        if r1 != r2 or c1 != c2:
            ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
        cl = ws.cell(r1, c1, val)
        cl.font      = font or NF
        cl.alignment = align
        cl.border    = border
        if fill:
            cl.fill = fill
        return cl

    # ── 1. 기본정보 ──────────────────────────────────────────────
    ws.merge_cells("A1:J1")
    c = ws["A1"]
    c.value, c.font, c.fill, c.alignment = "1. 기본정보", BF, G_FILL, LA

    mc(2, 1, 2, 3, "차 종",        BF, CA, B_ALL, H_FILL)
    mc(2, 4, 2, 10, "자동차 등록번호", BF, CA, B_ALL, H_FILL)
    mc(3, 1, 3, 3,  VEHICLE_NAME,  NF, CA, B_ALL)
    mc(3, 4, 3, 10, VEHICLE_NUMBER, NF, CA, B_ALL)

    # row 4 blank
    ws.append([None])

    # ── 2. 업무용 사용비율 계산 ──────────────────────────────────
    ws.merge_cells("A5:J5")
    c = ws["A5"]
    c.value, c.font, c.fill, c.alignment = "2. 업무용 사용비율 계산", BF, G_FILL, LA

    # ── 헤더 (row 6, 7, 8) ───────────────────────────────────────
    ws.row_dimensions[6].height = 20
    ws.row_dimensions[7].height = 20
    ws.row_dimensions[8].height = 30

    mc(6, 1, 8, 1,  "사용일자\n(요일)",         BF, CA, B_ALL, H_FILL)
    mc(6, 2, 6, 3,  "사용자",                   BF, CA, B_ALL, H_FILL)
    mc(7, 2, 8, 2,  "부서",                     BF, CA, B_ALL, H_FILL)
    mc(7, 3, 8, 3,  "성명",                     BF, CA, B_ALL, H_FILL)
    mc(6, 4, 6, 9,  "운 행 내 역",              BF, CA, B_ALL, H_FILL)
    mc(7, 4, 8, 4,  "주행 전\n계기판의 거리",    BF, CA, B_ALL, H_FILL)
    mc(7, 5, 8, 5,  "주행 후\n계기판의 거리",    BF, CA, B_ALL, H_FILL)
    mc(7, 6, 8, 6,  "주행거리\n(km)",            BF, CA, B_ALL, H_FILL)
    mc(7, 7, 7, 8,  "업무용 사용거리(km)",       BF, CA, B_ALL, H_FILL)
    mc(8, 7, 8, 7,  "출/퇴근용\n(km)",           BF, CA, B_ALL, H_FILL)
    mc(8, 8, 8, 8,  "일반업무용\n(km)",           BF, CA, B_ALL, H_FILL)
    mc(7, 9, 8, 9,  "비 고",                    BF, CA, B_ALL, H_FILL)
    mc(6, 10, 8, 10, "충전금액",                 BF, CA, B_ALL, H_FILL)

    # ── 데이터 행 ────────────────────────────────────────────────
    WEEKDAYS = ["월", "화", "수", "목", "금", "토", "일"]
    data_row = 9
    total_distance = 0.0
    total_charging = 0.0

    for log in logs:
        d = datetime.strptime(log["drive_date"], "%Y-%m-%d")
        weekday = WEEKDAYS[d.weekday()]
        date_label = f"{d.month:02d}월 {d.day:02d}일({weekday})"

        odo_s  = log["odometer_start"] or 0
        odo_e  = log["odometer_end"]   or 0
        dist   = odo_e - odo_s if odo_e and odo_s else 0
        charge = log["charging_amount"] if log["charging"] else 0

        total_distance += dist
        total_charging += charge

        row_data = [
            date_label,
            log["department"],
            log["name"],
            int(odo_s),
            int(odo_e),
            int(dist),
            "X",
            int(dist),
            log["destination"] or "",
            int(charge) if charge else "-",
        ]

        ws.row_dimensions[data_row].height = 16
        for col, val in enumerate(row_data, 1):
            cl = ws.cell(data_row, col, val)
            cl.font      = NF
            cl.alignment = CA
            cl.border    = B_ALL
            if isinstance(val, int) and col in (4, 5, 6, 8, 10):
                cl.number_format = "#,##0"

        data_row += 1

    # ── 합계 행 ──────────────────────────────────────────────────
    sr = data_row
    ws.row_dimensions[sr].height = 18
    ws.row_dimensions[sr + 1].height = 22

    mc(sr, 1, sr, 3,  "과세기간 총주행 거리 (km)",    BF, CA, B_ALL, H_FILL)
    mc(sr, 4, sr, 6,  "",                              BF, CA, B_ALL, H_FILL)
    ws.cell(sr, 4).value = int(total_distance)
    ws.cell(sr, 4).number_format = "#,##0"

    mc(sr, 7, sr, 8,  "과세기간 업무용 사용거리 (km)", BF, CA, B_ALL, H_FILL)
    mc(sr, 9, sr, 9,  "업무사용비율",                  BF, CA, B_ALL, H_FILL)
    mc(sr, 10, sr, 10, "총 충전금액",                  BF, CA, B_ALL, H_FILL)

    mc(sr+1, 4, sr+1, 6, int(total_distance),
       Font(bold=True, name="맑은 고딕", size=12), CA, B_ALL)
    ws.cell(sr+1, 4).number_format = "#,##0"

    mc(sr+1, 7, sr+1, 8, int(total_distance),
       Font(bold=True, name="맑은 고딕", size=12), CA, B_ALL)
    ws.cell(sr+1, 7).number_format = "#,##0"

    mc(sr+1, 9, sr+1, 9, "100%",
       Font(bold=True, name="맑은 고딕", size=12), CA, B_ALL)

    mc(sr+1, 10, sr+1, 10, int(total_charging),
       Font(bold=True, name="맑은 고딕", size=12), CA, B_ALL)
    ws.cell(sr+1, 10).number_format = "#,##0"

    # ── 열 너비 ──────────────────────────────────────────────────
    col_widths = [14, 16, 8, 13, 13, 10, 10, 11, 20, 11]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# ────────────────────────────────────────────────────────────────
# 세션 초기화
# ────────────────────────────────────────────────────────────────
def init_session():
    defs = {
        "logged_in":          False,
        "user_phone":         "",
        "user_name":          "",
        "user_department":    DEPARTMENTS[0],
        "user_employee_id":   "",
        "admin_logged_in":    False,
        "show_admin_modal":   False,
        "show_reg_modal":     False,
        "selected_cal_date":  str(date.today()),
        "cal_year":           date.today().year,
        "cal_month":          date.today().month,
        "post_drive_log_id":  None,
    }
    for k, v in defs.items():
        if k not in st.session_state:
            st.session_state[k] = v


def try_auto_login():
    """쿠키에서 사용자 정보 복원"""
    if st.session_state.logged_in:
        return
    phone  = cookie_get("v_phone", "")
    emp_id = cookie_get("v_empid", "")
    if phone and emp_id and auth_user(phone, emp_id):
        user = get_user(phone)
        if user:
            st.session_state.logged_in        = True
            st.session_state.user_phone       = user["phone"]
            st.session_state.user_name        = user["name"]
            st.session_state.user_department  = user["department"]
            st.session_state.user_employee_id = user["employee_id"]


# ────────────────────────────────────────────────────────────────
# CSS
# ────────────────────────────────────────────────────────────────
def inject_css():
    st.markdown(
        """
        <style>
        /* 헤더 */
        .top-bar {
            display: flex; justify-content: space-between; align-items: center;
            background: #1a3a5c; color: white;
            padding: 10px 20px; border-radius: 8px; margin-bottom: 16px;
        }
        .top-bar h2 { margin: 0; font-size: 1.3rem; }
        .user-info  { font-size: 0.9rem; opacity: 0.9; }

        /* 달력 */
        .cal-wrap { width: 100%; }
        .cal-nav  { display: flex; align-items: center; gap: 8px; margin-bottom: 8px; }
        .cal-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 3px; }
        .cal-hdr  { text-align: center; font-weight: 700; padding: 4px 0;
                    background: #eef2f7; border-radius: 4px; font-size: 0.82rem; }
        .cal-day  {
            min-height: 58px; border: 1px solid #dde3ed; border-radius: 6px;
            padding: 4px; background: white; font-size: 0.8rem;
        }
        .cal-day.reserved   { background: #fff0f0; border-color: #ffa0a0; }
        .cal-day.today-cell { border: 2px solid #1a7f37; }
        .cal-day.empty      { background: #f7f8fa; border-color: #f0f0f0; }
        .day-num  { font-weight: 700; font-size: 0.88rem; }
        .sun      { color: #d0302a; }
        .sat      { color: #1a5fad; }
        .res-chip {
            font-size: 0.70rem; background: #ffd0d0; color: #8b1a1a;
            padding: 1px 4px; border-radius: 3px; margin-top: 2px;
            overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# ────────────────────────────────────────────────────────────────
# 달력 렌더 (Streamlit 버튼 그리드)
# ────────────────────────────────────────────────────────────────
def render_calendar(year, month, all_res):
    by_date: dict[str, list] = {}
    for r in all_res:
        by_date.setdefault(r["res_date"], []).append(r)

    today = date.today()
    cal   = cal_module.Calendar(firstweekday=6)  # 일요일 시작
    weeks = cal.monthdayscalendar(year, month)
    day_names = ["일", "월", "화", "수", "목", "금", "토"]

    # 요일 헤더
    hdr_cols = st.columns(7)
    for i, n in enumerate(day_names):
        cls = " sun" if i == 0 else (" sat" if i == 6 else "")
        hdr_cols[i].markdown(
            f'<div class="cal-hdr{cls}">{n}</div>', unsafe_allow_html=True
        )

    selected = None
    for week in weeks:
        cols = st.columns(7)
        for i, day in enumerate(week):
            with cols[i]:
                if day == 0:
                    st.markdown('<div class="cal-day empty"></div>', unsafe_allow_html=True)
                    continue

                day_str  = f"{year:04d}-{month:02d}-{day:02d}"
                res_list = by_date.get(day_str, [])
                is_today = date(year, month, day) == today

                # 날짜 버튼 레이블
                label = f"{'📅 ' if is_today else ''}{day}"
                if res_list:
                    label += f"\n🔴 {len(res_list)}건"

                btn_type = "primary" if is_today else "secondary"
                if st.button(label, key=f"cal_{year}_{month}_{day}",
                             use_container_width=True, type=btn_type):
                    selected = day_str

    return selected


# ────────────────────────────────────────────────────────────────
# 사용자 등록 / 로그인 패널
# ────────────────────────────────────────────────────────────────
def render_user_panel():
    if st.session_state.logged_in:
        st.sidebar.success(
            f"**{st.session_state.user_name}** ({st.session_state.user_department})\n\n"
            f"📞 {st.session_state.user_phone}"
        )
        if st.sidebar.button("로그아웃"):
            for k in ["logged_in", "user_phone", "user_name",
                      "user_department", "user_employee_id"]:
                st.session_state[k] = "" if k != "logged_in" else False
            cookie_set("v_phone", "")
            cookie_set("v_empid", "")
            st.rerun()
        return

    st.sidebar.header("로그인 / 계정 등록")

    tab_login, tab_reg = st.sidebar.tabs(["로그인", "계정 등록"])

    with tab_login:
        phone  = st.text_input("전화번호 (ID)", key="si_phone",
                               placeholder="010-0000-0000")
        emp_id = st.text_input("사번 (PW)",    key="si_emp",
                               type="password", placeholder="사번 입력")
        if st.button("로그인", key="btn_login"):
            if auth_user(phone, emp_id):
                user = get_user(phone)
                st.session_state.logged_in        = True
                st.session_state.user_phone       = user["phone"]
                st.session_state.user_name        = user["name"]
                st.session_state.user_department  = user["department"]
                st.session_state.user_employee_id = user["employee_id"]
                cookie_set("v_phone", phone)
                cookie_set("v_empid", emp_id)
                st.rerun()
            else:
                st.error("전화번호 또는 사번이 올바르지 않습니다.")

    with tab_reg:
        st.caption("처음 사용 시 한 번만 등록하면 다음부터 자동 입력됩니다.")
        r_dept  = st.selectbox("부서", DEPARTMENTS, key="reg_dept")
        r_name  = st.text_input("이름",   key="reg_name")
        r_phone = st.text_input("전화번호", key="reg_phone",
                                placeholder="010-0000-0000")
        r_emp   = st.text_input("사번",   key="reg_emp",
                                type="password", placeholder="사번 입력")
        if st.button("등록 / 정보 수정", key="btn_reg"):
            if r_name and r_phone and r_emp:
                register_user(r_phone, r_emp, r_dept, r_name)
                st.session_state.logged_in        = True
                st.session_state.user_phone       = r_phone
                st.session_state.user_name        = r_name
                st.session_state.user_department  = r_dept
                st.session_state.user_employee_id = r_emp
                cookie_set("v_phone", r_phone)
                cookie_set("v_empid", r_emp)
                st.success("등록 완료!")
                st.rerun()
            else:
                st.warning("모든 항목을 입력해 주세요.")


# ────────────────────────────────────────────────────────────────
# 탭 1 : 예약하기
# ────────────────────────────────────────────────────────────────
def tab_reservation():
    st.subheader("차량 예약 현황")

    all_res = get_all_reservations()

    # 달력 월 이동
    col_prev, col_title, col_next = st.columns([1, 4, 1])
    with col_prev:
        if st.button("◀ 이전달", key="cal_prev"):
            m = st.session_state.cal_month - 1
            if m < 1:
                st.session_state.cal_month = 12
                st.session_state.cal_year -= 1
            else:
                st.session_state.cal_month = m
    with col_title:
        st.markdown(
            f"<h3 style='text-align:center;margin:0'>"
            f"{st.session_state.cal_year}년 {st.session_state.cal_month}월"
            f"</h3>",
            unsafe_allow_html=True,
        )
    with col_next:
        if st.button("다음달 ▶", key="cal_next"):
            m = st.session_state.cal_month + 1
            if m > 12:
                st.session_state.cal_month = 1
                st.session_state.cal_year += 1
            else:
                st.session_state.cal_month = m

    # 달력 렌더
    clicked = render_calendar(
        st.session_state.cal_year,
        st.session_state.cal_month,
        all_res,
    )
    if clicked:
        st.session_state.selected_cal_date = clicked

    st.divider()

    # 선택된 날짜 처리
    sel_date = st.session_state.selected_cal_date
    sel_res  = get_reservations_by_date(sel_date)

    col_info, col_form = st.columns([1, 1])

    with col_info:
        st.markdown(f"#### 📅 {sel_date} 예약 현황")
        if not sel_res:
            st.info("이 날짜에 예약이 없습니다. 오른쪽에서 예약하세요.")
        else:
            for r in sel_res:
                with st.expander(f"🕐 {r['res_time']} │ {r['department']} {r['name']}", expanded=True):
                    st.write(f"**부서:** {r['department']}")
                    st.write(f"**이름:** {r['name']}")
                    st.write(f"**전화번호:** {r['phone']}")
                    st.write(f"**시간:** {r['res_time']}")
                    st.write(f"**방문지:** {r['destination']}")
                    # 본인 예약 삭제
                    if (st.session_state.logged_in and
                            r["user_phone"] == st.session_state.user_phone):
                        if st.button("예약 취소", key=f"del_{r['id']}"):
                            delete_reservation(r["id"])
                            st.success("예약이 취소되었습니다.")
                            st.rerun()

    with col_form:
        st.markdown("#### ✏️ 예약 등록")
        if not st.session_state.logged_in:
            st.warning("예약하려면 왼쪽 사이드바에서 로그인하세요.")
            return

        f_dept = st.selectbox(
            "부서",
            DEPARTMENTS,
            index=DEPARTMENTS.index(st.session_state.user_department)
            if st.session_state.user_department in DEPARTMENTS else 0,
            key="res_dept",
        )
        f_name = st.text_input("이름", value=st.session_state.user_name, key="res_name")
        f_date = st.date_input(
            "날짜",
            value=datetime.strptime(sel_date, "%Y-%m-%d").date(),
            key="res_date",
        )
        f_time = st.time_input("시간", value=datetime.strptime("09:00", "%H:%M").time(),
                               step=1800, key="res_time")
        f_dest = st.text_input("방문지", key="res_dest")

        if st.button("예약 등록", type="primary", key="btn_add_res"):
            if f_name and f_dest:
                add_reservation(
                    st.session_state.user_phone,
                    f_dept, f_name,
                    st.session_state.user_phone,
                    str(f_date),
                    f_time.strftime("%H:%M"),
                    f_dest,
                )
                st.success("예약이 등록되었습니다!")
                st.rerun()
            else:
                st.warning("이름과 방문지를 입력해 주세요.")


# ────────────────────────────────────────────────────────────────
# 탭 2 : 주행 전 기록
# ────────────────────────────────────────────────────────────────
def tab_pre_drive():
    st.subheader("주행 전 기록")

    if not st.session_state.logged_in:
        st.warning("사이드바에서 로그인하세요.")
        return

    with st.form("form_pre"):
        col1, col2 = st.columns(2)
        with col1:
            p_dept = st.selectbox(
                "부서",
                DEPARTMENTS,
                index=DEPARTMENTS.index(st.session_state.user_department)
                if st.session_state.user_department in DEPARTMENTS else 0,
            )
            p_name = st.text_input("이름", value=st.session_state.user_name)
            p_date = st.date_input("주행 날짜", value=date.today())
        with col2:
            p_odo  = st.number_input("출발 시 계기판 거리 (km)", min_value=0.0,
                                     step=1.0, format="%.0f")
            p_dest = st.text_input("목적지")
            p_comp = st.text_input("동행인 (부서/이름, 여러 명은 쉼표 구분)",
                                   placeholder="예: DFX그룹/홍길동, 시공기술/이순신")

        submitted = st.form_submit_button("주행 전 기록 저장", type="primary")
        if submitted:
            if p_dest and p_odo >= 0:
                add_pre_drive(
                    st.session_state.user_phone,
                    p_dept, p_name,
                    st.session_state.user_phone,
                    str(p_date), p_odo, p_comp, p_dest,
                )
                st.success("주행 전 기록이 저장되었습니다.")
                st.rerun()
            else:
                st.warning("목적지와 계기판 거리를 입력해 주세요.")

    # 완료되지 않은 기록 목록
    pre_drives = get_pre_drives(st.session_state.user_phone)
    if pre_drives:
        st.divider()
        st.markdown("##### 주행 완료 대기 중인 기록")
        for p in pre_drives:
            st.info(
                f"📅 {p['drive_date']}  |  출발 계기판: **{p['odometer_start']:,.0f} km**  "
                f"|  목적지: {p['destination']}  |  동행: {p['companions'] or '-'}"
            )


# ────────────────────────────────────────────────────────────────
# 탭 3 : 주행 후 기록
# ────────────────────────────────────────────────────────────────
def tab_post_drive():
    st.subheader("주행 후 기록")

    if not st.session_state.logged_in:
        st.warning("사이드바에서 로그인하세요.")
        return

    pre_drives = get_pre_drives(st.session_state.user_phone)
    if not pre_drives:
        st.info("완료 대기 중인 주행 전 기록이 없습니다. 먼저 '주행 전 기록' 탭에서 기록하세요.")
        return

    options = {
        f"{p['drive_date']} │ {p['odometer_start']:,.0f}km 출발 │ {p['destination']}": p
        for p in pre_drives
    }
    sel_label = st.selectbox("완료할 주행 기록 선택", list(options.keys()))
    sel_log   = options[sel_label]

    st.markdown(
        f"> **출발 기준**  \n"
        f"> 날짜: {sel_log['drive_date']}  \n"
        f"> 출발 계기판: {sel_log['odometer_start']:,.0f} km  \n"
        f"> 목적지: {sel_log['destination']}  \n"
        f"> 동행인: {sel_log['companions'] or '없음'}"
    )

    with st.form("form_post"):
        col1, col2 = st.columns(2)
        with col1:
            q_date    = st.date_input("도착 날짜", value=date.today())
            q_odo_end = st.number_input(
                "도착 시 계기판 거리 (km)",
                min_value=float(sel_log["odometer_start"] or 0),
                step=1.0, format="%.0f",
            )
            if q_odo_end and sel_log["odometer_start"]:
                driven = q_odo_end - sel_log["odometer_start"]
                st.metric("주행 거리", f"{driven:,.0f} km")

        with col2:
            q_park    = st.text_input("주차 장소")
            q_comp    = st.text_input(
                "동행인",
                value=sel_log["companions"] or "",
                help="주행 전 입력 값이 기본으로 표시됩니다.",
            )
            q_charge  = st.checkbox("충전 여부")
            q_charge_amt = 0.0
            if q_charge:
                q_charge_amt = st.number_input("충전 금액 (원)", min_value=0.0, step=100.0,
                                               format="%.0f")

        submitted = st.form_submit_button("주행 후 기록 완료", type="primary")
        if submitted:
            if q_park and q_odo_end >= (sel_log["odometer_start"] or 0):
                complete_drive(
                    sel_log["id"],
                    q_odo_end,
                    int(q_charge),
                    q_charge_amt,
                    q_park,
                    q_comp,
                    str(q_date),
                )
                st.success("주행 기록이 완료 처리되었습니다!")
                st.rerun()
            else:
                st.warning("주차 장소를 입력하고 계기판 거리를 확인해 주세요.")


# ────────────────────────────────────────────────────────────────
# 관리자 화면
# ────────────────────────────────────────────────────────────────
def admin_panel():
    st.subheader("관리자 화면 — 운행기록 조회 및 내보내기")

    # 기간 선택
    col1, col2 = st.columns(2)
    with col1:
        start_d = st.date_input(
            "시작일",
            value=date.today().replace(day=1),
            key="adm_start",
        )
    with col2:
        end_d = st.date_input("종료일", value=date.today(), key="adm_end")

    if st.button("조회", type="primary"):
        st.session_state["adm_logs"] = get_logs_by_period(str(start_d), str(end_d))

    logs = st.session_state.get("adm_logs")
    if logs is None:
        return

    if not logs:
        st.warning("해당 기간에 완료된 운행 기록이 없습니다.")
        return

    # 테이블 표시
    WEEKDAYS = ["월", "화", "수", "목", "금", "토", "일"]
    rows = []
    for log in logs:
        d    = datetime.strptime(log["drive_date"], "%Y-%m-%d")
        wday = WEEKDAYS[d.weekday()]
        dist = (log["odometer_end"] or 0) - (log["odometer_start"] or 0)
        rows.append({
            "날짜":         f"{log['drive_date']} ({wday})",
            "부서":         log["department"],
            "이름":         log["name"],
            "출발 계기판":  f"{log['odometer_start']:,.0f}" if log["odometer_start"] else "-",
            "도착 계기판":  f"{log['odometer_end']:,.0f}"   if log["odometer_end"]   else "-",
            "주행거리(km)": f"{dist:,.0f}",
            "목적지(비고)": log["destination"] or "",
            "충전금액":     f"{int(log['charging_amount']):,}" if log["charging"] else "-",
        })

    df = pd.DataFrame(rows)
    st.dataframe(df, use_container_width=True, hide_index=True)

    total_dist = sum(
        (l["odometer_end"] or 0) - (l["odometer_start"] or 0) for l in logs
    )
    total_chg  = sum(l["charging_amount"] for l in logs if l["charging"])
    c1, c2, c3 = st.columns(3)
    c1.metric("총 운행 건수", f"{len(logs)} 건")
    c2.metric("총 주행 거리", f"{total_dist:,.0f} km")
    c3.metric("총 충전 금액", f"{int(total_chg):,} 원")

    # 엑셀 다운로드
    excel_buf = make_excel(logs, str(start_d), str(end_d))
    st.download_button(
        label="📥 엑셀 운행기록부 다운로드",
        data=excel_buf,
        file_name=f"운행기록부_{start_d}_{end_d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )


# ────────────────────────────────────────────────────────────────
# MAIN
# ────────────────────────────────────────────────────────────────
def main():
    init_db()
    init_session()
    try_auto_login()
    inject_css()

    # ── 상단 바 ──────────────────────────────────────────────────
    user_str = (
        f"{st.session_state.user_name} ({st.session_state.user_department})"
        if st.session_state.logged_in
        else "로그인 필요"
    )
    col_title, col_user, col_admin = st.columns([4, 3, 1])
    with col_title:
        st.markdown(
            "<h2 style='margin:0;color:#1a3a5c'>🚗 공용차량 관리 시스템</h2>"
            f"<small style='color:#888'>차종: {VEHICLE_NAME} │ 번호: {VEHICLE_NUMBER}</small>",
            unsafe_allow_html=True,
        )
    with col_user:
        if st.session_state.logged_in:
            st.markdown(
                f"<div style='text-align:right;padding-top:8px;color:#1a3a5c'>"
                f"👤 {user_str}</div>",
                unsafe_allow_html=True,
            )
    with col_admin:
        if st.session_state.admin_logged_in:
            if st.button("🔓 관리자", type="secondary"):
                st.session_state.admin_logged_in = False
                st.rerun()
        else:
            if st.button("🔒 관리자", type="secondary"):
                st.session_state.show_admin_modal = True

    # 사이드바에 로그인 패널
    render_user_panel()

    # ── 관리자 로그인 모달 ───────────────────────────────────────
    if st.session_state.show_admin_modal and not st.session_state.admin_logged_in:
        with st.container(border=True):
            st.markdown("#### 관리자 인증")
            pw = st.text_input("비밀번호", type="password", key="adm_pw_input")
            c1, c2 = st.columns(2)
            if c1.button("확인", key="adm_pw_ok"):
                if pw == ADMIN_PASSWORD:
                    st.session_state.admin_logged_in  = True
                    st.session_state.show_admin_modal = False
                    st.rerun()
                else:
                    st.error("비밀번호가 틀렸습니다.")
            if c2.button("취소", key="adm_pw_cancel"):
                st.session_state.show_admin_modal = False
                st.rerun()

    st.divider()

    # ── 탭 분기 ──────────────────────────────────────────────────
    if st.session_state.admin_logged_in:
        tabs = st.tabs(["📅 예약하기", "🚀 주행 전 기록", "🏁 주행 후 기록", "🔐 관리자 화면"])
        with tabs[0]: tab_reservation()
        with tabs[1]: tab_pre_drive()
        with tabs[2]: tab_post_drive()
        with tabs[3]: admin_panel()
    else:
        tabs = st.tabs(["📅 예약하기", "🚀 주행 전 기록", "🏁 주행 후 기록"])
        with tabs[0]: tab_reservation()
        with tabs[1]: tab_pre_drive()
        with tabs[2]: tab_post_drive()


if __name__ == "__main__":
    main()
