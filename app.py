# 공용차량 예약 및 주행거리 정산 시스템
import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime, date
from io import BytesIO
import calendar as cal_module
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="공용차량 관리",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded",
)

DEPARTMENTS = [
    "생산기술팀", "공법기술", "의장기술", "운반종합기술",
    "시공기술", "족장기술", "DFX그룹", "도장기술",
]
ADMIN_PASSWORD  = "1111"
DB_PATH         = "vehicle_management.db"
VEHICLE_NAME    = "EV3"
VEHICLE_NUMBER  = "05하 7211"
WEEKDAYS        = ["월", "화", "수", "목", "금", "토", "일"]


# ════════════════════════════════════════════════════════════════
# DB
# ════════════════════════════════════════════════════════════════
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
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            user_phone   TEXT NOT NULL,
            department   TEXT NOT NULL,
            name         TEXT NOT NULL,
            phone        TEXT NOT NULL,
            res_date     TEXT NOT NULL,
            res_time     TEXT NOT NULL,
            res_time_end TEXT NOT NULL DEFAULT '',
            destination  TEXT NOT NULL,
            created_at   TEXT DEFAULT (datetime('now','localtime'))
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
            charging_amount  REAL DEFAULT 0,
            parking_location TEXT,
            status           TEXT DEFAULT 'pre',
            created_at       TEXT DEFAULT (datetime('now','localtime'))
        );
        """)
        for ddl in [
            "ALTER TABLE reservations ADD COLUMN res_time_end TEXT NOT NULL DEFAULT ''",
            "ALTER TABLE driving_logs ADD COLUMN charging_amount REAL DEFAULT 0",
        ]:
            try:
                conn.execute(ddl)
            except Exception:
                pass
        conn.commit()


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


# ── 사용자 ──────────────────────────────────────────────────────
def get_user(phone):
    return _query("SELECT * FROM users WHERE phone=?", (phone,), one=True)

def register_user(phone, emp_id, dept, name):
    _exec("INSERT OR REPLACE INTO users VALUES (?,?,?,?)", (phone, emp_id, dept, name))

def auth_user(phone, emp_id):
    return _query(
        "SELECT 1 FROM users WHERE phone=? AND employee_id=?",
        (phone, emp_id), one=True,
    ) is not None


# ── 예약 ────────────────────────────────────────────────────────
def get_all_reservations():
    return _query("SELECT * FROM reservations ORDER BY res_date, res_time")

def get_reservations_by_date(d):
    return _query(
        "SELECT * FROM reservations WHERE res_date=? ORDER BY res_time", (d,))

def add_reservation(user_phone, dept, name, phone,
                    res_date, res_time, res_time_end, dest):
    _exec(
        "INSERT INTO reservations "
        "(user_phone,department,name,phone,res_date,res_time,res_time_end,destination) "
        "VALUES (?,?,?,?,?,?,?,?)",
        (user_phone, dept, name, phone, res_date, res_time, res_time_end, dest),
    )

def update_reservation(rid, dept, name, res_date, res_time, res_time_end, dest):
    _exec(
        "UPDATE reservations "
        "SET department=?,name=?,res_date=?,res_time=?,res_time_end=?,destination=? "
        "WHERE id=?",
        (dept, name, res_date, res_time, res_time_end, dest, rid),
    )

def delete_reservation(rid):
    _exec("DELETE FROM reservations WHERE id=?", (rid,))


# ── 운행 기록 ────────────────────────────────────────────────────
def get_pre_drives(user_phone):
    return _query(
        "SELECT * FROM driving_logs WHERE user_phone=? AND status='pre' "
        "ORDER BY drive_date DESC", (user_phone,))

def get_user_all_logs(user_phone):
    return _query(
        "SELECT * FROM driving_logs WHERE user_phone=? "
        "ORDER BY drive_date DESC, created_at DESC", (user_phone,))

def add_pre_drive(user_phone, dept, name, phone,
                  drive_date, odo_start, companions, dest):
    _exec(
        "INSERT INTO driving_logs "
        "(user_phone,department,name,phone,drive_date,"
        "odometer_start,companions,destination) VALUES (?,?,?,?,?,?,?,?)",
        (user_phone, dept, name, phone, drive_date, odo_start, companions, dest),
    )

def complete_drive(lid, odo_end, charge_amt, parking, companions, drive_date):
    _exec(
        "UPDATE driving_logs "
        "SET odometer_end=?,charging_amount=?,parking_location=?,"
        "companions=?,drive_date=?,status='complete' WHERE id=?",
        (odo_end, charge_amt, parking, companions, drive_date, lid),
    )

def update_drive_log(lid, drive_date, odo_start, odo_end,
                     companions, dest, charge_amt, parking, status):
    _exec(
        "UPDATE driving_logs "
        "SET drive_date=?,odometer_start=?,odometer_end=?,companions=?,"
        "destination=?,charging_amount=?,parking_location=?,status=? WHERE id=?",
        (drive_date, odo_start, odo_end, companions,
         dest, charge_amt, parking, status, lid),
    )

def delete_drive_log(lid):
    _exec("DELETE FROM driving_logs WHERE id=?", (lid,))

def get_logs_by_period(start, end):
    return _query(
        "SELECT * FROM driving_logs WHERE drive_date BETWEEN ? AND ? "
        "AND status='complete' ORDER BY drive_date, created_at",
        (start, end),
    )




# ════════════════════════════════════════════════════════════════
# 엑셀
# ════════════════════════════════════════════════════════════════
def make_excel(logs, start_date, end_date):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "운행기록부"
    ws.sheet_view.showGridLines = False

    thin   = Side(style="thin")
    B_ALL  = Border(left=thin, right=thin, top=thin, bottom=thin)
    CA     = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LA     = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    BF     = Font(bold=True, name="맑은 고딕", size=10)
    NF     = Font(name="맑은 고딕", size=10)
    H_FILL = PatternFill("solid", fgColor="BDD7EE")
    G_FILL = PatternFill("solid", fgColor="D9D9D9")

    def mc(r1, c1, r2, c2, val="", font=None, align=CA, border=B_ALL, fill=None):
        if r1 != r2 or c1 != c2:
            ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
        cl = ws.cell(r1, c1, val)
        cl.font, cl.alignment, cl.border = font or NF, align, border
        if fill:
            cl.fill = fill
        return cl

    ws.merge_cells("A1:J1")
    ws["A1"].value, ws["A1"].font = "1. 기본정보", BF
    ws["A1"].fill,  ws["A1"].alignment = G_FILL, LA
    mc(2,1,2,3,"차 종",BF,CA,B_ALL,H_FILL); mc(2,4,2,10,"자동차 등록번호",BF,CA,B_ALL,H_FILL)
    mc(3,1,3,3,VEHICLE_NAME,NF,CA,B_ALL);   mc(3,4,3,10,VEHICLE_NUMBER,NF,CA,B_ALL)
    ws.append([None])
    ws.merge_cells("A5:J5")
    ws["A5"].value, ws["A5"].font = "2. 업무용 사용비율 계산", BF
    ws["A5"].fill,  ws["A5"].alignment = G_FILL, LA
    for r in (6,7,8): ws.row_dimensions[r].height = 24
    mc(6,1,8,1,"사용일자\n(요일)",BF,CA,B_ALL,H_FILL)
    mc(6,2,6,3,"사용자",BF,CA,B_ALL,H_FILL); mc(7,2,8,2,"부서",BF,CA,B_ALL,H_FILL)
    mc(7,3,8,3,"성명",BF,CA,B_ALL,H_FILL);   mc(6,4,6,9,"운 행 내 역",BF,CA,B_ALL,H_FILL)
    mc(7,4,8,4,"주행 전\n계기판의 거리",BF,CA,B_ALL,H_FILL)
    mc(7,5,8,5,"주행 후\n계기판의 거리",BF,CA,B_ALL,H_FILL)
    mc(7,6,8,6,"주행거리\n(km)",BF,CA,B_ALL,H_FILL)
    mc(7,7,7,8,"업무용 사용거리(km)",BF,CA,B_ALL,H_FILL)
    mc(8,7,8,7,"출/퇴근용\n(km)",BF,CA,B_ALL,H_FILL); mc(8,8,8,8,"일반업무용\n(km)",BF,CA,B_ALL,H_FILL)
    mc(7,9,8,9,"비 고",BF,CA,B_ALL,H_FILL);   mc(6,10,8,10,"충전금액",BF,CA,B_ALL,H_FILL)

    dr = 9; td = tc = 0.0
    for log in logs:
        d = datetime.strptime(log["drive_date"], "%Y-%m-%d")
        lbl = f"{d.month:02d}월 {d.day:02d}일({WEEKDAYS[d.weekday()]})"
        s = log["odometer_start"] or 0; e = log["odometer_end"] or 0
        dist = e - s; chrg = log.get("charging_amount") or 0
        td += dist; tc += chrg
        ws.row_dimensions[dr].height = 16
        for col, val in enumerate(
            [lbl, log["department"], log["name"],
             int(s), int(e), int(dist), "X", int(dist),
             log["destination"] or "", int(chrg) if chrg else "-"], 1
        ):
            cl = ws.cell(dr, col, val)
            cl.font, cl.alignment, cl.border = NF, CA, B_ALL
            if isinstance(val, int) and col in (4,5,6,8,10):
                cl.number_format = "#,##0"
        dr += 1

    sr = dr
    ws.row_dimensions[sr].height = 18; ws.row_dimensions[sr+1].height = 22
    mc(sr,1,sr,3,"과세기간 총주행 거리 (km)",BF,CA,B_ALL,H_FILL)
    mc(sr,4,sr,6,int(td),BF,CA,B_ALL,H_FILL); ws.cell(sr,4).number_format="#,##0"
    mc(sr,7,sr,8,"과세기간 업무용 사용거리 (km)",BF,CA,B_ALL,H_FILL)
    mc(sr,9,sr,9,"업무사용비율",BF,CA,B_ALL,H_FILL)
    mc(sr,10,sr,10,"총 충전금액",BF,CA,B_ALL,H_FILL)
    BF2 = Font(bold=True, name="맑은 고딕", size=12)
    mc(sr+1,4,sr+1,6,int(td),BF2,CA,B_ALL); ws.cell(sr+1,4).number_format="#,##0"
    mc(sr+1,7,sr+1,8,int(td),BF2,CA,B_ALL); ws.cell(sr+1,7).number_format="#,##0"
    mc(sr+1,9,sr+1,9,"100%",BF2,CA,B_ALL)
    mc(sr+1,10,sr+1,10,int(tc),BF2,CA,B_ALL); ws.cell(sr+1,10).number_format="#,##0"
    for i, w in enumerate([14,16,8,13,13,10,10,11,20,11],1):
        ws.column_dimensions[get_column_letter(i)].width = w
    out = BytesIO(); wb.save(out); out.seek(0)
    return out


# ════════════════════════════════════════════════════════════════
# 세션 초기화
# ════════════════════════════════════════════════════════════════
def init_session():
    defs = {
        "logged_in":        False,
        "user_phone":       "",
        "user_name":        "",
        "user_department":  DEPARTMENTS[0],
        "user_employee_id": "",
        "admin_logged_in":  False,
        "show_admin_modal": False,
        "selected_cal_date": str(date.today()),
        "date_picker_main": date.today(),   # 달력 클릭 → date_input 자동 반영용
        "cal_year":         date.today().year,
        "cal_month":        date.today().month,
        "adm_logs":         None,
        "confirm_del_log":  None,
        "confirm_del_res":  None,
        "editing_res_id":   None,
        "editing_log_id":   None,
    }
    for k, v in defs.items():
        if k not in st.session_state:
            st.session_state[k] = v


# ════════════════════════════════════════════════════════════════
# 쿼리 파라미터 처리 (달력 클릭 / 월 이동)
# ════════════════════════════════════════════════════════════════
def handle_query_params():
    params = st.query_params
    changed = False

    if "cal_date" in params:
        try:
            d = datetime.strptime(params["cal_date"], "%Y-%m-%d").date()
            st.session_state.selected_cal_date = str(d)
            st.session_state.cal_year  = d.year
            st.session_state.cal_month = d.month
            # date_input 위젯 키에도 직접 주입 → 달력 클릭 시 자동 반영
            st.session_state["date_picker_main"] = d
        except ValueError:
            pass
        changed = True

    if changed:
        st.query_params.clear()


# ════════════════════════════════════════════════════════════════
# 달력 월 네비게이션 — st.button (전체 페이지 로딩 없이 달력만 갱신)
# ════════════════════════════════════════════════════════════════
def render_month_nav(year, month):
    # 모바일에서도 한 줄 유지
    st.markdown("""
    <style>
    div[data-testid="stHorizontalBlock"].month-nav {
        flex-wrap: nowrap !important;
        align-items: center !important;
        gap: 4px !important;
    }
    div[data-testid="stHorizontalBlock"].month-nav > div {
        min-width: 0 !important;
    }
    </style>
    <div class="month-nav">
    """, unsafe_allow_html=True)

    col_prev, col_label, col_next = st.columns([2, 5, 2])
    with col_prev:
        if st.button("◀ 이전달", key="btn_cal_prev", use_container_width=True):
            if month == 1:
                st.session_state.cal_year  = year - 1
                st.session_state.cal_month = 12
            else:
                st.session_state.cal_month = month - 1
            st.rerun()
    with col_label:
        st.markdown(
            f"<div style='text-align:center;font-size:1.3rem;font-weight:700;"
            f"color:#1a3a5c;padding:6px 0'>{year}년 {month}월</div>",
            unsafe_allow_html=True,
        )
    with col_next:
        if st.button("다음달 ▶", key="btn_cal_next", use_container_width=True):
            if month == 12:
                st.session_state.cal_year  = year + 1
                st.session_state.cal_month = 1
            else:
                st.session_state.cal_month = month + 1
            st.rerun()


# ════════════════════════════════════════════════════════════════
# HTML 달력 (셀 클릭 → ?cal_date=YYYY-MM-DD)
# ════════════════════════════════════════════════════════════════
def build_html_calendar(year, month, all_res):
    by_date: dict = {}
    for r in all_res:
        by_date.setdefault(r["res_date"], []).append(r)

    today = date.today()
    sel   = st.session_state.selected_cal_date
    weeks = cal_module.Calendar(firstweekday=6).monthdayscalendar(year, month)

    html = """
    <style>
    .vc-cal{width:100%;border-collapse:collapse;font-family:'Malgun Gothic',sans-serif}
    .vc-cal th{background:#1a3a5c;color:#fff;padding:7px 0;text-align:center;
               font-size:0.85rem;white-space:nowrap}
    .vc-cal td{border:1px solid #dde3ed;vertical-align:top;padding:4px;
               min-height:64px;width:14.28%;font-size:0.78rem;
               cursor:pointer;transition:background .12s}
    .vc-cal td:hover{background:#e8f0fe !important}
    .vc-empty{background:#f7f8fa;cursor:default !important}
    .vc-empty:hover{background:#f7f8fa !important}
    .vc-today{border:2px solid #1a7f37 !important;background:#f0fff4}
    .vc-reserved{background:#fff5f5}
    .vc-selected{background:#dbeafe !important;border:2px solid #2563eb !important}
    .vc-num{font-weight:700;font-size:0.9rem;display:block}
    .vc-sun{color:#d0302a}.vc-sat{color:#1a5fad}
    .vc-chip{display:block;font-size:0.67rem;background:#ffd0d0;color:#7b1515;
             padding:1px 4px;border-radius:3px;margin-top:2px;
             overflow:hidden;white-space:nowrap;text-overflow:ellipsis}
    </style>
    <table class="vc-cal">
    <tr>
      <th class="vc-sun">일</th><th>월</th><th>화</th><th>수</th>
      <th>목</th><th>금</th><th class="vc-sat">토</th>
    </tr>
    """
    for week in weeks:
        html += "<tr>"
        for i, day in enumerate(week):
            if day == 0:
                html += '<td class="vc-empty"></td>'
                continue
            day_str  = f"{year:04d}-{month:02d}-{day:02d}"
            res_list = by_date.get(day_str, [])
            is_today = date(year, month, day) == today
            is_sel   = day_str == sel

            cls = []
            if is_sel:        cls.append("vc-selected")
            elif is_today:    cls.append("vc-today")
            if res_list and not is_sel: cls.append("vc-reserved")

            nc    = "vc-sun" if i == 0 else ("vc-sat" if i == 6 else "")
            chips = "".join(
                f'<span class="vc-chip">'
                f'{r["res_time"]}~{r.get("res_time_end","")} {r["name"]}'
                f'</span>'
                for r in res_list[:3]
            )
            if len(res_list) > 3:
                chips += f'<span class="vc-chip">+{len(res_list)-3}건 더</span>'

            onclick = f"window.location.href='?cal_date={day_str}'"
            html += (
                f'<td class="{" ".join(cls)}" onclick="{onclick}">'
                f'<span class="vc-num {nc}">{day}</span>'
                f'{chips}</td>'
            )
        html += "</tr>"
    html += "</table>"
    return html


# ════════════════════════════════════════════════════════════════
# 사이드바
# ════════════════════════════════════════════════════════════════
def render_user_panel():
    st.sidebar.title("👤 사용자")
    if st.session_state.logged_in:
        st.sidebar.success(
            f"**{st.session_state.user_name}**\n\n"
            f"{st.session_state.user_department}\n\n"
            f"📞 {st.session_state.user_phone}"
        )
        if st.sidebar.button("로그아웃", use_container_width=True):
            for k in ["logged_in","user_phone","user_name",
                      "user_department","user_employee_id"]:
                st.session_state[k] = False if k=="logged_in" else ""
            st.rerun()
        return

    tab_l, tab_r = st.sidebar.tabs(["로그인", "계정 등록"])
    with tab_l:
        phone  = st.text_input("전화번호 (ID)", key="si_phone",
                               placeholder="010-0000-0000")
        emp_id = st.text_input("비밀번호", key="si_emp",
                               type="password", placeholder="사번 입력")
        if st.button("로그인", key="btn_login", use_container_width=True):
            if auth_user(phone, emp_id):
                user = get_user(phone)
                st.session_state.update(
                    logged_in=True, user_phone=user["phone"],
                    user_name=user["name"], user_department=user["department"],
                    user_employee_id=user["employee_id"],
                )
                st.rerun()
            else:
                st.error("전화번호 또는 비밀번호가 올바르지 않습니다.")
    with tab_r:
        st.caption("처음 사용 시 한 번만 등록하면 됩니다.")
        r_dept  = st.selectbox("부서", DEPARTMENTS, key="reg_dept")
        r_name  = st.text_input("이름", key="reg_name")
        r_phone = st.text_input("전화번호", key="reg_phone",
                                placeholder="010-0000-0000")
        r_emp   = st.text_input("비밀번호 (사번 입력)", key="reg_emp",
                                type="password")
        if st.button("등록 / 정보 수정", key="btn_reg", use_container_width=True):
            if r_name and r_phone and r_emp:
                register_user(r_phone, r_emp, r_dept, r_name)
                st.session_state.update(
                    logged_in=True, user_phone=r_phone, user_name=r_name,
                    user_department=r_dept, user_employee_id=r_emp,
                )
                st.success("등록 완료!")
                st.rerun()
            else:
                st.warning("모든 항목을 입력해 주세요.")


# ════════════════════════════════════════════════════════════════
# 탭 1 : 예약하기
# ════════════════════════════════════════════════════════════════
def _dept_idx(dept):
    return DEPARTMENTS.index(dept) if dept in DEPARTMENTS else 0


def _time_val(t_str, fallback="09:00"):
    try:
        return datetime.strptime(t_str or fallback, "%H:%M").time()
    except ValueError:
        return datetime.strptime(fallback, "%H:%M").time()


def tab_reservation():
    st.subheader("차량 예약 현황")
    all_res = get_all_reservations()

    # ── 달력 월 이동 (항상 한 줄 HTML) ──────────────────────────
    render_month_nav(st.session_state.cal_year, st.session_state.cal_month)

    # ── HTML 달력 ────────────────────────────────────────────────
    st.markdown(
        build_html_calendar(st.session_state.cal_year,
                            st.session_state.cal_month, all_res),
        unsafe_allow_html=True,
    )
    st.caption("날짜 클릭 → 해당 날짜로 이동  │  🔵=선택됨  │  붉은 배경=예약 있음")
    st.markdown("---")

    col_left, col_right = st.columns(2)

    # ── 예약 확인 ──────────────────────────────────────────────
    with col_left:
        st.markdown("#### 📅 예약 확인")
        st.caption("달력에서 날짜를 클릭하면 자동으로 반영됩니다.")

        # 달력 클릭 시 date_picker_main 키로 값이 주입되어 자동 반영
        sel = st.date_input(
            "날짜",
            key="date_picker_main",
            label_visibility="collapsed",
        )
        # 수동으로 날짜를 바꿨을 때 달력 하이라이트도 동기화
        if sel and str(sel) != st.session_state.selected_cal_date:
            st.session_state.selected_cal_date = str(sel)

        sel_res = get_reservations_by_date(str(sel))

        st.markdown(f"##### 📅 {sel} 예약 현황")
        if not sel_res:
            st.info("이 날짜에 예약이 없습니다.")
        else:
            for r in sel_res:
                tr = f"{r['res_time']} ~ {r.get('res_time_end','')}"
                is_mine = (st.session_state.logged_in and
                           r["user_phone"] == st.session_state.user_phone)
                hdr = f"🕐 {tr}  │  {r['department']} {r['name']}"
                if is_mine: hdr += "  ✏️"

                with st.expander(hdr, expanded=True):
                    st.write(f"**부서:** {r['department']}")
                    st.write(f"**이름:** {r['name']}")
                    st.write(f"**전화번호:** {r['phone']}")
                    st.write(f"**예약 시간:** {tr}")
                    st.write(f"**방문지:** {r['destination']}")

                    if is_mine:
                        ba, bb = st.columns(2)
                        if ba.button("✏️ 수정", key=f"edit_res_btn_{r['id']}"):
                            st.session_state.editing_res_id = (
                                None if st.session_state.editing_res_id == r["id"]
                                else r["id"]
                            )
                            st.rerun()
                        if bb.button("🗑 취소", key=f"del_res_{r['id']}"):
                            delete_reservation(r["id"])
                            st.success("예약이 취소되었습니다.")
                            st.rerun()

                    if st.session_state.editing_res_id == r["id"]:
                        st.markdown("---")
                        st.markdown("**예약 수정**")
                        with st.form(f"form_edit_res_{r['id']}"):
                            e_dept = st.selectbox("부서", DEPARTMENTS,
                                                  index=_dept_idx(r["department"]))
                            e_name = st.text_input("이름", value=r["name"])
                            e_date = st.date_input(
                                "날짜",
                                value=datetime.strptime(r["res_date"], "%Y-%m-%d").date(),
                            )
                            ec1, ec2 = st.columns(2)
                            with ec1:
                                e_ts = st.time_input("시작 시간",
                                                     value=_time_val(r["res_time"]),
                                                     step=1800)
                            with ec2:
                                e_te = st.time_input("종료 시간",
                                                     value=_time_val(r.get("res_time_end"),
                                                                     "18:00"),
                                                     step=1800)
                            e_dest = st.text_input("방문지", value=r["destination"])
                            sa, sb = st.columns(2)
                            if sa.form_submit_button("저장", type="primary",
                                                     use_container_width=True):
                                if e_name and e_dest:
                                    update_reservation(
                                        r["id"], e_dept, e_name, str(e_date),
                                        e_ts.strftime("%H:%M"),
                                        e_te.strftime("%H:%M"), e_dest,
                                    )
                                    st.session_state.editing_res_id = None
                                    st.success("수정되었습니다.")
                                    st.rerun()
                                else:
                                    st.warning("이름과 방문지를 입력해 주세요.")
                            if sb.form_submit_button("취소", use_container_width=True):
                                st.session_state.editing_res_id = None
                                st.rerun()

    # ── 예약 등록 ──────────────────────────────────────────────
    with col_right:
        st.markdown("#### ✏️ 예약 등록")
        if not st.session_state.logged_in:
            st.warning("예약하려면 왼쪽 사이드바에서 로그인하세요.")
        else:
            with st.form("form_reservation", clear_on_submit=True):
                f_dept = st.selectbox("부서", DEPARTMENTS,
                                      index=_dept_idx(st.session_state.user_department))
                f_name = st.text_input("이름", value=st.session_state.user_name)
                f_date = st.date_input("날짜", value=sel)
                tc1, tc2 = st.columns(2)
                with tc1:
                    f_ts = st.time_input("시작 시간",
                                         value=datetime.strptime("09:00","%H:%M").time(),
                                         step=1800)
                with tc2:
                    f_te = st.time_input("종료 시간",
                                         value=datetime.strptime("18:00","%H:%M").time(),
                                         step=1800)
                f_dest = st.text_input("방문지")
                if st.form_submit_button("예약 등록", type="primary",
                                          use_container_width=True):
                    if f_name and f_dest:
                        add_reservation(
                            st.session_state.user_phone,
                            f_dept, f_name, st.session_state.user_phone,
                            str(f_date),
                            f_ts.strftime("%H:%M"), f_te.strftime("%H:%M"),
                            f_dest,
                        )
                        st.success("예약이 등록되었습니다!")
                        st.rerun()
                    else:
                        st.warning("이름과 방문지를 입력해 주세요.")


# ════════════════════════════════════════════════════════════════
# 탭 2 : 주행 전 기록
# ════════════════════════════════════════════════════════════════
def tab_pre_drive():
    st.subheader("주행 전 기록")
    if not st.session_state.logged_in:
        st.warning("사이드바에서 로그인하세요."); return

    with st.form("form_pre", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            p_dept = st.selectbox("부서", DEPARTMENTS,
                                  index=_dept_idx(st.session_state.user_department))
            p_name = st.text_input("이름", value=st.session_state.user_name)
            p_date = st.date_input("주행 날짜", value=date.today())
        with c2:
            p_odo  = st.number_input("출발 시 계기판 거리 (km)",
                                     min_value=0.0, step=1.0, format="%.0f")
            p_dest = st.text_input("목적지")
            p_comp = st.text_input("동행인",
                                   placeholder="부서/이름 형식, 쉼표로 구분")
        if st.form_submit_button("주행 전 기록 저장", type="primary",
                                  use_container_width=True):
            if p_dest and p_odo > 0:
                add_pre_drive(
                    st.session_state.user_phone, p_dept, p_name,
                    st.session_state.user_phone, str(p_date), p_odo, p_comp, p_dest,
                )
                st.success("주행 전 기록이 저장되었습니다.")
                st.rerun()
            else:
                st.warning("목적지와 계기판 거리를 입력해 주세요.")

    pre = get_pre_drives(st.session_state.user_phone)
    if pre:
        st.divider()
        st.markdown("##### 완료 대기 중인 주행 기록")
        for p in pre:
            st.info(
                f"📅 **{p['drive_date']}**  │  "
                f"출발: **{p['odometer_start']:,.0f} km**  │  "
                f"목적지: {p['destination']}  │  동행: {p['companions'] or '없음'}"
            )


# ════════════════════════════════════════════════════════════════
# 탭 3 : 주행 후 기록
# ════════════════════════════════════════════════════════════════
def tab_post_drive():
    st.subheader("주행 후 기록")
    if not st.session_state.logged_in:
        st.warning("사이드바에서 로그인하세요."); return

    pre_drives = get_pre_drives(st.session_state.user_phone)
    if not pre_drives:
        st.info("완료 대기 중인 주행 전 기록이 없습니다.\n\n"
                "'주행 전 기록' 탭에서 먼저 기록하세요.")
        return

    options = {
        f"{p['drive_date']}  │  {p['odometer_start']:,.0f} km 출발  │  {p['destination']}": p
        for p in pre_drives
    }
    sel_label = st.selectbox("완료할 주행 기록 선택", list(options.keys()))
    sel_log   = options[sel_label]
    lid       = sel_log["id"]
    odo_s     = float(sel_log["odometer_start"] or 0)

    with st.container(border=True):
        st.markdown(
            f"**출발 날짜:** {sel_log['drive_date']}  \n"
            f"**출발 계기판:** {odo_s:,.0f} km  \n"
            f"**목적지:** {sel_log['destination']}  \n"
            f"**동행인:** {sel_log['companions'] or '없음'}"
        )

    st.markdown("---")
    c1, c2 = st.columns(2)
    with c1:
        q_date = st.date_input("도착 날짜", value=date.today(), key=f"pd_date_{lid}")
        q_odo_end = st.number_input(
            "도착 시 계기판 거리 (km)",
            min_value=odo_s, value=odo_s, step=1.0, format="%.0f",
            key=f"pd_odo_{lid}",
        )
        driven = q_odo_end - odo_s
        st.metric("주행 거리", f"{driven:,.0f} km",
                  help="도착 계기판 입력 시 자동 계산")
    with c2:
        q_park       = st.text_input("주차 장소", key=f"pd_park_{lid}")
        q_comp       = st.text_input("동행인",
                                     value=sel_log["companions"] or "",
                                     key=f"pd_comp_{lid}")
        q_charge_amt = st.number_input(
            "충전 금액 (원, 없으면 0)",
            min_value=0.0, step=100.0, format="%.0f",
            key=f"pd_charge_{lid}",
        )

    st.markdown("")
    if st.button("주행 후 기록 완료", type="primary",
                 use_container_width=True, key=f"pd_submit_{lid}"):
        if not q_park:
            st.warning("주차 장소를 입력해 주세요.")
        elif driven < 0:
            st.warning("도착 계기판이 출발 계기판보다 작습니다.")
        else:
            complete_drive(lid, q_odo_end, q_charge_amt,
                           q_park, q_comp, str(q_date))
            st.success("주행 기록이 완료 처리되었습니다!")
            st.rerun()


# ════════════════════════════════════════════════════════════════
# 탭 4 : 내 주행기록
# ════════════════════════════════════════════════════════════════
def _edit_log_form(log):
    lid    = log["id"]
    status = log["status"]
    odo_s  = float(log["odometer_start"] or 0)

    st.markdown("---")
    st.markdown("**주행기록 수정**")
    with st.form(f"form_edit_log_{lid}"):
        ec1, ec2 = st.columns(2)
        with ec1:
            e_date  = st.date_input(
                "날짜",
                value=datetime.strptime(log["drive_date"], "%Y-%m-%d").date(),
            )
            e_odo_s = st.number_input(
                "출발 계기판 (km)", min_value=0.0, value=odo_s,
                step=1.0, format="%.0f",
            )
            e_dest  = st.text_input("목적지", value=log["destination"] or "")
            e_comp  = st.text_input("동행인", value=log["companions"] or "")
        with ec2:
            e_odo_e = None
            if status == "complete":
                e_odo_e = st.number_input(
                    "도착 계기판 (km)", min_value=0.0,
                    value=float(log["odometer_end"] or odo_s),
                    step=1.0, format="%.0f",
                )
                driven = e_odo_e - e_odo_s
                st.metric("주행 거리", f"{driven:,.0f} km")
                e_charge = st.number_input(
                    "충전 금액 (원)",
                    min_value=0.0, value=float(log.get("charging_amount") or 0),
                    step=100.0, format="%.0f",
                )
                e_park = st.text_input("주차 장소",
                                       value=log["parking_location"] or "")
            else:
                st.info("미완료 기록: 출발 정보만 수정 가능합니다.")
                e_charge = 0.0
                e_park   = log["parking_location"] or ""

        sa, sb = st.columns(2)
        if sa.form_submit_button("저장", type="primary", use_container_width=True):
            odo_end = e_odo_e if status == "complete" else None
            update_drive_log(
                lid, str(e_date), e_odo_s, odo_end,
                e_comp, e_dest, e_charge, e_park, status,
            )
            st.session_state.editing_log_id = None
            st.success("수정되었습니다.")
            st.rerun()
        if sb.form_submit_button("취소", use_container_width=True):
            st.session_state.editing_log_id = None
            st.rerun()


def tab_my_logs():
    st.subheader("내 주행기록")
    if not st.session_state.logged_in:
        st.warning("사이드바에서 로그인하세요."); return

    all_logs = get_user_all_logs(st.session_state.user_phone)
    if not all_logs:
        st.info("등록된 주행 기록이 없습니다."); return

    complete = [l for l in all_logs if l["status"] == "complete"]
    pending  = [l for l in all_logs if l["status"] == "pre"]

    total_dist = sum((l["odometer_end"] or 0) - (l["odometer_start"] or 0)
                     for l in complete)
    total_chrg = sum(l.get("charging_amount") or 0 for l in complete)
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("완료된 운행",   f"{len(complete)} 건")
    m2.metric("대기 중인 운행",f"{len(pending)} 건")
    m3.metric("누적 주행거리", f"{total_dist:,.0f} km")
    m4.metric("누적 충전금액", f"{int(total_chrg):,} 원")
    st.divider()

    if complete:
        st.markdown("#### ✅ 완료된 운행 기록")
        for log in complete:
            d    = datetime.strptime(log["drive_date"], "%Y-%m-%d")
            dist = (log["odometer_end"] or 0) - (log["odometer_start"] or 0)
            chrg = log.get("charging_amount") or 0
            hdr  = (f"📅 {log['drive_date']} ({WEEKDAYS[d.weekday()]})  │  "
                    f"{dist:,.0f} km  │  {log['destination'] or '-'}")
            if st.session_state.editing_log_id == log["id"]:
                hdr += "  ✏️"

            with st.expander(hdr):
                ca, cb = st.columns(2)
                with ca:
                    st.write(f"**날짜:** {log['drive_date']} ({WEEKDAYS[d.weekday()]})")
                    st.write(f"**출발 계기판:** {log['odometer_start']:,.0f} km")
                    st.write(f"**도착 계기판:** {log['odometer_end']:,.0f} km")
                    st.write(f"**주행 거리:** {dist:,.0f} km")
                with cb:
                    st.write(f"**목적지:** {log['destination'] or '-'}")
                    st.write(f"**동행인:** {log['companions'] or '없음'}")
                    st.write(f"**주차 장소:** {log['parking_location'] or '-'}")
                    st.write("**충전 금액:** " + (f"{int(chrg):,} 원" if chrg else "-"))

                ba, bb = st.columns(2)
                if ba.button("✏️ 수정", key=f"edit_log_{log['id']}"):
                    st.session_state.editing_log_id = (
                        None if st.session_state.editing_log_id == log["id"]
                        else log["id"]
                    )
                    st.rerun()
                if bb.button("🗑 삭제", key=f"del_log_{log['id']}"):
                    st.session_state.confirm_del_log = log["id"]
                    st.rerun()

                if st.session_state.confirm_del_log == log["id"]:
                    st.warning("정말 삭제하시겠습니까?")
                    y, n = st.columns(2)
                    if y.button("삭제 확인", key=f"yes_log_{log['id']}", type="primary"):
                        delete_drive_log(log["id"])
                        st.session_state.confirm_del_log = None
                        st.rerun()
                    if n.button("취소", key=f"no_log_{log['id']}"):
                        st.session_state.confirm_del_log = None
                        st.rerun()

                if st.session_state.editing_log_id == log["id"]:
                    _edit_log_form(log)

    if pending:
        st.divider()
        st.markdown("#### ⏳ 주행 완료 대기 중")
        for log in pending:
            d   = datetime.strptime(log["drive_date"], "%Y-%m-%d")
            hdr = (f"📅 {log['drive_date']} ({WEEKDAYS[d.weekday()]})  │  "
                   f"출발 {log['odometer_start']:,.0f} km  │  "
                   f"{log['destination'] or '-'}")
            with st.expander(hdr):
                st.write(f"**날짜:** {log['drive_date']}")
                st.write(f"**출발 계기판:** {log['odometer_start']:,.0f} km")
                st.write(f"**목적지:** {log['destination'] or '-'}")
                st.write(f"**동행인:** {log['companions'] or '없음'}")
                st.caption("'주행 후 기록' 탭에서 완료 처리하세요.")

                ba, bb = st.columns(2)
                if ba.button("✏️ 수정", key=f"edit_pre_{log['id']}"):
                    st.session_state.editing_log_id = (
                        None if st.session_state.editing_log_id == log["id"]
                        else log["id"]
                    )
                    st.rerun()
                if bb.button("🗑 삭제", key=f"del_pre_{log['id']}"):
                    st.session_state.confirm_del_log = log["id"]
                    st.rerun()

                if st.session_state.confirm_del_log == log["id"]:
                    st.warning("정말 삭제하시겠습니까?")
                    y, n = st.columns(2)
                    if y.button("삭제 확인", key=f"yes_pre_{log['id']}", type="primary"):
                        delete_drive_log(log["id"])
                        st.session_state.confirm_del_log = None
                        st.rerun()
                    if n.button("취소", key=f"no_pre_{log['id']}"):
                        st.session_state.confirm_del_log = None
                        st.rerun()

                if st.session_state.editing_log_id == log["id"]:
                    _edit_log_form(log)


# ════════════════════════════════════════════════════════════════
# 관리자 화면
# ════════════════════════════════════════════════════════════════
def admin_panel():
    st.subheader("관리자 화면 — 운행기록 조회 및 내보내기")
    c1, c2 = st.columns(2)
    with c1:
        start_d = st.date_input("시작일", value=date.today().replace(day=1),
                                key="adm_start")
    with c2:
        end_d = st.date_input("종료일", value=date.today(), key="adm_end")

    if st.button("조회", type="primary"):
        st.session_state.adm_logs = get_logs_by_period(str(start_d), str(end_d))

    logs = st.session_state.get("adm_logs")
    if logs is None: return
    if not logs:
        st.warning("해당 기간에 완료된 운행 기록이 없습니다."); return

    rows = []
    for log in logs:
        d    = datetime.strptime(log["drive_date"], "%Y-%m-%d")
        dist = (log["odometer_end"] or 0) - (log["odometer_start"] or 0)
        chrg = log.get("charging_amount") or 0
        rows.append({
            "날짜":         f"{log['drive_date']} ({WEEKDAYS[d.weekday()]})",
            "부서":         log["department"],
            "이름":         log["name"],
            "출발 계기판":  f"{log['odometer_start']:,.0f}" if log["odometer_start"] else "-",
            "도착 계기판":  f"{log['odometer_end']:,.0f}"   if log["odometer_end"]   else "-",
            "주행거리(km)": f"{dist:,.0f}",
            "목적지(비고)": log["destination"] or "",
            "충전금액(원)": f"{int(chrg):,}" if chrg else "-",
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    td = sum((l["odometer_end"] or 0)-(l["odometer_start"] or 0) for l in logs)
    tc = sum(l.get("charging_amount") or 0 for l in logs)
    m1, m2, m3 = st.columns(3)
    m1.metric("총 운행 건수", f"{len(logs)} 건")
    m2.metric("총 주행 거리", f"{td:,.0f} km")
    m3.metric("총 충전 금액", f"{int(tc):,} 원")

    st.download_button(
        "📥 엑셀 운행기록부 다운로드",
        data=make_excel(logs, str(start_d), str(end_d)),
        file_name=f"운행기록부_{start_d}_{end_d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )


# ════════════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════════════
def main():
    init_db()
    init_session()
    handle_query_params()

    col_title, col_adm = st.columns([8, 1])
    with col_title:
        st.markdown(
            "<h2 style='margin:0;color:#1a3a5c'>🚗 공용차량 관리 시스템</h2>"
            f"<small style='color:#888'>차종: {VEHICLE_NAME} │ 번호: {VEHICLE_NUMBER}</small>",
            unsafe_allow_html=True,
        )
    with col_adm:
        lbl = "🔓 관리자" if st.session_state.admin_logged_in else "🔒 관리자"
        if st.button(lbl, key="btn_admin", use_container_width=True):
            if st.session_state.admin_logged_in:
                st.session_state.admin_logged_in  = False
                st.session_state.show_admin_modal = False
            else:
                st.session_state.show_admin_modal = True
            st.rerun()

    if st.session_state.show_admin_modal and not st.session_state.admin_logged_in:
        with st.container(border=True):
            st.markdown("#### 관리자 인증")
            pw = st.text_input("비밀번호", type="password", key="adm_pw")
            b1, b2 = st.columns(2)
            if b1.button("확인", key="adm_ok"):
                if pw == ADMIN_PASSWORD:
                    st.session_state.admin_logged_in  = True
                    st.session_state.show_admin_modal = False
                    st.rerun()
                else:
                    st.error("비밀번호가 틀렸습니다.")
            if b2.button("취소", key="adm_cancel"):
                st.session_state.show_admin_modal = False
                st.rerun()

    render_user_panel()
    st.divider()

    if st.session_state.admin_logged_in:
        tabs = st.tabs(["📅 예약하기", "🚀 주행 전 기록",
                        "🏁 주행 후 기록", "📋 내 주행기록", "🔐 관리자 화면"])
        with tabs[0]: tab_reservation()
        with tabs[1]: tab_pre_drive()
        with tabs[2]: tab_post_drive()
        with tabs[3]: tab_my_logs()
        with tabs[4]: admin_panel()
    else:
        tabs = st.tabs(["📅 예약하기", "🚀 주행 전 기록",
                        "🏁 주행 후 기록", "📋 내 주행기록"])
        with tabs[0]: tab_reservation()
        with tabs[1]: tab_pre_drive()
        with tabs[2]: tab_post_drive()
        with tabs[3]: tab_my_logs()


if __name__ == "__main__":
    main()
