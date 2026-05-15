# 공용차량 관리 시스템 (YardChacha)

공용차량 예약 및 운행기록 관리 Streamlit 앱

## 주요 기능

| 탭 | 기능 |
|---|---|
| 📅 예약하기 | JS 달력 조회(날짜 클릭 → 자동 반영), 예약 등록/수정/취소, 시간 중복 방지 |
| 🚀 주행 전 기록 | 출발 계기판, 목적지, 출발 시간 기록, 오늘 예약 자동 채우기 |
| 🏁 주행 후 기록 | 도착 계기판, 도착 시간, 충전 금액, 주차 장소, 주행 거리 자동 계산 |
| 📋 내 주행기록 | 본인 운행기록 조회/수정/삭제 |
| 🔐 관리자 | 기간별 운행기록 조회, 엑셀 내보내기, 전체 예약 수정/삭제 |

## 예약 규칙

- 동일 날짜/시간대 중복 예약 불가 (시간 겹침 자동 검사)
- 인접한 예약(전 예약 종료 = 신규 예약 시작)은 허용
- 같은 날짜 주행 전 기록 중복 등록 불가

## 등록/수정/삭제 확인

- 모든 저장·완료 처리 전 팝업(`@st.dialog`)으로 내용 확인 후 진행
- 삭제 시 팝업에서 내용 확인 후 삭제 확인

## DB 스키마

### reservations (예약)

| 컬럼 | 설명 |
|---|---|
| res_date | 예약 날짜 (YYYY-MM-DD) |
| res_time / res_time_end | 시작/종료 시간 (HH:MM) |
| destination | 방문지 |
| purpose | 방문 목적 |
| created_at | 입력 시간 (관리자만 확인) |

### driving_logs (운행 기록)

| 컬럼 | 설명 |
|---|---|
| drive_date | 운행 날짜 |
| odometer_start / odometer_end | 출발/도착 계기판 (km) |
| depart_time / arrive_time | 출발/도착 시간 (HH:MM) |
| companions | 동행인 |
| destination | 목적지 |
| purpose | 방문 목적 |
| charging_amount | 충전 금액 (원) |
| parking_location | 주차 장소 |
| status | pre(주행 전) / complete(완료) |
| created_at | 입력 시간 (관리자만 확인) |

## 파일 구조

```
app.py                        # 메인 앱 (단일 파일)
calendar_component/
  index.html                  # JS 달력 컴포넌트 (Streamlit Custom Component)
requirements.txt
.gitignore
```

## 배포

- Streamlit Community Cloud
- GitHub push 시 자동 배포

## 의존성

```
streamlit>=1.36.0
openpyxl>=3.1.2
pandas>=2.0.0
psycopg[binary]>=3.1.0
```

## 주의사항

- 엑셀 내보내기 양식은 세금 신고용이므로 형식 변경 금지
