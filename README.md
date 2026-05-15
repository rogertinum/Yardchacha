# 공용차량 관리 시스템 (YardChacha)

공용차량(EV3, 05하 7211) 예약 및 운행기록 관리 Streamlit 앱

## 주요 기능

| 탭 | 기능 |
|---|---|
| 📅 예약하기 | 달력 조회, 예약 등록/수정/취소, 시간 중복 방지 |
| 🚀 주행 전 기록 | 출발 계기판, 목적지, 출발 시간 기록 |
| 🏁 주행 후 기록 | 도착 계기판, 도착 시간, 충전 금액, 주차 장소 기록 |
| 📋 내 주행기록 | 본인 운행기록 조회/수정/삭제 |
| 🔐 관리자 (비밀번호 필요) | 기간별 운행기록 조회, 엑셀 내보내기 |

## 예약 규칙
- 동일 날짜/시간대 중복 예약 불가 (시간 겹침 자동 검사)
- 인접한 예약(전 예약 종료 = 신규 예약 시작)은 허용

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
| charging_amount | 충전 금액 (원) |
| parking_location | 주차 장소 |
| status | pre(주행 전) / complete(완료) |
| created_at | 입력 시간 (관리자만 확인) |

## 배포
- Streamlit Community Cloud
- GitHub push 시 자동 배포

## 의존성
```
streamlit>=1.32.0
openpyxl>=3.1.2
pandas>=2.0.0
```

## 주의사항
- `vehicle_management.db` 파일은 .gitignore에 포함 (로컬 보관)
- 엑셀 내보내기 양식은 세금 신고용이므로 형식 변경 금지
- 관리자 초기 비밀번호: 1111
