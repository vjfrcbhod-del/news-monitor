# 엔젤로보틱스 CB/CPS IM — 작업노트 (내부용, 배포 금지)

> 배포용 IM에는 들어가지 않는 내부 진행 메모임.
> 기준일: 2026.08.21 / 대상 문서: `IM_엔젤로보틱스_CB_CPS_v4_20260821.docx`

---

## 1. 확인 및 보완 필요 항목

| 우선순위 | 구분 | 확인 필요 항목 |
|---|---|---|
| **높음** | 딜 조건 | 발행규모, 트랜치별 배분, 만기, 표면이율·YTM, Put/Call, 리픽싱 한도, 우선배당률, 청약수수료 |
| **높음** | 회계 판단 | **CPS 자본 분류 가능 여부에 대한 감사인(신한회계법인) 사전 의견.** 상장 전 RCPS가 부채로 계상되어 자본총계가 음(-)이 된 이력이 있어 발행조건 설계 전 확인 필수 |
| **높음** | 정관 확인 | 우선주 발행한도(발행주식총수 1/4 = 3,848,245주) 및 종류주식 발행 근거 조항 원문 |
| **높음** | 자금 배분 | 3개 사용목적별 구체적 금액 배분 및 집행 일정 |
| **높음** | 실적 전망 | 2026년 하반기 및 2027년 매출·손익 전망, 해외 수주 파이프라인 |
| 중간 | 과제 정보 | 브레인-투-로봇 과제 총 사업비, 정부출연금 및 민간부담금, 단계별 마일스톤 |
| 중간 | 해외 사업 | 아세안 3개국 지정대리점 계약 현황, 유럽 CE MDR 진행 단계 |
| 중간 | 연간 손익 | 2024·2025년 연간 매출원가 및 판관비 세부 항목 (사업보고서 원문) |
| 낮음 | 주가 | 2026.07.31 종가 20,150원 (증권정보 서비스 기준, 미검증) |

---

## 2. 산업분석 도표 — 삽입 위치 및 출처

본문 IV장에 빈 표로 자리를 잡아두었음. 아래 출처에서 이미지를 확보하여 교체.

| # | 본문 위치 | 도표 | 출처 |
|---|---|---|---|
| 1 | IV.1 글로벌 시장 규모 | 글로벌 재활로봇 시장 규모 추이 및 전망 | ⭐ **한국IR협의회 기업분석보고서 p.7~8** (사내 보유)<br>대안: [GII Korea](https://www.giikorea.co.kr/report/gmi1858972-rehabilitation-robots-market-opportunity-growth.html) |
| 2 | IV.2 국내 수요 기반 | 국내 고령인구 비율 추이 및 장래인구추계 | [국가데이터처 2025 고령자 통계](https://www.kostat.go.kr/board.es?mid=a10301010000&bid=10820&list_no=438832)<br>[KOSIS 100대 지표 — 고령인구](https://kosis.kr/visual/nsportalStats/detailContents.do?listId=A&statJipyoId=3634) |
| 3 | IV.2 국내 수요 기반 | 뇌졸중 환자수 및 재활 수요 추이 | [심평원 의료빅데이터 개방시스템](https://opendata.hira.or.kr/op/opc/selectMedInfoSvcList.do) |
| 4 | IV.3 수가 체계 및 정책 | 보행재활로봇 건강보험 수가 적용 구조 | ⭐ **한국IR협의회 보고서 p.6** (수가코드 도식)<br>[국립재활원 재활로봇중개연구사업](https://www.nrc.go.kr/research/html/content.do?depth=rp&menu_cd=03_04_00_01) |
| 5 | IV.4 경쟁 구도 | 국내 재활 웨어러블 로봇 시장점유율 (원그래프) | ⭐ **한국IR협의회 보고서 p.13** |
| 6 | VII.1 주가 추이 | 상장 이후 주가 추이 (코스닥 대비 상대주가) | [KRX 정보데이터시스템](http://data.krx.co.kr) |

**⭐ 표시 4건은 이미 보유하신 한국IR협의회 보고서(2024.09.26)에서 바로 캡처 가능.** 도표 품질이 가장 우수하므로 우선 사용 권장.

### 참고: 시장규모 수치 편차
조사기관별 편차가 매우 큼. 본문은 동사 분석보고서가 인용한 **GII·technavio 기준**(재활로봇 '23년 8.63억달러 → '28년 46.4억달러, CAGR 39.98%)을 사용.
- [Fortune Business Insights](https://www.fortunebusinessinsights.com/industry-reports/rehabilitation-robots-market-101013)
- [Mordor Intelligence](https://www.mordorintelligence.kr/industry-reports/rehabilitation-robots-market)

---

## 3. 문서 구조 메모

- Executive Summary: 거래 개요 → 투자 논거(보유 자산/변화 시점/하방 방어) → 핵심 지표 → Term Sheet → 자금용도
- Investment Highlights 4축(MECE): 시장 지위 / 성장 동력 / 인적 자원 / 하방 방어
- 핵심 논거 축: **CPS 자본확충 → 계속사업손실 요건(2027~) 방어 / 해외 매출 실행 → 매출액 요건(2029~) 방어**
- 재생성: `python3 out/build_im.py`
