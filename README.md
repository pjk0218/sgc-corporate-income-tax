# SGC Arabia 법인세 계산 웹앱 (Corporate Income Tax Calculation)

연말 결산 때 1년에 한 번 쓰는 웹앱. 엑셀 자료를 넣으면 Corporate Income Tax와 Deferred Tax를 계산 과정과 함께 보여준다. 화면은 영어(팀 공유용).
서버 없이 브라우저에서만 동작하며, 자료는 팀 공유 폴더(또는 브라우저) 밖으로 나가지 않는다. 급여 웹앱(sgc-wps-payroll)과 같은 방식.

## 파일

| 파일 | 설명 |
|---|---|
| `index.html` | 앱 본체 (코드만, 회사 데이터 없음, 외부 라이브러리 없음) |
| `tests/verify.py` | 검증 스크립트 — 엑셀 워킹 파일을 앱에 넣고 엑셀 계산값과 줄별 비교, 공유 폴더 저장·History·동시 수정·Final 잠금 점검 |

## 넣는 자료 (Inputs 체크리스트)

| 자료 | 필수 | 쓰이는 곳 |
|---|---|---|
| Trial Balance (연간) | 필수 | 순이익, Entertainment, Tax & Dues, Prepaid Tax |
| 1.Provision Movement | 필수 | 충당금 가산/차감, 이연법인세자산 |
| 2.PPE- Depreciation | 필수 | 감가상각 차이, 고정자산 세무상 장부가 |
| 3.Normal Guarantee | 선택 | 보증수수료 부인액. 수수료율 = 시장요율이면 "Not applicable" 체크 |
| Key-in: Fixed assets as per the Financial statements | 필수 | Deferred Tax Note 1 |
| Key-in: Deferred tax of prior year as per audited FS | 필수 | 이연법인세 당기 변동 |

- 한 워크북에 시트가 다 들어 있어도 되고, 파일을 나눠 넣어도 된다. 시트는 이름이 아니라 표 머리글로 알아본다.
- `.xlsx`만 지원(옛 `.xls`는 엑셀에서 다른 이름으로 저장).
- 빠진 자료·확인이 필요한 사항은 Notices 탭에 표시한다. 앱이 값을 임의로 채우지 않는다.

## 저장과 공유 (연도별 보관)

- 오른쪽 위 **Connect shared folder**로 팀 공유 폴더(네트워크 폴더)를 한 번 연결하면, 같은 폴더를 연결한 사람 모두가 같은 자료를 본다. 크롬·엣지(PC) 필요.
- 폴더를 연결하지 않으면 그 브라우저에만 저장된다(공유 안 됨).
- 회계연도마다 따로 저장하고, **History** 탭에서 지난 연도를 언제든 열어 본다.
- 끝난 연도는 **Mark as final**로 잠근다(읽기 전용, 필요하면 Unlock).
- 두 사람이 동시에 고치면 나중 사람의 변경은 저장하지 않고 "Load the latest data"를 안내한다(덮어쓰기 방지).
- 전년도가 앱에 저장돼 있으면: 전기 이연법인세 금액 제안, PPE·충당금 기초 잔액과 전년 기말 비교 Notice.

```
<공유 폴더>/
  FY2025/data.json      입력 자료·키인 값·계정 매핑
  FY2025/source/        올린 엑셀 원본 사본
  FY2026/...
```

## 수동 조정 (Manual adjustments)

- 기본은 수동 조정 없이 올린 자료를 그대로 계산한다.
- 제공 자료에 수동 조정이 있을 때만 Inputs 탭의 "Manual adjustments"에 줄(CIT A1~A6, 기타 세무조정, 이연법인세 고정자산·충당금), 사유, 금액을 키인한다.
- 넣은 조정은 계산표에 별도 줄로 보이고 Notices에도 표시된다.

## 계산식

- Net profit = Income(002-002, 002-0043 Cr YTD) − Expense(002-004, 002-005, 002-0046, 002-0049 Dr YTD)
- Taxable profit = Net profit + 보증수수료 부인액 + Entertainment(01-32, 02-028) + Tax & Dues(01-28, 02-027) + Provisions made − Provisions utilized + (GL 감가상각 − ZATCA 감가상각)
- Income tax = Taxable profit × 20%, Due to ZATCA = Income tax − Prepaid tax(001-0025 당기 대변)
- Deferred tax: 고정자산 세무상 장부가 = PPE 차기이월 잔액 + 당기 취득 50% − 당기 처분대가 50%; Net = (충당금 기말 × 세율 − [−(세무상 − 재무제표)] × 세율) × 비사우디 지분율; 당기 변동 = Net − 전기 잔액
- 계정 코드는 Inputs 탭의 "Account mapping"에서 바꿀 수 있다.

## 검증 (2025 워킹 파일 기준)

```
python3 tests/verify.py <Corporate_Income_Tax_Calculation_2025.xlsx>
```

- 24개 항목 모두 엑셀과 일치. 단, 엑셀 Normal Guarantee 시트 합계 수식에 들어 있던 수동 조정(−0.01, +0.02, −0.03)은 앱에서 쓰지 않는다(수동 조정 없이 단순 합계).
  그래서 부인액·과세소득은 엑셀보다 0.03, 법인세는 0.006 SAR 크다. 소수 둘째 자리로는 법인세 16,482,381.69, Due to ZATCA 13,535,235.45로 엑셀과 같다.
- 공유 폴더 저장, 동료 화면에서 같은 값 확인, 동시 수정 충돌, 다른 연도 TB 거부, Final 잠금까지 자동 점검.

## 닐슨 휴리스틱 10원칙 반영

1. 시스템 상태 표시 — 체크리스트 진행률(6개 중 n개), 탭의 Notices 개수, 자료별 출처 파일·시트, 저장 위치와 저장 시각
2. 실제 업무 용어 — 엑셀 워킹과 같은 줄 이름·순서(Net Profit → Disallowed → Taxable → Income Tax → Due to ZATCA)
3. 사용자 통제 — 자료 제거·매핑 변경·초기화 뒤 Undo, 초기화·Final 전 확인 창, Final 잠금 해제
4. 일관성 — 모든 계산표가 Ref / Line / How it is calculated / SAR 네 열
5. 오류 예방 — 숫자 입력 검증, 자료가 빠지면 계산을 막고 필요한 항목 안내, 다른 연도 TB 거부, 끝난 연도 잠금, 동시 수정 덮어쓰기 방지
6. 기억보다 인식 — 넣어야 할 자료를 미리 화면에 나열, 줄마다 출처와 산식 표시
7. 유연성 — 드래그 앤 드롭, 여러 파일 한 번에, 전년도 값 제안, 데이터 파일 내려받기/가져오기
8. 간결한 디자인 — 탭 구분, 세부 계정은 접어 두기
9. 오류 복구 — Notice마다 무엇이 문제이고 무엇을 하면 되는지(What to do) 표시
10. 도움말 — Help 탭(사용 순서, 엑셀 양식, 계산식)

## 주의

- 회사 데이터(엑셀, data.json, FY 폴더)는 저장소에 올리지 않는다(.gitignore).
- 장부에 계상된 법인세·이연법인세와의 차이는 참고용으로만 표시한다. 계산은 올린 자료 기준.

## 변경 이력

- 2026-11-16 최초 작성 (2025 워킹 엑셀 기준, 엑셀 값과 일치 검증)
- 2026-11-16 수동 반올림 조정 제거, 팀 공유 폴더 저장·연도별 History·Final 잠금 추가, 파일명을 index.html로 변경
- 2026-11-16 수동 조정 키인(Manual adjustments) 추가
