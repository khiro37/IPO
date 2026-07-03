# 공모주 자동수집

`ipo_watch.py`는 네이버 금융 IPO 일정에서 작업일 다음날 `개인청약`이 시작되는 공모기업을 찾고,
DART 투자설명서에서 핵심 공모주 지표를 추출합니다.

## 실행

```bash
.venv/bin/python ipo/ipo_watch.py
```

특정 일자를 기준으로 테스트:

```bash
.venv/bin/python ipo/ipo_watch.py --date 2026-07-03 --dry-run
```

## 저장 파일

- `ipo/ipo_watch_results.csv`: 기존 공모주 투자 엑셀 양식에 맞춘 결과
- `ipo/ipo_watch_results.xlsx`: 기존 공모주 투자 엑셀 양식에 맞춘 결과
- `ipo/ipo_watch_raw.csv`: DART 링크와 추출 근거를 포함한 내부 확인용 원자료

기존 원자료가 있으면 `작업일 + 회사 + 개인청약_시작일 + 접수번호` 기준으로 병합합니다.

## 추출 항목

- 공모가
- 수요예측 경쟁률
- 의무확약비율(전)
- 시가총액
- 유통가능비율

결과 파일에는 기존 엑셀과 같은 컬럼만 저장합니다.
시가총액은 억원 단위 정수로 반올림합니다.

## DART API 키

`DART_API_KEY` 환경변수를 사용합니다.
GitHub Actions에서는 저장소의 `Settings > Secrets and variables > Actions > Repository secrets`에
`DART_API_KEY` 이름으로 등록하세요.
