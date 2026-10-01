# 주간 루틴

## 월 · 착수
1. 웹 Claude → 이번 주 과제 파일 생성 → `docs/weeks/WNN.md` 에 저장 (목표·A·B·완료 기준만)
2. 각자 브랜치 · `git checkout -b wNN-A` / `wNN-B`
3. 터미널 → `claude` → `/week A` (또는 `/week B`)

## 화~목 · 구현
- CC가 코드 작성 → VS2022 Rebuild → 오류 붙여넣기 → CC 수정 (반복)
- `scripts/deploy.ps1` → Inventor 실행 → 테스트 모델로 확인 → 스크린샷 저장 (`docs/shots/WNN/`)
- 결정·문제가 생기면 CC가 `decisions.md` `issues.md` 에 append (CLAUDE.md 규칙)

## 금 · 마감
1. `/wrap` → 변경 목록 · 기록 누락 확인 · 상태 한 줄
2. `git checkout main` → 두 브랜치 merge → `git tag WNN`
3. `scripts/snapshot.ps1 -Week NN -State "상태 한 줄"`
4. 웹 Claude로 · `decisions.md` `issues.md` 이번 주 항목 + 스크린샷 → 상세보고서 · Weekly report · PPT

## 두 사람
- GitHub private repo 하나 · 각자 clone
- 브랜치 `wNN-A` / `wNN-B` · 금요일 main merge
- `models/` 는 용량 문제로 git 제외 · 공유 드라이브로
