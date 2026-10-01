# HoleCheck Capstone · CLAUDE.md

Inventor 조립품에서 마주 보는 두 부품의 구멍 짝(위치·지름·탭/클리어런스)을 검사하는 Add-in.
2인 팀 · 14주 · 매주 정해진 범위만 구현한다.

## 절대 규칙

1. **이번 주 범위 밖은 구현하지 않는다.** `docs/weeks/` 의 최신 파일이 이번 주 과제다. 거기 없는 기능·최적화·리팩터를 먼저 제안하지 않는다. 필요해 보이면 `docs/issues.md`에 한 줄 적고 넘어간다.
2. **아래 경로는 열지도 읽지도 않는다.** 이 프로젝트는 처음부터 만드는 것이다.
   - `D:\choiseungjun\inventor_addins\` 이하 전부
   - `*완성본*`, `*판단대장*`, `*DesignIntent*`, `*실측프로토콜*` 이름의 파일
3. **빌드는 사용자가 VS2022에서 한다.** `dotnet build` / `msbuild` 를 직접 실행하지 않는다 (COMReference 비호환). 코드를 쓰고 나면 "VS2022에서 Rebuild Solution 후 오류를 붙여넣어 달라"고 요청한다.
4. **파일을 수정하기 전에 현재 내용을 먼저 읽는다.** 추정으로 덮어쓰지 않는다.
5. **자체 메모리 파일(MEMORY.md 등)을 만들지 않는다.** 기록은 아래 두 파일에만.

## 기술 스택

- VB.NET · .NET Framework 4.8 · Class Library · Visual Studio 2022
- Autodesk Inventor 2027 · `Autodesk.Inventor.Interop` (Embed Interop Types=False · Copy Local=False)
- `Option Strict On` · `Option Explicit On` — 모든 .vb 파일 첫 줄
- **프로젝트 기본 Imports가 꺼져 있다.** `Exception`·`Environment`·`GC`·`StringComparison`·`vbCrLf` 등은 `System.` 으로 완전 수식한다. `Imports Inventor` 와 충돌하는 이름(`Environment`)은 특히 주의.
- 이벤트 배선은 `AddHandler` (디자이너 `Handles` 절 사용 안 함)
- 명명 · 헝가리안 접두 (`dbl` `str` `bol` `col` `inv` `int` `obj`) · 축약 금지

## 단위·좌표

- Inventor API 내부 단위는 **cm**. 읽는 즉시 ×10 하여 **내부는 전부 mm**.
- 파트 좌표 → `occ.Transformation` 곱 → 조립품 좌표. `AllLeafOccurrences`의 리프 프록시는 누적 변환을 이미 포함한다.

## 폴더

```
src/HoleCheckAddIn/      VS 프로젝트 (여기만 코드)
docs/weeks/WNN.md        이번 주 과제 (사용자가 매주 넣음)
docs/decisions.md        결정 사항 누적 · 아래 형식으로 append
docs/issues.md           겪은 문제 누적 · 아래 형식으로 append
models/                  테스트 모델 (.iam/.ipt) · 읽기만
scripts/deploy.ps1       빌드 산출 DLL → Inventor Addins 폴더 복사
scripts/snapshot.ps1     주 말미 zip 스냅샷
```

## 기록 규칙 (매 작업마다)

**결정이 생기면** `docs/decisions.md` 끝에 append:
```
## WNN · (한 줄 제목)
- 결정:
- 대안:
- 이유:
- 영향 파일:
```

**문제를 겪으면** `docs/issues.md` 끝에 append:
```
## WNN · (증상 한 줄)
- 증상:
- 원인:
- 처리:
- 왜 예측 못 했나:
```

"왜 예측 못 했나" 칸을 비워두지 않는다. 이게 주간 보고서의 재료다.

## 응답 형식

- 코드 변경 전에 **되돌리기 난이도**(🟢 파일 하나 / 🟡 여러 파일 / 🟠 구조) 한 줄.
- 변경한 파일과 함수를 목록으로. 설명은 짧게.
- 진행 가부를 되묻지 않는다. 범위 안이면 그냥 한다. 범위 밖이면 issues.md에 적고 멈춘다.

## 배포 (사용자 실행)

1. Inventor 완전 종료 (실행 중이면 DLL 잠김)
2. VS2022 Rebuild Solution
3. `scripts/deploy.ps1` → `%APPDATA%\Autodesk\Inventor 2027\Addins\HoleCheckAddIn\`
4. Inventor 실행 → Assembly 문서 → 리본 확인
