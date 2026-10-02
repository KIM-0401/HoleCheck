# 겪은 문제 누적

(CLAUDE.md 형식으로 append · "왜 예측 못 했나" 필수)

## W04 · src 아래 VS 프로젝트가 두 개 생김 (HoleCheck · HoleCheckAddIn)
- 증상: src/HoleCheck/(17:06 생성)와 src/HoleCheckAddIn/(18:10 생성)이 둘 다 있다. 둘 다 Class1.vb만 들어 있고 sln도 각각 따로 있다. W04 기준은 src\HoleCheckAddIn 단일 프로젝트
- 원인: (추정 · 확인 필요) 처음에 프로젝트 이름을 HoleCheck로 만들었고, 이름을 바꾸지 않고 새로 만든 뒤 이전 폴더를 지우지 않음
- 처리: src/HoleCheck 폴더를 사람이 삭제함 (W04 마감 다음 날 확인). src/HoleCheckAddIn 하나만 남음
- 왜 예측 못 했나: VS 새 프로젝트 대화상자에서 프로젝트 이름이 곧 폴더명·AssemblyName·RootNamespace가 된다는 점을 생성 전에 확인하지 않았다. 과제 파일도 경로만 적고 프로젝트 이름은 명시하지 않았다

## W04 · 생성된 vbproj 설정이 CLAUDE.md 규칙과 다름
- 증상: HoleCheckAddIn.vbproj가 OptionStrict=Off이고 기본 Imports(System, Microsoft.VisualBasic, System.Linq 등 9개)가 켜져 있다. Inventor Interop 참조도 아직 없다
- 원인: VS 클래스 라이브러리 템플릿 기본값 그대로임. 프로젝트 속성 > 컴파일 / 참조 단계를 아직 진행하지 않음
- 처리: 미처리. W04 B 항목(Option Strict=On, Interop 참조 Embed=False/CopyLocal=False)과 CLAUDE.md(기본 Imports 끔)에 맞춰 VS 속성 창에서 사람이 바꿔야 함
- 왜 예측 못 했나: 템플릿 기본값이 Strict Off이고 Imports가 켜진 상태라는 걸 몰랐다. 프로젝트를 만든 시점에 CLAUDE.md 기술 스택 항목과 대조하는 체크 단계가 없었다

## W04 · .addin이 bin\Debug에 복사되지 않음
- 증상: vbproj에 HoleCheckAddIn.addin(CopyToOutputDirectory=PreserveNewest)을 추가하고 빌드했는데 bin\Debug에 .addin이 없었다. deploy.ps1이 *.addin을 복사하지 못한다
- 원인: (추정) 프로젝트 파일을 바꾼 뒤 증분 빌드가 새 복사 단계를 반영하지 않았다
- 처리: 솔루션 다시 빌드(Rebuild Solution) 후 bin\Debug에 복사됨
- 왜 예측 못 했나: "빌드 성공" 메시지만 보고 출력 폴더 내용은 확인하지 않았다

## W04 · PowerShell에서 한글이 깨져 보임
- 증상: PowerShell에서 .vb/.addin/.md 파일을 출력하면 한글이 깨져 보인다
- 원인: BOM 없는 UTF-8 파일을 Windows PowerShell 5.1이 기본 코드페이지(CP949)로 읽는다. 파일 자체는 정상이다(Get-Content -Encoding UTF8로 확인)
- 처리: 동작에는 영향이 없어 그대로 둔다
- 왜 예측 못 했나: 화면 표시 문제와 파일 손상을 구분하지 못했다

## W04 · 원통면 수가 구멍 수와 다름
- 증상: 원통면 수가 구멍 수와 다르다
- 원인: 카운터보어는 원통면 2개, 카운터싱크는 원뿔(기타) 면, 관통홀은 위·아래 원형 edge 2개
- 처리: 이번 주는 기록만
- 왜 예측 못 했나: 구멍 1개를 면 1개로 가정했음

## W04 · HoleFeature 수가 실제 구멍 수보다 적음
- 증상: HoleFeature 수가 실제 구멍 수보다 적다
- 원인: 패턴·미러로 복제된 홀은 피처 수에 안 잡힘
- 처리: 기록만
- 왜 예측 못 했나: 패턴 안의 홀이 별도 HoleFeature로 나오는지 확인하지 않았음

## W04 · 탭 홀에서 FastenerSize·HoleDiameter 읽기 실패
- 증상: 탭 홀에서 FastenerSize·HoleDiameter 읽기가 실패한다
- 원인: 탭 홀에는 ClearanceInfo와 HoleDiameter가 비어 있음
- 처리: 분기해서 읽도록 설계
- 왜 예측 못 했나: 계획서가 탭/클리어런스를 같은 방식으로 읽는다고 가정했음

## W04 · 카운터싱크는 원통면이 아니라 원뿔 면이 생김
- 증상: 카운터싱크 홀에는 원통면이 아니라 원뿔 면이 생긴다
- 원인: 5주차 축 기반 매칭은 원통면만 수집하므로 놓칠 수 있음
- 처리: 5주차 한계로 기록
- 왜 예측 못 했나: 홀 종류별 면 구성을 사전에 확인하지 않았음

## W04 · 재귀로 부모 변환을 직접 곱하면 값이 틀어짐
- 증상: SubOccurrences 재귀로 부모 변환을 직접 곱하면 값이 틀어진다
- 원인: 리프 Transformation이 누적 변환을 포함
- 처리: AllLeafOccurrences 값을 그대로 사용
- 왜 예측 못 했나: 리프 프록시가 누적 변환을 포함하는지 코드 작성 전에 확인하지 않았음

## W04 · 탭 구멍 크기가 호칭 지름과 다르게 보임(M6 → 4.917 등)
- 증상: 탭 구멍 크기가 호칭 지름과 다르게 보인다(M6 → 4.917 등)
- 원인: TapInfo의 골지름 계열 값일 가능성(추정)
- 처리: 12주차에 정리
- 왜 예측 못 했나: 탭 홀 지름의 기준 값을 정의하지 않았음

## W04 · 작업 지시가 가리킨 과제 파일 W04-05.md가 없음
- 증상: B 1·2단계와 A 작업 지시가 모두 docs/weeks/W04-05.md를 먼저 읽으라고 했지만 docs/weeks에는 W03.md·W04.md만 있다. ApiProbe.vb 상단 주석도 "W04-05 임시 진단 코드"로 남았다
- 원인: (추정 · 확인 필요) 작업 지시에 4·5주차 통합 과제 파일 이름을 미리 적었고 실제 파일은 만들지 않음
- 처리: 매번 W04.md 기준으로 진행. 파일명·주석은 그대로 둠
- 왜 예측 못 했나: 작업 지시를 쓸 때 docs/weeks의 실제 파일명과 대조하지 않았다

## W04 · api_paths.md가 api_paths.md.txt로 저장돼 있었음
- 증상: A 작업 시점에 docs/api_paths.md가 없고 docs/api_paths.md.txt만 있었다. 지시의 "있으면 읽기" 조건에 걸리지 않을 뻔했다. 마감 시점에는 docs/api_paths.md로 바뀌어 커밋됨
- 원인: (추정) 확장자가 숨겨진 탐색기에서 텍스트 파일 이름을 바꿔 .txt가 남음
- 처리: 사람이 이름을 바꿈 (커밋 ffe5a14 기준 api_paths.md)
- 왜 예측 못 했나: 탐색기에서 보이는 파일명만 확인하고 실제 확장자는 확인하지 않았다

## W04 · 탭도 클리어런스도 아닌 일반 드릴 홀도 FastenerSize 읽기가 E_FAIL
- 증상: probe_block-1에서 FastenerSize E_FAIL이 11건인데 탭 홀은 10개다. 나머지 1건은 파트-1 구멍6(Tapped False · IsClearanceHole False · HoleDiameter 2.1)이다
- 원인: (추정) 클리어런스 홀이 아니면 ClearanceInfo가 비어 있음. 탭 여부와는 무관
- 처리: 미처리. decisions.md "Hole 피처 속성은 탭이면 TapInfo, 클리어런스면 ClearanceInfo"는 두 갈래만 다룬다. 일반 드릴 홀 갈래(HoleDiameter만 읽기)를 다음 주에 추가 검토
- 왜 예측 못 했나: 실측 결과를 정리할 때 탭/클리어런스 두 종류만 있다고 보고, 오류 건수와 탭 홀 수를 대조하지 않았다

## W04 · (해결 확인) vbproj 설정이 CLAUDE.md 규칙과 다름
- 증상: 위 "W04 · 생성된 vbproj 설정이 CLAUDE.md 규칙과 다름" 항목이 "처리: 미처리"로 남아 있었다
- 원인: 설정을 고친 뒤 issues.md에 처리 결과를 다시 적지 않음
- 처리: 커밋 47ef685에서 해결됨 (OptionStrict Off→On · 기본 Imports 9개 제거 · Inventor Interop 참조 Embed=False/Private=False 추가)
- 왜 예측 못 했나: issues.md가 append 전용이라 처리 결과를 원래 항목에 반영하는 절차가 없었다
