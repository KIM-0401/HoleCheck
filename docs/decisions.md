# 결정 사항 누적

(CLAUDE.md 형식으로 append)

## W04 · 검사 대상 범위 정의
- 결정: 직사각형 블록(판재) 형태 부품에 Hole 피처 또는 압출컷으로 가공된 원형 홀을 대상으로, 2장 이상 겹쳐 볼트·핀으로 체결되는 조립품의 홀 짝을 검사한다. 항목은 위치(중심 거리) · 지름 · 종류 정합(탭↔클리어런스, 핀↔핀홀). 제외: STEP/IGES 변환 모델 · 구매품 · 장공·비원형 홀 · 경사 축 홀(관찰만)
- 대안: 범용 설비 조립품 전체
- 이유: 받은 모델이 이 형태이고 홀 간 거리를 수치로 확인할 수 있어 기대값을 정량으로 만들 수 있음
- 영향 파일: 없음

## W04 · 4주차는 단일 프로젝트로 시작
- 결정: src/HoleCheckAddIn 단일 Class Library
- 대안: 처음부터 2계층(Add-in 등록 / 로직 라이브러리)
- 이유: 2계층 분리는 15주차 범위 · 지금 나누면 로드 문제의 원인 후보가 늘어남
- 영향 파일: 없음

## W04 · W04 과제 범위 재합의 (과제 파일 개정)
- 결정: A는 "API 조사"에서 "임시 진단 버튼 API Probe로 실측 5항목을 텍스트 파일로 출력 + 모델별 개수 대조표"로 바꾼다. B는 "단일 프로젝트 + 매니페스트 + Assembly 리본(Tab HoleCheck / Panel 검사 / 버튼 Hole Check) + 빈 폼"으로 정한다. deploy.ps1 검증, iLogic 집계, 구멍 1개 기하 정리, "부품 수·구멍 수 출력" 완료 기준은 보류(대조표로 대체)
- 대안: 원래 W04(API 조사 정리 + iLogic 집계 + deploy.ps1 검증)
- 이유: 조사 문서보다 실측 숫자가 W05 매칭 설계의 기대값 근거가 됨 · 2인 1주 분량에 맞춤
- 영향 파일: docs/weeks/W04.md

## W04 · Add-in GUID 확정
- 결정: Add-in 클래스 GUID = 475D83A4-5AE7-414F-ADED-296ED976490A. 들어간 위치 3곳 · (1) src/HoleCheckAddIn/EntryPoints.vb StandardAddInServer의 GuidAttribute (2) src/HoleCheckAddIn/HoleCheckAddIn.addin의 ClassId (3) 같은 파일의 ClientId (2·3은 중괄호 형식). 진입 클래스 이름은 StandardAddInServer이고 네임스페이스는 RootNamespace(HoleCheckAddIn)를 그대로 쓴다. AssemblyInfo.vb의 Assembly Guid(a1b2e484-…, typelib ID)와 Assembly ComVisible(False)는 그대로 둔다. 클래스에 ComVisible(True)를 붙여 덮어쓰므로 충돌하지 않는다
- 대안: AssemblyInfo의 Assembly Guid를 클래스 GUID로 같이 쓰기 · Assembly ComVisible(True)로 바꾸기
- 이유: 한 번 배포하면 Inventor가 ClientId로 Add-in을 식별하므로 고정값이 필요하다. typelib GUID와 클래스 GUID를 같은 값으로 쓰면 COM 등록 시 ID가 겹친다. Assembly 전체를 ComVisible로 열 필요는 없다
- 영향 파일: src/HoleCheckAddIn/EntryPoints.vb, src/HoleCheckAddIn/HoleCheckAddIn.addin

## W04 · API 실측을 Add-in 진단 버튼으로 수행
- 결정: 실측 6항목(단위 · 순회 A/B 비교 · 면·엣지 개수 · HoleFeature 속성 · 패턴·미러 수 · 프록시 좌표계)을 Add-in의 임시 버튼 "API Probe"(ApiProbe.Run)로 실행하고 결과를 바탕화면 HoleCheck_Probe_<문서이름>_<시각>.txt에 API 원값 그대로 저장한다. 같은 부품 정의를 여러 번 배치한 경우 실측 3·4·5는 정의당 1회만 하고 배치 수를 함께 적는다
- 대안: iLogic Rule로 집계
- 이유: 실제 Add-in이 실행되는 경로(Interop 참조 · Option Strict · 프록시 객체)와 같은 환경에서 값을 확인해야 W05 이후 코드에 그대로 옮길 수 있다. iLogic은 늦은 바인딩이라 타입 변환 문제가 드러나지 않는다
- 영향 파일: src/HoleCheckAddIn/ApiProbe.vb

## W04 · API 길이 단위는 cm이고 읽는 즉시 mm로 변환
- 결정: API에서 읽은 길이는 cm로 보고 읽는 즉시 ×10 하여 mm로 바꾼다
- 대안: 문서 기준 단위를 가정만 하고 쓰기
- 이유: 블록 긴 변 150.000mm 측정값과 Probe RangeBox 15가 일치, HoleDiameter 0.55·0.66·0.9가 M5·M6·M8 클리어런스 5.5·6.6·9.0mm와 일치
- 영향 파일: 없음

## W04 · 조립품 순회는 AllLeafOccurrences 사용, 리프의 Transformation은 부모 변환을 곱하지 않고 그대로 사용
- 결정: 조립품 순회는 AllLeafOccurrences를 쓰고, 리프의 Transformation에 부모 변환을 곱하지 않는다
- 대안: SubOccurrences 재귀로 단계별 Transformation을 직접 곱하기
- 이유: nested_top 실측(36개 리프)에서 직접 곱한 값이 AllLeafOccurrences 값과 102.315~129.036(API 원값 cm · 36개 리프 최소~최대) 어긋났고 리프 단독 값은 일치(차이 0). 리프 Transformation이 누적 변환을 이미 포함하므로 직접 곱하면 이중 적용됨(추정)
- 영향 파일: 없음

## W04 · 프록시 원통면 Geometry는 조립품 좌표로 취급하고 변환을 곱하지 않음
- 결정: 리프 Occurrence의 SurfaceBodies에서 얻은 원통면(프록시) Geometry는 조립품 좌표로 보고 Transformation을 곱하지 않는다
- 대안: 
- 이유: 블록-1(36개 리프)과 nested_top(이동+회전, 108건) 모두에서 조립품 좌표로 판정, 부품 좌표 0건
- 영향 파일: 없음

## W04 · Hole 피처 속성은 탭이면 TapInfo, 클리어런스면 ClearanceInfo로 갈라 읽는다
- 결정: 탭 홀은 TapInfo, 클리어런스 홀은 ClearanceInfo로 나눠 읽는다
- 대안: 
- 이유: 탭 홀은 FastenerSize 읽기에 E_FAIL, HoleDiameter 읽기에 Object reference not set 오류
- 영향 파일: 없음

## W04 · Hole 피처 속성 읽기를 탭 / 클리어런스 / 그 외 드릴 홀 세 갈래로 보완
- 결정: 위 "Hole 피처 속성은 탭이면 TapInfo, 클리어런스면 ClearanceInfo로 갈라 읽는다"를 세 갈래로 보완한다. (1) 탭(Tapped True) → TapInfo (2) 클리어런스(IsClearanceHole True) → ClearanceInfo와 HoleDiameter (3) 그 외 드릴 홀(둘 다 False) → HoleDiameter만 읽고 ClearanceInfo는 읽지 않는다
- 대안: 탭 / 클리어런스 두 갈래만 유지
- 이유: probe_block-1에서 FastenerSize E_FAIL이 11건인데 탭 홀은 10개다. 나머지 1건은 파트-1 구멍6(Tapped False · IsClearanceHole False)이고, 이 홀의 HoleDiameter는 2.1로 읽혔다. 두 갈래만 두면 이런 홀에서 ClearanceInfo를 읽다가 실패한다
- 영향 파일: 없음 (W05 이후 홀 읽기 코드)

## W04 · 폼 프레임워크는 WinForms
- 결정: HoleCheckForm은 WinForms로 구현한다
- 대안: WPF
- 이유: 5주차에 필요한 결과 그리드(DataGridView)와 행 더블클릭 이동에 WinForms로 충분하다
- 영향 파일: src/HoleCheckAddIn/HoleCheckForm.vb

## W04 · 4·5주차 통합 진행
- 결정: 4주차와 5주차를 한 과제 파일(docs/weeks/W04-05.md)로 진행한다. 4주차 완료 기준은 관문으로 두고 태그를 남긴다 (관문 1 w04-done · API 실측 w04-probe-done · 관문 2 w05-done)
- 대안: 
- 이유: 5주차가 4주차 결과(폼·API 경로)에 직접 의존
- 영향 파일: docs/weeks/W04-05.md, docs/weeks/W04.md

## W05 · 매칭 공차 2종 정의
- 결정: 축간 공차 = 두 축선 사이 최단거리(mm), 각도 공차 = 축 방향 평행 허용각(도)
- 대안: 
- 이유: 축 기반 판정은 평행 여부가 먼저 필요
- 영향 파일: (예정) src/HoleCheckAddIn/AxisMatching.vb, src/HoleCheckAddIn/HoleCheckForm.vb

## W05 · 축간 거리는 평행 판정 후 점-직선 거리로 계산
- 결정: 두 축이 평행한지 먼저 판정하고, 평행하면 한 축 위의 점에서 다른 축선까지의 수직 거리(점-직선 거리)를 축간 거리로 쓴다
- 대안: skew line 일반 공식
- 이유: 구멍 짝은 평행이 기본인데 일반 공식은 평행에서 분모가 0이 되고, 거의 평행이면 먼 곳에서 교차해 거리 0이 나올 수 있음
- 영향 파일: (예정) src/HoleCheckAddIn/AxisMatching.vb

## W05 · 공차 2개는 후보 범위와 정렬 공차
- 결정: 축간 거리 ≤ 정렬 공차는 OK, 정렬 공차 초과이면서 후보 범위 이내는 위치 불일치, 후보 범위 초과는 짝 아님
- 대안: 공차 1개
- 이유: 1개면 M2(0.707mm)가 "짝 없음"이 되어 위치 NG로 보고되지 않음. 조건은 정렬 공차 < 0.707 < 후보 범위 < 최소 구멍 간격의 절반
- 영향 파일: src/HoleCheckAddIn/AxisMatching.vb (RunAxisMatching 매개변수 dblCandidateRange · dblAlignmentTolerance)

## W05 · 서로 다른 리프 occurrence끼리 전수 비교, 중복 제거 없음
- 결정: 원통면 축은 서로 다른 리프 occurrence끼리 전부 비교하고, 한 구멍에 여러 상대가 걸려도 중복을 제거하지 않는다
- 대안: 가장 가까운 상대 하나만 남기기
- 이유: 축 방식의 한계를 줄 수로 기록하기 위함
- 영향 파일: (예정) src/HoleCheckAddIn/AxisMatching.vb

## W05 · 결과 구조에는 숫자와 문자열만 담음
- 결정: AxisMatchResult에는 숫자와 문자열만 담고 Face 프록시 등 COM 객체는 담지 않는다
- 대안: Face 프록시 보관
- 이유: A·B 병렬 작업의 접점. 문서가 바뀌면 COM 참조가 무효화될 수 있음(추정)
- 영향 파일: src/HoleCheckAddIn/AxisMatching.vb

## W05 · 후보 범위 2.0mm 확정
- 결정: 후보 범위 시작값 2.0mm, 정렬 공차 0.1mm 유지
- 대안: 후보 범위를 더 넓히기
- 이유: Axis txt 실측에서 같은 occurrence 안 구멍 사이 최소 축간 거리가 10mm(파트-4:4 등)라 절반인 5mm가 상한. 2.0은 그 안이고, 위치 불일치 최대값은 1.5mm(파트-1:3 ↔ 파트-2:4)라 2.0 안에 모두 들어옴
- 영향 파일: 없음 (폼 입력 시작값)
