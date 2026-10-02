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
- 이유: nested_top 실측(36개 리프)에서 직접 곱한 값이 AllLeafOccurrences 값과 약 100~118 어긋났고 리프 단독 값은 일치(차이 0). 리프 Transformation이 누적 변환을 이미 포함하므로 직접 곱하면 이중 적용됨(추정)
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
