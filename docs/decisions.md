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
