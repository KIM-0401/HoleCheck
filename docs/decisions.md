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

## W06 · 피처 필터는 수집과 짝 계산 사이에 끼우고 RunAxisMatching 시그니처는 유지
- 결정: CollectCylinderAxes에서 원통면마다 피처 판정(홀 / 홀 아님 / 피처 읽기 실패)을 문자열로 적고, 새 단계 SelectHoleAxes가 홀 면만 골라 MatchAxes에 넘긴다. MatchAxes와 RunAxisMatching 시그니처는 바꾸지 않아 폼(HoleCheckForm)은 수정하지 않는다. 진단 txt 이름은 HoleCheck_Feature_로 바꿔 W05 Axis txt와 나란히 대조한다
- 대안: 축 기반·피처 기반을 함께 돌리는 별도 함수 + 폼 선택 입력
- 이유: W06 목표가 짝짓기 로직을 고정하고 재료만 바꾸는 것. W05 Axis txt가 이미 대조 기준으로 남아 있어 두 경로를 동시에 유지할 필요가 없고, 폼 변경은 A 범위 밖
- 영향 파일: src/HoleCheckAddIn/AxisMatching.vb

## W06 · ParentFeatures는 패턴·미러 4종에서만 읽음
- 결정: ParentFeatures는 RectangularPatternFeature · CircularPatternFeature · MirrorFeature · SketchDrivenPatternFeature만 TypeOf로 변환해 읽고, 그 밖의 피처는 부모 없음으로 본다. CreatedByFeature가 깊이 0, 깊이 6에서 멈추고, 방문 키는 피처 이름(부품 문서 안에서 고유)으로 쓴다
- 대안: 리플렉션(InvokeMember)으로 모든 피처의 ParentFeatures 시도
- 이유: ParentFeatures는 PartFeature 공통 멤버가 아니고, Option Strict On에서는 늦은 바인딩을 쓸 수 없음. 피처 경로를 진단 txt에 남기므로 빠진 피처 종류가 있으면 실측에서 보임
- 영향 파일: src/HoleCheckAddIn/AxisMatching.vb (FindHoleAncestor · GetParentFeatures)

## W06 · 바깥 링 묶음의 축 일치 기준은 W05 동축 임계값 재사용
- 결정: 같은 리프 · 같은 HoleFeature 이름 · 평행(0.1도) · 축간 거리 0.001mm 이하이면 한 묶음으로 보고, 최소 반지름 면만 남긴다. 나머지는 "바깥 링 제외"로 표시한다
- 대안: 새 임계값 정의
- 이유: W05 진단(같은 occurrence 안 구멍 간격)에서 이미 쓰는 동축 기준(dblCoaxialDistanceThreshold)이고, 실측 최소 간격 10mm와 거리가 충분히 멂. 패턴·미러 복제는 HoleFeature 이름이 같아도 축이 달라 따로 묶인다
- 영향 파일: src/HoleCheckAddIn/AxisMatching.vb (SelectHoleAxes)

## W06 · 결과 줄 수는 137을 기준으로 삼음
- 결정: 홀간 거리 확인-1 결과 줄 기준을 137로 기록한다
- 대안: 보고서 기재값 148에 맞추기 위해 바깥 링 필터를 완화
- 이유: 감소 24줄이 바깥 링 제외 면 24개와 일치하고 한 구멍이 한 줄로 나옴. :1 배치와 :2 배치가 동일하게 줄어 필터가 일관됨. 148로 맞추려면 :2 배치를 걸러내지 않는 상태가 필요해 한 구멍이 두 줄이 됨. 148은 보고서 기재값이며 계산 조건이 문서에 없어 재현 불가
- 영향 파일: docs/weeks/W06.md

## W06 · 색 표시용 Face 프록시는 호출자가 넘긴 목록으로 받음
- 결정: RunAxisMatching에 오버로드를 하나 추가해 호출자가 만든 빈 List(Of Inventor.Face) 두 개(홀 면 · 바깥 링 면)를 채우고, 진단 txt 경로를 ByRef로 돌려준다. 기존 3인자 시그니처는 남겨 두고 새 오버로드를 부른다. Face 프록시는 내부 CylinderAxis.invFace에만 담고 AxisMatchResult에는 넣지 않는다. 홀 면 = 비교 투입 면(SelectHoleAxes 반환), 바깥 링 면 = strFeatureState가 "바깥 링 제외"인 면
- 대안: 모듈 수준 속성(마지막 실행의 면 목록)에 보관 · AxisMatchResult에 Face 추가
- 이유: W05 결정(결과 구조에는 숫자와 문자열만)을 지킴. 모듈 수준 보관은 실행이 끝난 뒤에도 COM 참조가 남고 문서가 바뀌면 무효가 될 수 있어(추정) 실행 한 번 동안만 쓰는 지역 목록이 안전함. 진단 txt는 RunAxisMatching 안에서 먼저 저장되므로 색 표시 결과 줄은 폼이 그 경로에 덧붙인다
- 영향 파일: src/HoleCheckAddIn/AxisMatching.vb (CylinderAxis · RunAxisMatching · CollectCylinderAxes · WriteDiagnostics), src/HoleCheckAddIn/HoleCheckForm.vb (OnRunClick · HighlightFaces · AddFacesToHighlightSet)

## W06 · 하이라이트 세트를 폼 필드로 보관
- 결정: invHoleSet · invOuterRingSet을 HighlightFaces의 지역 변수에서 HoleCheckForm의 Private 필드로 옮긴다. 실행할 때마다 새 세트로 덮어쓰기만 하고 이전 세트는 지우지 않는다. 같은 실행에서 세트 Count · 시도 건수 · 첫 아이템 TypeName을 진단 txt 끝에 적는다
- 대안: 지역 변수 유지 · 문서 참조와 함께 보관하고 이전 세트를 정리
- 이유: 지역 변수 수명 확인용 진단. 함수가 끝나 참조가 사라지면 색이 사라지는지 가르기 위함. 10주차에서 문서 참조와 함께 보관하는 방식으로 대체 예정
- 영향 파일: src/HoleCheckAddIn/HoleCheckForm.vb (필드 invHoleSet · invOuterRingSet · HighlightFaces · ReadHighlightSetCount · ReadFirstItemTypeName)
- 확인: 수명 연장 후 파트-1 계열에서 초록, 파트-6:1 에서 작은 원 초록 · 큰 원 주황 표시 확인. 진단 Count 293 / 24 가 시도 건수와 일치

## W06 · 검사 대상 조립품 선정
- 결정: Buffer Prep Skid.iam 을 검사 대상 후보 조립품으로 선정한다
  경로: C:\Users\localadmin\Downloads\autodesk_inventor_2022_samples\autodesk_inventor_2022_samples\Models\Tube & Pipe\Buffer Prep Skid\Buffer Prep Skid.iam
- 대안: 스캔 스크립트(find-sample-assemblies.ps1 · 참조 파일명 수 6~60 기준)에서 후보 88개가 나왔다.
  (Models 249개 중. 상위: Blower Frame · PC Chassis · Stapler Channel 등)
  대안 후보는 열어서 비교하지 않았다. 탈락 사유는 확인하지 않았다(미검토).
- 이유: v4 6주차 조건 중 "Inventor 2022 샘플" · "Inventor 원본(샘플 폴더)" 은 충족. 사진상 탱크 뚜껑 등에 원형 볼트 구멍이 보인다.
  v4 11주차의 상한 도달 · Region 비교 · 15~16주차 대형 설비 성능 과제가 큰 모델을 전제로 하므로 규모가 큰 모델이 어울린다고 판단했다(추정).
- 미확인(11주차 실행 전 Inventor 에서 확인 필요): 판재형 2장 이상 겹침 · 볼트 체결 구조 · 구멍 형상(원형/장공 비율) · 구매품 비율 · 실제 리프 부품 수
- 리스크: 3주차 코멘트("너무 복잡하지 않은 조립품")와 달리 규모가 큼(폴더 기준 .iam 97 · .ipt 312 · 130.5 MB).
  스크립트의 참조 수 35 는 최상위 파일에 적힌 이름 수로 보이며(추정) 실제 리프 수가 아니다.
  6주차 빌드에는 비교 상한이 없어 이번 주에는 실행하지 않는다.
- 영향 파일: 없음 (실행은 11주차)
- 비고: 샘플은 사용자 PC 다운로드 폴더에 있음 · git 에 올리지 않음.
  팀원과 공유하려면 폴더 전체를 통째로 전달(.ipj 와 상대 경로 유지). 2022 → 2027 열 때 변환 안내가 뜰 수 있으며 저장하지 않는다.

## W06 · 피처만으로는 부족 · 기하로 후보를 모으고 피처는 정체 확인용
- 결정: 다음 주부터 구멍 후보는 기하(면 · 루프)로 모으고, 피처(HoleFeature 조상)는 정체 확인용으로 쓴다. 필터로 걸러낸 정보(바깥 링 지름 등)는 버리지 않고 보관한다. 압출컷 구멍은 이번 주에 억지로 잡지 않고 한계로 기록한다. 시각 확인은 색 표시(HighlightSet)로 한다. 다음 주에 파일을 분리한다
- 대안: W06 피처 기반 필터 유지 · 이번 주에 압출컷을 피처 외 방법으로 추가 수집
- 이유: M7 압출컷은 HoleFeature 조상이 없어 W05 OK 2줄 → W06 0줄로 짝이 사라짐. 바깥 링 제외 24면의 지름(10 mm 11개 · 11 mm 6개 · 15 mm 6개 · 8 mm 1개)이 결과에 남지 않음. 색 표시로 홀 면 293 · 바깥 링 24를 화면에서 확인할 수 있었음. 압출컷 잡기 · 바깥 링 지름 보존 · EdgeLoop 수집 · 파일 분리는 W06.md 범위 밖
- 영향 파일: 없음 (이번 주 코드 변경 없음 · 7주차 이후)

## W07 · 파일 분리 단위와 배치
- 결정: 파일 분리 단위는 A = CollectFromPart · MatchingEngine · ResultModel · GeometryHelpers / B = Globals · HoleCheckForm · EntryPoints · Zoom. ApiProbe.vb 는 그대로 둔다(15주차 제거 예정). EntryPoints.vb(StandardAddInServer) 는 이미 분리돼 있어 이번 주 변경 없음. My Project 아래 파일도 건드리지 않음
- 대안: 한 파일 유지 · ApiProbe 까지 포함해 재배치
- 이유: issues "W06 · 코드가 한 파일(AxisMatching.vb)에 몰려 오류 위치 특정이 어려움"
- 영향 파일: src 전체

## W07 · 좌표는 변환을 곱하지 않고 cm→mm 만 · 계획서 문구와 다름
- 결정: 프록시 Geometry 는 이미 조립품 좌표 (W04 결정). Transformation 을 곱하지 않고 cm→mm 만 한다. 계획서 v4 의 "파트→조립품 변환" 문구와 다르며 nested_top 으로 검증한다
- 대안: 계획서대로 Transformation 곱하기
- 이유: W04 결정 실측 108건(nested_top 이동+회전 · 부품 좌표 0건) · 프록시에 변환을 다시 곱하면 이중 적용 위험
- 영향 파일: CollectFromPart.vb

## W07 · 회귀 기준본은 024033
- 결정: 7-2 이후 회귀 비교 기준은 HoleCheck_Feature_..._024033.txt (색 표시 포함 · 분리 전 마지막 실행)
- 대안: 014759 유지
- 이유: 색 표시 코드 이후의 출력 형식과 같아 줄 단위 diff 가 가능
- 영향 파일: docs/probe/

## W07 · 매칭 미연결 기간 폼 처리
- 결정: 7-6 부터 실행 버튼은 RunLoopCollection 만 부른다. 그리드는 열 정의를 남기고 행을 비운다. 결과 목록 · FillGrid · GetResultAt · 더블클릭 핸들러는 AxisMatchResult 대신 MatchPair 형식으로 바꿔 컴파일만 되게 두고(행이 없어 도달하지 않음), 더블클릭의 Zoom 호출은 MatchPair 에 목표점이 없어 뺐다(Zoom.vb 는 유지). 상태 라벨은 "리프 N · 홀 묶음 N · 지름 2개 이상 묶음 N · <Loop txt 파일명>". 색 표시 함수와 하이라이트 필드는 남기되 호출하지 않는다. 공차 입력 2개와 입력 검증은 그대로 둔다
- 대안: 그리드 · 핸들러 · 색 표시 함수 삭제 · 묶음 목록을 그리드에 표시
- 이유: W07.md 시작 전 합의(매칭 미연결 기간 그리드는 비우고 상태 라벨에 요약 · 상세는 진단 txt). W08 에서 MatchPair 로, 10주차 색 표시 정리에서 HighlightSet 으로 재연결할 자리를 남김. 공차 입력은 W08 Cap/Tol 로 이어짐
- 영향 파일: src/HoleCheckAddIn/HoleCheckForm.vb (colResults · OnRunClick · FillGrid · OnGridSelectionChanged · OnGridCellDoubleClick · GetResultAt)

## W07 · 피처 추적 함수 보관 위치
- 결정: FindHoleAncestor · GetParentFeatures · intMaxFeatureDepth 는 CollectFromPart.vb 에 그대로 두고 "W06 피처 추적 · W07 수집에서는 호출하지 않음 · W08 HoleFeature 연계에서 재사용" 주석을 붙인다. 이 둘만 쓰던 피처 상태 상수 4개 · ClassifyFaceFeature · SelectHoleAxes 는 축 경로와 함께 제거
- 대안: 별도 파일(FeatureTrace.vb 등)로 분리 · 삭제 후 W08 에서 git 이력에서 복원
- 이유: W07.md A 항목 "W06 피처 추적은 수집에서 빼고 보관 (W08 재사용)". 수집과 같은 파일에 두면 W08 에서 묶음과 HoleFeature 를 연결할 때 위치를 찾기 쉽고, 이번 주 파일 분리 단위(W07 decisions "파일 분리 단위와 배치")를 바꾸지 않음
- 영향 파일: src/HoleCheckAddIn/CollectFromPart.vb

## W07 · EdgeLoop 루프 판정과 Loop 진단 출력 규칙
- 결정: 안쪽 루프의 모든 edge 가 kCircleCurve 일 때만 채택하고, 원 정보는 첫 edge 의 Geometry 를 Inventor.Circle 로 읽는다. 원호와 다른 곡선이 함께 있으면 "원호 포함"으로 하나만 센다. HoleCluster.AddDiameter 는 지름을 오름차순으로 넣고 intCircleCount 를 같이 늘린다(같은 지름도 버리지 않음). 묶음 번호는 리프 안에서 0부터. Loop 진단 txt 는 Feature 진단과 같은 폴더 · 인코딩 · 이름 규칙(접두 Loop)이며 저장 알림 메시지박스는 띄우지 않고 폼 상태 라벨에 파일명을 적는다
- 대안: 루프의 모든 edge 원을 비교해 중심이 같은지 확인 · 탈락 사유 복수 집계 · 메시지박스 표시
- 이유: 7-5 실측에서 edge 2개 이상 채택 루프 0 · 지름 2개 이상 묶음 0 이라 첫 edge 로 충분했다. 수집만 하는 주간이라 판정에 쓰이지 않는 집계는 단순하게 둠. 실행마다 메시지박스가 둘 뜨지 않게 함
- 영향 파일: src/HoleCheckAddIn/CollectFromPart.vb (CollectHoleClusters · ReadCircleLoop · AddToCluster) · ResultModel.vb (HoleCluster.AddDiameter) · MatchingEngine.vb (WriteLoopDiagnostics)

## W07 · 리프 쌍별 최근접 묶음 거리는 진단 txt 에만 · 판정 없음
- 결정: 서로 다른 리프 쌍마다 묶음 중심 사이 3D 거리 최솟값 · 두 묶음 번호 · 지름 리스트를 거리 오름차순(같으면 리프 번호 순)으로 전부 쓴다. 필터 · 판정 · MatchPair 사용 없음. 묶음이 없는 리프가 낀 쌍은 줄 없이 개수만 적는다. 리프별 탈락 수는 수집 로직을 바꾸지 않고 리프 시작 전 카운터와의 차이로 기록한다
- 대안: 후보 범위로 걸러 짝 후보만 출력 · MatchPair 로 그리드 표시
- 이유: W07 은 매칭 · 판정을 연결하지 않는 주간(W07.md 목표). 전체 출력이라 W06 쌍별 비교 18쌍과 대조할 수 있었음(3 mm 미만 18줄 = W06 18쌍)
- 영향 파일: src/HoleCheckAddIn/MatchingEngine.vb (WriteLeafPairNearest · CompareLeafPairNearest · FormatDiameterList) · CollectFromPart.vb (CollectHoleClusters)

## W08 · 연계 1차 카운터 의미와 경로 기록 방식
- 결정: 연계 카운터는 묶음 단위로 새 묶음을 만든 첫 루프의 결과만 센다. intLinkProxySuccessCount · intLinkNativeSuccessCount 는 CreatedByFeature 를 어느 경로에서 얻었는지, intLinkFailCount 는 프록시 · NativeObject 모두 Nothing(옆벽 없음 포함)인 묶음 수. HoleFeature 를 찾은 1차 성공 / 실패 묶음 수는 HoleCluster.intLinkStage 로 진단에서 따로 센다. 경로와 CreatedByFeature Type 이름은 ReadCircleLoop 의 ByRef 인자로 호출자에게 넘기고, Type 이름별 묶음 수는 LoopCounters.colCreatedByFeatureTypeCounts(Dictionary(Of String, Integer))에 담는다. 시드 역추적의 첫 부모는 인덱서 대신 For Each 첫 항목으로 얻는다
- 대안: intLinkFailCount = 1차 실패(HoleFeature 못 찾음) 묶음 수 · 경로를 LoopCircleInfo 필드로 추가 · 루프 단위 집계
- 이유: 8-3 지시의 진단 카운터 3개(프록시 직접 · NativeObject 경유 · 둘 다 Nothing)와 LoopCounters 필드 3개를 맞춤. 연계 필드 5개 외의 진단 값을 HoleCluster 로 복사하지 않으려고 ByRef 로 넘김. ObjectCollection 인덱서 시작 번호가 미확인(api_paths.md)
- 영향 파일: src/HoleCheckAddIn/CollectFromPart.vb (CollectHoleClusters · ReadCircleLoop · LinkHoleFeatureBySideWall · ReadCreatedByFeature · FindHoleFeatureBySeed · CountClusterLink · AddToCluster) · ResultModel.vb (LoopCounters) · MatchingEngine.vb (WriteLoopDiagnostics · WriteLinkSection)

## W08 · 연계 2차와 교차 검증 계산 방식
- 결정: 연계 2차는 리프 루프 수집이 끝난 뒤 리프마다 한 번 HoleFeature 전체 HoleCenterPoints 를 조립품 좌표 mm 후보 목록(숫자 · 문자열만)으로 읽고, 1차 실패 묶음에만 가장 가까운 후보(≤ 0.02 mm)를 적용한다. 교차 검증은 같은 후보 목록으로 모든 묶음에 대해 센다. 점은 SketchPoint.Geometry3d · SketchPoint3D.Geometry · WorkPoint.Point · Point 로 TypeOf 분기하고 모르는 타입은 "기타:TypeName" 으로 개수만 센다. 점은 Copy() 후 TransformBy(리프 Transformation) 한 번. 탭은 HoleFeature 단위로 한 번 읽어 후보에 담는다. 1차 성공 묶음의 탭 읽기 실패 사유는 strLinkNote 에 남고, 묶음별 탭 목록 마지막 열(비고)에 출력한다. 묶음별 1차 CreatedByFeature Type 이름은 CollectHoleClusters 안 리프 단위 리스트로만 들고 HoleCluster 에 넣지 않는다
- 대안: 묶음마다 HoleFeatures 를 다시 순회 · 2차 결과를 1차 성공 묶음에도 덮어씀 · 1차 Type 이름을 HoleCluster 필드로 추가
- 이유: COM 호출을 리프당 한 번으로 줄이고 COM 객체를 함수 밖에 두지 않음(W05 결정 · W08 시작 전 합의). 연계 필드 5개 외의 진단 값은 모델에 넣지 않음(W08 decisions "연계 1차 카운터 의미와 경로 기록 방식")
- 영향 파일: src/HoleCheckAddIn/CollectFromPart.vb (CollectHoleClusters · LinkHoleFeatureBySideWall · ReadHoleTap · LinkHoleFeatureByCenter · ReadHoleCenterCandidates · ReadCenterPoint · FindNearestCandidate · HoleCenterCandidate) · Globals.vb (dblLinkCenterToleranceMillimeter) · ResultModel.vb (LoopCounters) · MatchingEngine.vb (WriteLinkSection · WriteLinkStage2Section · WriteCountDictionary)

## W08 · 그리디 배타 매칭의 동률 순서 · 후보 수 전달 · 진단 형식
- 결정: 정렬 키는 거리 하나이고, 같은 거리는 후보 생성 순서(리프 A · 리프 B · 묶음 A · 묶음 B 이중 루프 순)로 둔다. Cap 비교는 거리 ≤ dblCap 그대로(경계 보정 없음). 사용 집합은 A쪽 · B쪽 HashSet 두 개 · 키 "리프 번호:리프 안 묶음 번호". 후보 수는 Private RunMatchingWithCount 의 ByRef 로 받고, Public RunMatching(colClusters, dblCap, dblTol) 은 그 래퍼. 진단 매칭 섹션은 요약 4줄 · 리프 쌍별 확정/OK/NG(리프 번호 순) · 확정 줄 목록(거리 오름차순 · 거리와 중점 R 형식 · 연계 A/B · 탭 A/B)
- 대안: List.Sort 를 그대로 씀(불안정 정렬이라 동률 순서가 입력 순서와 무관) · LINQ OrderBy(기본 Imports 꺼짐 · System.Core 참조 의존) · 후보 수를 진단에서 다시 셈
- 이유: 동률 3키 정렬은 14주차라 키를 늘리지 않되, 실행마다 같은 결과가 나오게 생성 순서로만 묶음. 지시된 RunMatching 시그니처를 지키면서 진단에 후보 수를 쓰려고 래퍼로 나눔
- 영향 파일: src/HoleCheckAddIn/MatchingEngine.vb (RunLoopCollection · RunMatching · RunMatchingWithCount · CompareMatchCandidate · CreateMatchPair · FormatLinkStage · FormatTap · WriteLoopDiagnostics · WriteMatchingSection · CompareLeafPairKey · MatchCandidate) · HoleCheckForm.vb (OnRunClick)
