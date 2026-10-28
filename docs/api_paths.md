# API 접근 경로표 (Inventor 2027 Interop 기준)

| # | 목적 | 확인한 이름 | 타입 | 비고 |
|---|---|---|---|---|
| 1 | 활성 조립품 | | | |
| 2 | 순회(재귀) | ComponentOccurrence.SubOccurrences | ComponentOccurrencesEnumerator | |
| 3 | 순회(평탄) | ComponentOccurrences.AllLeafOccurrences(Optional Object) | ComponentOccurrencesEnumerator | 괄호 형태 속성 |
| 4 | 부품 위치 | ComponentOccurrence.Transformation | Inventor.Matrix | 단위·누적 여부는 실측 |
| 5 | 경계 상자 | ComponentOccurrence.RangeBox | Inventor.Box | MinPoint/MaxPoint 확인 대기 |
| 6 | 면 | ComponentOccurrence.SurfaceBodies | Inventor.SurfaceBodies | SurfaceBody.Faces 확인 대기 |
| 7 | 면 종류 | Face.SurfaceType | SurfaceTypeEnum | kCylinderSurface 확인, kPlaneSurface 대기 |
| 8 | 원통 축 | Cylinder.BasePoint | Inventor.Point | 단위·좌표계는 실측 대기 |
| 8 | 원통 축 | Cylinder.AxisVector | Inventor.UnitVector | |
| 8 | 원통 축 | Cylinder.Radius | Double | |
| 9 | 면 테두리 | Face.EdgeLoops / Face.Edges | Inventor.EdgeLoops / Inventor.Edges (get only) | EdgeLoops 는 IEnumerable · Item(Int32) As EdgeLoop · For Each 가능 (W07 · Interop 리플렉션으로 확인 · 실행 미검증) |
| 9 | 면 테두리 | EdgeLoop.IsOuterEdgeLoop | Boolean (get only) | 안쪽 루프 = False (W07 · Interop 리플렉션으로 확인 · 실행 미검증) |
| 9 | 면 테두리 | EdgeLoop.Edges | Inventor.Edges (get only) | Edges 는 IEnumerable · Item(Int32) As Edge · For Each 가능 (W07 · Interop 리플렉션으로 확인 · 실행 미검증) |
| 추가 | 면 | Face.CreatedByFeature / Face.ThreadInfos | PartFeature / ObjectCollection | 6주차 이후 재료 · CreatedByFeature 는 Inventor.PartFeature (get only) · Interop XML 설명: "FaceProxy 에서는 지원 안 함 · Nothing 반환" → 조립품 프록시 면은 NativeObject 로 파트 면을 얻은 뒤 호출해야 할 수 있음 · W06 실측은 프록시에서 값 반환 · 8-3 에서 재확인 (W08 · Interop 리플렉션·XML로 확인 · 실행 미검증) |
| 추가 | 면의 인접 면 | Edge.Faces | Inventor.Faces (get only) | Faces 는 IEnumerable · Item(Int32) As Face · Count Int32 · For Each 가능 (W08 · Interop 리플렉션으로 확인 · 실행 미검증) |
| 10 | 엣지 종류 | Edge.GeometryType | Inventor.CurveTypeEnum (get only) | kCircleCurve = 5124 · kCircularArcCurve = 5125 둘 다 있음 (W07 · Interop 리플렉션으로 확인 · 실행 미검증) |
| 10 | 엣지 종류 | Edge.Geometry | Object (get only) | Inventor.Circle 로 CType 필요 (W07 · Interop 리플렉션으로 확인 · 실행 미검증) |
| 11 | 원 정보 | Circle.Center | Inventor.Point (get/set) | cm · 조립품 좌표(프록시 · W04 결정) (W07 · Interop 리플렉션으로 확인 · 실행 미검증) |
| 11 | 원 정보 | Circle.Radius | Double (get/set) | cm (W07 · Interop 리플렉션으로 확인 · 실행 미검증) |
| 11 | 원 정보 | Circle.Normal | Inventor.UnitVector (get/set) | (W07 · Interop 리플렉션으로 확인 · 실행 미검증) |
| 추가 | 면 종류 | SurfaceTypeEnum.kPlaneSurface | SurfaceTypeEnum (값 5890) | 7행 "kPlaneSurface 대기" 확인 · kCylinderSurface = 5891 (W07 · Interop 리플렉션으로 확인 · 실행 미검증) |
| 12 | 부품 정의 | ComponentOccurrence.Definition | ComponentDefinition | PartComponentDefinition 변환 필요 |
| 추가 | 부품 피처 | PartComponentDefinition.Features | Inventor.PartFeatures (get only) | 13행 경로의 앞 단계 (W08 · Interop 리플렉션으로 확인 · 실행 미검증) |
| 13 | 구멍 피처 | PartFeatures.HoleFeatures | HoleFeatures | get only · HoleFeatures 는 IEnumerable · Item(Object) As HoleFeature · Count Int32 · For Each 가능 (W08 · Interop 리플렉션으로 확인 · 실행 미검증) |
| 14 | 탭 | HoleFeature.Tapped | Boolean | get/set (W08 · Interop 리플렉션으로 확인 · 실행 미검증) |
| 14 | 탭 | HoleFeature.TapInfo | Object (HoleTapInfo 또는 TaperedThreadInfo) | 변환 필요 · get/set · System.Object · HoleTapInfo · TaperedThreadInfo 타입 둘 다 존재 → TypeOf 확인 후 CType (W08 · Interop 리플렉션으로 확인 · 실행 미검증) |
| 14 | 탭 | HoleTapInfo.ThreadDesignation | String | 나사 호칭 · get only (W08 · Interop 리플렉션으로 확인 · 실행 미검증) |
| 15 | 클리어런스 | HoleFeature.IsClearanceHole | Boolean | |
| 15 | 클리어런스 | HoleFeature.ClearanceInfo | HoleClearanceInfo | 타입 명확 |
| 15 | 클리어런스 | HoleClearanceInfo.FastenerSize | String | |
| 15 | 클리어런스 | HoleClearanceInfo.FastenerType / FastenerStandard | String | |
| 15 | 클리어런스 | HoleClearanceInfo.FastenerFitType | FastenerFitType | |
| 추가 | 구멍 | HoleFeature.HoleDiameter / HoleType / HoleCenterPoints | Parameter / HoleTypeEnum / ObjectCollection | 8주차 재료 · HoleCenterPoints 는 Inventor.ObjectCollection (get/set) · 항목 타입은 리플렉션상 Object 로만 보임 (구체 타입 미확정 · SketchPoint · SketchPoint3D · WorkPoint 타입은 존재 · 8-4 실행에서 TypeName 확인) · 좌표는 파트 좌표 · cm (W08 · Interop 리플렉션으로 확인 · 실행 미검증) |
| 추가 | 컬렉션 | ObjectCollection.Item(Int32 Index) | Object (get only) | Count Int32 · IEnumerable · For Each 가능 · 시작 번호(1 또는 0)는 리플렉션·XML 설명으로 판단 불가 → For Each 로 순회하면 무관 · 인덱서를 쓰면 실행에서 확인 (Inventor 컬렉션 관례상 1부터 · 추정) (W08 · Interop 리플렉션으로 확인 · 실행 미검증) |
| 추가 | 좌표 변환 | Point.TransformBy(Matrix As Matrix) | Sub (반환 없음) | 점 자체를 바꿈 → 원본 보존 필요하면 Copy 후 호출 (W08 · Interop 리플렉션으로 확인 · 실행 미검증) |
| 추가 | 좌표 변환 | Matrix.Invert() | Sub (반환 없음 · 인자 없음) | 행렬 자체를 역행렬로 바꿈 → occ.Transformation 에 직접 호출 금지 · Copy 후 호출 (W08 · Interop 리플렉션으로 확인 · 실행 미검증) |
| 16 | 패턴·미러 | PartFeatures.RectangularPatternFeatures | RectangularPatternFeatures | |
| 16 | 패턴·미러 | PartFeatures.CircularPatternFeatures | CircularPatternFeatures | |
| 16 | 패턴·미러 | PartFeatures.MirrorFeatures | MirrorFeatures | |
| 추가 | 패턴 | PartFeatures.SketchDrivenPatternFeatures | SketchDrivenPatternFeatures | 계획서에 없음 |
| 추가 | 스케치 기반 패턴 부모 피처 | SketchDrivenPatternFeature.Definition.ParentFeatures | ObjectCollection | 피처 직접 속성 아님 · 실행 미검증 |
| 추가 | 직사각형 패턴 부모 피처 | RectangularPatternFeature.ParentFeatures | ObjectCollection | 피처 직접 속성 · 실행 검증 · 홀간 거리 확인-1 의 피처 경로에서 확인 |
| 추가 | 원형 패턴 부모 피처 | CircularPatternFeature.ParentFeatures | ObjectCollection | 피처 직접 속성 · 실행 미검증 · 빌드만 통과 · 원형 패턴 모델 없음 |
| 추가 | 미러 부모 피처 | MirrorFeature.ParentFeatures | ObjectCollection | 피처 직접 속성 · 실행 검증 · 홀간 거리 확인-1 의 피처 경로에서 확인 |
| 추가 | 압출컷 | PartFeatures.ExtrudeFeatures | ExtrudeFeatures | 피처 없는 구멍 확인용 |
| 17 | 활성 뷰 | Application.ActiveView | Inventor.View | W05 · Interop 리플렉션으로 확인 (W05 · 존재만 확인, 동작은 W10에서 확정) |
| 17 | 카메라 | View.Camera | Inventor.Camera | 읽기 전용 속성 (W05 · 존재만 확인, 동작은 W10에서 확정) |
| 17 | 화면 맞춤 | Camera.Fit() | Sub | View.Fit(Optional DoUpdate As Boolean)도 있음 (W05 · 존재만 확인, 동작은 W10에서 확정) |
| 17 | 목표점 지정 | Camera.Target | Inventor.Point (get/set) | cm · 조립품 좌표 (W05 · 존재만 확인, 동작은 W10에서 확정) |
| 17 | 적용 | Camera.Apply() / Camera.ApplyWithoutTransition() | Sub | Apply는 전환 애니메이션 있음(추정) (W05 · 존재만 확인, 동작은 W10에서 확정) |
| 17 | 시점 | Camera.Eye / Camera.UpVector | Point / UnitVector (get/set) | W05 범위 밖 · 이름만 확인 (W05 · 존재만 확인, 동작은 W10에서 확정) |
| 17 | 점 생성 | TransientGeometry.CreatePoint(Optional X, Y, Z As Double) | Inventor.Point | Application.TransientGeometry (W05 · 존재만 확인, 동작은 W10에서 확정) |
| 추가 | 색 표시 세트 | Document.HighlightSets (AssemblyDocument.HighlightSets) | Inventor.HighlightSets (get only) | W06 · Interop 리플렉션으로 확인 |
| 추가 | 색 표시 세트 | HighlightSets.Add() | Inventor.HighlightSet | 인자 없음 (W06) |
| 추가 | 색 표시 세트 | HighlightSet.Color | Inventor.Color (get/set) | W06 |
| 추가 | 색 표시 세트 | HighlightSet.AddItem(Entity As Object) | Sub | 조립품에서는 Face 프록시를 넣음 (W06) |
| 추가 | 색 표시 세트 | HighlightSet.Clear() / HighlightSet.Delete() | Sub | 이름만 확인 · 누적 정리는 W06 범위 밖 |
| 추가 | 색 생성 | Application.TransientObjects | Inventor.TransientObjects (get only) | W06 |
| 추가 | 색 생성 | TransientObjects.CreateColor(Red As Byte, Green As Byte, Blue As Byte, Optional Opacity As Double = 1) | Inventor.Color | W06 |
