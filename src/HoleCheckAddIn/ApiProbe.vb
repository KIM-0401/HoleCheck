Option Strict On
Option Explicit On

' W04-05 임시 진단 코드
' 임시 진단 버튼 "API Probe" · Inventor API 접근 경로를 실측해 바탕화면 텍스트 파일로 출력한다 (W04.md A 담당)
' 값을 찍어보는 용도 · 수집·매칭·필터링 로직은 만들지 않는다
' 출력은 API 원값 그대로 · mm 변환하지 않는다
' "(추정)" 주석이 붙은 멤버는 개체 브라우저로 확인하지 못한 이름 · 빌드 오류로 확인 필요
Public Module ApiProbe

    Private Const intMaxCylinderFacesPerLeaf As Integer = 3
    Private Const dblCompareTolerance As Double = 0.000001

    ' 부품 정의 하나의 요약 숫자 · 마지막 요약 줄에 쓴다
    Private Class DefinitionStats
        Public intPlacementCount As Integer
        Public strHoleFeatureCount As String = "-"
        Public intCylinderFaceCount As Integer
        Public intCircleEdgeCount As Integer
        Public intCircularInnerLoopCount As Integer
    End Class

    Public Sub Run(ByVal invAssemblyDocument As Inventor.AssemblyDocument)
        Dim objText As New System.Text.StringBuilder()
        objText.AppendLine("단위 미변환(API 원값)")
        objText.AppendLine("문서: " & invAssemblyDocument.FullFileName)
        objText.AppendLine("시각: " & System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        objText.AppendLine()

        Dim invOccurrences As Inventor.ComponentOccurrences = invAssemblyDocument.ComponentDefinition.Occurrences

        ' 경로 B · AllLeafOccurrences
        Dim colLeaves As New System.Collections.Generic.List(Of Inventor.ComponentOccurrence)()
        For Each invLeaf As Inventor.ComponentOccurrence In invOccurrences.AllLeafOccurrences()
            colLeaves.Add(invLeaf)
        Next
        objText.AppendLine("리프 수 (AllLeafOccurrences): " & colLeaves.Count.ToString())
        objText.AppendLine()

        ' 같은 부품 정의가 여러 번 배치되면 실측 3·4·5는 정의당 1회
        Dim colDefinitionKeys As New System.Collections.Generic.List(Of String)()
        Dim colLeafKeys As New System.Collections.Generic.List(Of String)()
        Dim colStats As New System.Collections.Generic.Dictionary(Of String, DefinitionStats)()
        Dim colFirstLeaf As New System.Collections.Generic.Dictionary(Of String, Inventor.ComponentOccurrence)()
        For Each invLeaf As Inventor.ComponentOccurrence In colLeaves
            Dim strKey As String = GetDefinitionKey(invLeaf)
            colLeafKeys.Add(strKey)
            If Not colStats.ContainsKey(strKey) Then
                colStats.Add(strKey, New DefinitionStats())
                colFirstLeaf.Add(strKey, invLeaf)
                colDefinitionKeys.Add(strKey)
            End If
            colStats(strKey).intPlacementCount += 1
        Next

        WriteUnitSection(objText, colLeaves)
        WriteTraversalSection(objText, invOccurrences, colLeaves)
        WriteFaceEdgeSection(objText, colDefinitionKeys, colFirstLeaf, colStats)
        WriteHoleFeatureSection(objText, colDefinitionKeys, colFirstLeaf, colStats)
        WritePatternSection(objText, colDefinitionKeys, colFirstLeaf, colStats)
        WriteProxySection(objText, colLeaves)
        WriteSummarySection(objText, colLeaves, colLeafKeys, colStats)

        Dim strDocumentName As String = System.IO.Path.GetFileNameWithoutExtension(invAssemblyDocument.DisplayName)
        For Each chrInvalid As Char In System.IO.Path.GetInvalidFileNameChars()
            strDocumentName = strDocumentName.Replace(chrInvalid, "_"c)
        Next
        Dim strFileName As String = "HoleCheck_Probe_" & strDocumentName & "_" & System.DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".txt"
        Dim strPath As String = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.DesktopDirectory), strFileName)

        ' BOM 포함 UTF-8 · 메모장·PowerShell에서 한글이 깨지지 않게
        System.IO.File.WriteAllText(strPath, objText.ToString(), New System.Text.UTF8Encoding(True))

        System.Windows.Forms.MessageBox.Show("저장됨" & System.Environment.NewLine & strPath, "API Probe")
    End Sub

    ' 실측 1 · 리프 RangeBox Min/Max 원값
    Private Sub WriteUnitSection(ByVal objText As System.Text.StringBuilder, ByVal colLeaves As System.Collections.Generic.List(Of Inventor.ComponentOccurrence))
        objText.AppendLine("==== 실측 1 · 단위 (리프 RangeBox 원값) ====")
        For Each invLeaf As Inventor.ComponentOccurrence In colLeaves
            objText.AppendLine(invLeaf.Name & " | " & ReadSafe(Function() FormatBox(invLeaf.RangeBox)))
        Next
        objText.AppendLine()
    End Sub

    ' 실측 2 · 경로 A(SubOccurrences 재귀 + 단계별 Transformation 누적) vs 경로 B(AllLeafOccurrences 리프 Transformation)
    Private Sub WriteTraversalSection(ByVal objText As System.Text.StringBuilder, ByVal invOccurrences As Inventor.ComponentOccurrences, ByVal colLeaves As System.Collections.Generic.List(Of Inventor.ComponentOccurrence))
        objText.AppendLine("==== 실측 2 · 순회 비교 (원점 변환 결과) ====")
        objText.AppendLine("A누적 = 루트부터 리프까지 각 단계 Transformation을 직접 곱함")
        objText.AppendLine("A리프 = SubOccurrences로 얻은 리프 자신의 Transformation만 적용")
        objText.AppendLine("B     = AllLeafOccurrences 리프의 Transformation만 적용")
        objText.AppendLine("A누적=B 이고 A리프=B 이면 SubOccurrences 프록시도 누적 변환을 이미 포함한다는 뜻")

        Dim colPathA As New System.Collections.Generic.Dictionary(Of String, Double())()
        Try
            WalkOccurrences(invOccurrences, "", New System.Collections.Generic.List(Of Inventor.Matrix)(), colPathA)
        Catch objException As System.Exception
            objText.AppendLine("경로 A 순회 실패: " & objException.Message)
        End Try

        Dim colSeenPaths As New System.Collections.Generic.HashSet(Of String)()
        For Each invLeaf As Inventor.ComponentOccurrence In colLeaves
            Dim strPath As String = GetOccurrencePath(invLeaf)
            colSeenPaths.Add(strPath)

            Dim dblB As Double() = Nothing
            Try
                dblB = TransformPoint(New Double() {0, 0, 0}, invLeaf.Transformation)
            Catch objException As System.Exception
                objText.AppendLine(strPath & " | B 읽기 실패: " & objException.Message)
                Continue For
            End Try

            If Not colPathA.ContainsKey(strPath) Then
                objText.AppendLine(strPath & " | A 없음 | B" & FormatTriple(dblB))
                Continue For
            End If

            Dim dblA As Double() = colPathA(strPath)
            Dim dblAccumulated As Double() = New Double() {dblA(0), dblA(1), dblA(2)}
            Dim dblLeafOnly As Double() = New Double() {dblA(3), dblA(4), dblA(5)}
            objText.AppendLine(strPath &
                               " | A누적" & FormatTriple(dblAccumulated) &
                               " | A리프" & FormatTriple(dblLeafOnly) &
                               " | B" & FormatTriple(dblB) &
                               " | 차이(A누적-B) 최대 " & FormatDouble(MaxAbsDifference(dblAccumulated, dblB)) &
                               " | 차이(A리프-B) 최대 " & FormatDouble(MaxAbsDifference(dblLeafOnly, dblB)))
        Next

        For Each strPath As String In colPathA.Keys
            If Not colSeenPaths.Contains(strPath) Then
                objText.AppendLine(strPath & " | B에 없음 (경로 A에만 있는 리프)")
            End If
        Next
        objText.AppendLine()
    End Sub

    ' 결과 · 경로 → {A누적 x,y,z, A리프 x,y,z}
    ' ComponentOccurrences·ComponentOccurrencesEnumerator를 IEnumerable로 넘긴다 (Interop 인터페이스가 IEnumerable 상속 · 추정)
    Private Sub WalkOccurrences(ByVal colOccurrences As System.Collections.IEnumerable, ByVal strParentPath As String, ByVal colParentMatrices As System.Collections.Generic.List(Of Inventor.Matrix), ByVal colResults As System.Collections.Generic.Dictionary(Of String, Double()))
        For Each invOccurrence As Inventor.ComponentOccurrence In colOccurrences
            Dim strPath As String = If(strParentPath = "", invOccurrence.Name, strParentPath & "/" & invOccurrence.Name)
            Dim colMatrices As New System.Collections.Generic.List(Of Inventor.Matrix)(colParentMatrices)
            colMatrices.Add(invOccurrence.Transformation)

            If invOccurrence.SubOccurrences.Count = 0 Then
                Dim dblAccumulated As Double() = New Double() {0, 0, 0}
                For intIndex As Integer = colMatrices.Count - 1 To 0 Step -1
                    dblAccumulated = TransformPoint(dblAccumulated, colMatrices(intIndex))
                Next
                Dim dblLeafOnly As Double() = TransformPoint(New Double() {0, 0, 0}, invOccurrence.Transformation)
                colResults(strPath) = New Double() {dblAccumulated(0), dblAccumulated(1), dblAccumulated(2), dblLeafOnly(0), dblLeafOnly(1), dblLeafOnly(2)}
            Else
                WalkOccurrences(invOccurrence.SubOccurrences, strPath, colMatrices, colResults)
            End If
        Next
    End Sub

    ' 실측 3 · 면·엣지 개수 (정의당 1회 · 첫 배치의 SurfaceBodies)
    Private Sub WriteFaceEdgeSection(ByVal objText As System.Text.StringBuilder, ByVal colDefinitionKeys As System.Collections.Generic.List(Of String), ByVal colFirstLeaf As System.Collections.Generic.Dictionary(Of String, Inventor.ComponentOccurrence), ByVal colStats As System.Collections.Generic.Dictionary(Of String, DefinitionStats))
        objText.AppendLine("==== 실측 3 · 면·엣지 개수 (정의당 1회) ====")
        objText.AppendLine("EdgeLoop 수와 원형 안쪽 루프 수는 평면 Face 기준 · Edge 수는 SurfaceBody.Edges 기준(면끼리 공유 엣지 중복 없음)")
        For Each strKey As String In colDefinitionKeys
            Dim invLeaf As Inventor.ComponentOccurrence = colFirstLeaf(strKey)
            Dim objStats As DefinitionStats = colStats(strKey)
            objText.AppendLine("[정의] " & strKey & " · 배치 " & objStats.intPlacementCount.ToString() & " · 대표 " & invLeaf.Name)

            Try
                Dim intBodyCount As Integer = 0
                Dim intFaceCount As Integer = 0
                Dim intPlaneFaceCount As Integer = 0
                Dim intCylinderFaceCount As Integer = 0
                Dim intOtherFaceCount As Integer = 0
                Dim intOuterLoopCount As Integer = 0
                Dim intInnerLoopCount As Integer = 0
                Dim intCircularInnerLoopCount As Integer = 0
                Dim intCircleEdgeCount As Integer = 0
                Dim intLineEdgeCount As Integer = 0
                Dim intOtherEdgeCount As Integer = 0

                For Each invBody As Inventor.SurfaceBody In invLeaf.SurfaceBodies
                    intBodyCount += 1

                    For Each invFace As Inventor.Face In invBody.Faces ' SurfaceBody.Faces (추정)
                        intFaceCount += 1
                        Select Case invFace.SurfaceType
                            Case Inventor.SurfaceTypeEnum.kPlaneSurface ' (추정)
                                intPlaneFaceCount += 1
                                For Each invLoop As Inventor.EdgeLoop In invFace.EdgeLoops
                                    If invLoop.IsOuterEdgeLoop Then ' (추정)
                                        intOuterLoopCount += 1
                                    Else
                                        intInnerLoopCount += 1
                                        If IsAllCircleLoop(invLoop) Then
                                            intCircularInnerLoopCount += 1
                                        End If
                                    End If
                                Next
                            Case Inventor.SurfaceTypeEnum.kCylinderSurface
                                intCylinderFaceCount += 1
                            Case Else
                                intOtherFaceCount += 1
                        End Select
                    Next

                    For Each invEdge As Inventor.Edge In invBody.Edges ' SurfaceBody.Edges (추정)
                        Select Case invEdge.GeometryType ' (추정)
                            Case Inventor.CurveTypeEnum.kCircleCurve ' (추정)
                                intCircleEdgeCount += 1
                            Case Inventor.CurveTypeEnum.kLineSegmentCurve, Inventor.CurveTypeEnum.kLineCurve ' (추정)
                                intLineEdgeCount += 1
                            Case Else
                                intOtherEdgeCount += 1
                        End Select
                    Next
                Next

                objStats.intCylinderFaceCount = intCylinderFaceCount
                objStats.intCircleEdgeCount = intCircleEdgeCount
                objStats.intCircularInnerLoopCount = intCircularInnerLoopCount

                objText.AppendLine("  SurfaceBody " & intBodyCount.ToString() & " · Face " & intFaceCount.ToString() &
                                   " (평면 " & intPlaneFaceCount.ToString() & " / 원통 " & intCylinderFaceCount.ToString() & " / 기타 " & intOtherFaceCount.ToString() & ")")
                objText.AppendLine("  평면 EdgeLoop 바깥 " & intOuterLoopCount.ToString() & " / 안쪽 " & intInnerLoopCount.ToString() &
                                   " · 모든 Edge가 원인 안쪽 루프 " & intCircularInnerLoopCount.ToString())
                objText.AppendLine("  Edge 원 " & intCircleEdgeCount.ToString() & " / 직선 " & intLineEdgeCount.ToString() & " / 기타 " & intOtherEdgeCount.ToString())
            Catch objException As System.Exception
                objText.AppendLine("  읽기 실패: " & objException.Message)
            End Try
        Next
        objText.AppendLine()
    End Sub

    Private Function IsAllCircleLoop(ByVal invLoop As Inventor.EdgeLoop) As Boolean
        Dim intEdgeCount As Integer = 0
        For Each invEdge As Inventor.Edge In invLoop.Edges ' EdgeLoop.Edges (추정)
            intEdgeCount += 1
            If invEdge.GeometryType <> Inventor.CurveTypeEnum.kCircleCurve Then
                Return False
            End If
        Next
        Return intEdgeCount > 0
    End Function

    ' 실측 4 · HoleFeature (정의당 1회 · PartComponentDefinition 변환 실패하면 건너뜀)
    Private Sub WriteHoleFeatureSection(ByVal objText As System.Text.StringBuilder, ByVal colDefinitionKeys As System.Collections.Generic.List(Of String), ByVal colFirstLeaf As System.Collections.Generic.Dictionary(Of String, Inventor.ComponentOccurrence), ByVal colStats As System.Collections.Generic.Dictionary(Of String, DefinitionStats))
        objText.AppendLine("==== 실측 4 · HoleFeature (정의당 1회) ====")
        For Each strKey As String In colDefinitionKeys
            Dim objStats As DefinitionStats = colStats(strKey)
            objText.AppendLine("[정의] " & strKey & " · 배치 " & objStats.intPlacementCount.ToString())

            Dim invPartDefinition As Inventor.PartComponentDefinition = TryCast(colFirstLeaf(strKey).Definition, Inventor.PartComponentDefinition)
            If invPartDefinition Is Nothing Then
                objText.AppendLine("  PartComponentDefinition 아님 · 건너뜀")
                Continue For
            End If

            Dim invFeatures As Inventor.PartFeatures = Nothing
            Try
                invFeatures = invPartDefinition.Features ' PartComponentDefinition.Features (추정)
            Catch objException As System.Exception
                objText.AppendLine("  Features 읽기 실패: " & objException.Message)
                Continue For
            End Try

            objStats.strHoleFeatureCount = ReadSafe(Function() invFeatures.HoleFeatures.Count.ToString())
            objText.AppendLine("  HoleFeatures " & objStats.strHoleFeatureCount & " · ExtrudeFeatures " & ReadSafe(Function() invFeatures.ExtrudeFeatures.Count.ToString()))

            Try
                For Each invHole As Inventor.HoleFeature In invFeatures.HoleFeatures
                    objText.AppendLine("  - " & ReadSafe(Function() invHole.Name))
                    objText.AppendLine("      Tapped: " & ReadSafe(Function() invHole.Tapped.ToString()))
                    objText.AppendLine("      TapInfo: " & ReadSafe(Function() DescribeTapInfo(invHole.TapInfo)))
                    objText.AppendLine("      IsClearanceHole: " & ReadSafe(Function() invHole.IsClearanceHole.ToString()))
                    objText.AppendLine("      FastenerSize: " & ReadSafe(Function() invHole.ClearanceInfo.FastenerSize))
                    objText.AppendLine("      HoleType: " & ReadSafe(Function() invHole.HoleType.ToString()))
                    objText.AppendLine("      HoleDiameter 원값: " & ReadSafe(Function() FormatDouble(CType(invHole.HoleDiameter.Value, Double))))
                Next
            Catch objException As System.Exception
                objText.AppendLine("  HoleFeatures 순회 실패: " & objException.Message)
            End Try
        Next
        objText.AppendLine()
    End Sub

    Private Function DescribeTapInfo(ByVal objTapInfo As Object) As String
        If objTapInfo Is Nothing Then
            Return "Nothing"
        End If
        If TypeOf objTapInfo Is Inventor.HoleTapInfo Then
            Return "HoleTapInfo · ThreadDesignation=" & DirectCast(objTapInfo, Inventor.HoleTapInfo).ThreadDesignation
        End If
        If TypeOf objTapInfo Is Inventor.TaperedThreadInfo Then
            ' TaperedThreadInfo 멤버는 개체 브라우저 목록에 없어 읽지 않는다
            Return "TaperedThreadInfo · ThreadDesignation 미확인"
        End If
        Return "알 수 없는 타입"
    End Function

    ' 실측 5 · 패턴·미러 피처 수 (정의당 1회)
    Private Sub WritePatternSection(ByVal objText As System.Text.StringBuilder, ByVal colDefinitionKeys As System.Collections.Generic.List(Of String), ByVal colFirstLeaf As System.Collections.Generic.Dictionary(Of String, Inventor.ComponentOccurrence), ByVal colStats As System.Collections.Generic.Dictionary(Of String, DefinitionStats))
        objText.AppendLine("==== 실측 5 · 패턴·미러 피처 수 (정의당 1회) ====")
        For Each strKey As String In colDefinitionKeys
            Dim strHeader As String = "[정의] " & strKey & " · 배치 " & colStats(strKey).intPlacementCount.ToString()

            Dim invPartDefinition As Inventor.PartComponentDefinition = TryCast(colFirstLeaf(strKey).Definition, Inventor.PartComponentDefinition)
            If invPartDefinition Is Nothing Then
                objText.AppendLine(strHeader & " | PartComponentDefinition 아님 · 건너뜀")
                Continue For
            End If

            objText.AppendLine(strHeader &
                               " | 직사각형 " & ReadSafe(Function() invPartDefinition.Features.RectangularPatternFeatures.Count.ToString()) &
                               " | 원형 " & ReadSafe(Function() invPartDefinition.Features.CircularPatternFeatures.Count.ToString()) &
                               " | 스케치 기반 " & ReadSafe(Function() invPartDefinition.Features.SketchDrivenPatternFeatures.Count.ToString()) &
                               " | 미러 " & ReadSafe(Function() invPartDefinition.Features.MirrorFeatures.Count.ToString()))
        Next
        objText.AppendLine()
    End Sub

    ' 실측 6 · 프록시 원통면 BasePoint vs 정의 쪽 BasePoint × Transformation (리프당 원통면 3개까지)
    Private Sub WriteProxySection(ByVal objText As System.Text.StringBuilder, ByVal colLeaves As System.Collections.Generic.List(Of Inventor.ComponentOccurrence))
        objText.AppendLine("==== 실측 6 · 프록시 좌표계 ====")
        For Each invLeaf As Inventor.ComponentOccurrence In colLeaves
            objText.AppendLine("[리프] " & invLeaf.Name)
            Try
                Dim intFoundCount As Integer = 0
                For Each invBody As Inventor.SurfaceBody In invLeaf.SurfaceBodies
                    For Each invFace As Inventor.Face In invBody.Faces ' SurfaceBody.Faces (추정)
                        If invFace.SurfaceType <> Inventor.SurfaceTypeEnum.kCylinderSurface Then
                            Continue For
                        End If
                        intFoundCount += 1
                        objText.AppendLine("  원통면 " & intFoundCount.ToString() & " | " & ReadSafe(Function() DescribeProxyCylinder(invFace, invLeaf.Transformation)))
                        If intFoundCount >= intMaxCylinderFacesPerLeaf Then
                            Exit For
                        End If
                    Next
                    If intFoundCount >= intMaxCylinderFacesPerLeaf Then
                        Exit For
                    End If
                Next
                If intFoundCount = 0 Then
                    objText.AppendLine("  원통면 없음")
                End If
            Catch objException As System.Exception
                objText.AppendLine("  읽기 실패: " & objException.Message)
            End Try
        Next
        objText.AppendLine()
    End Sub

    Private Function DescribeProxyCylinder(ByVal invFace As Inventor.Face, ByVal invTransformation As Inventor.Matrix) As String
        Dim invProxyCylinder As Inventor.Cylinder = CType(invFace.Geometry, Inventor.Cylinder)
        Dim dblProxy As Double() = PointToArray(invProxyCylinder.BasePoint)

        Dim invFaceProxy As Inventor.FaceProxy = TryCast(invFace, Inventor.FaceProxy)
        If invFaceProxy Is Nothing Then
            Return "프록시" & FormatTriple(dblProxy) & " · FaceProxy 아님 · 판정 불가"
        End If

        Dim invNativeFace As Inventor.Face = CType(invFaceProxy.NativeObject, Inventor.Face) ' FaceProxy.NativeObject (추정)
        Dim invNativeCylinder As Inventor.Cylinder = CType(invNativeFace.Geometry, Inventor.Cylinder)
        Dim dblNative As Double() = PointToArray(invNativeCylinder.BasePoint)
        Dim dblNativeTransformed As Double() = TransformPoint(dblNative, invTransformation)

        Dim bolSameAsTransformed As Boolean = MaxAbsDifference(dblProxy, dblNativeTransformed) < dblCompareTolerance
        Dim bolSameAsNative As Boolean = MaxAbsDifference(dblProxy, dblNative) < dblCompareTolerance

        Dim strVerdict As String
        If bolSameAsTransformed AndAlso bolSameAsNative Then
            strVerdict = "판정 불가 (Transformation이 이 점을 옮기지 않음)"
        ElseIf bolSameAsTransformed Then
            strVerdict = "프록시 Geometry는 조립품 좌표"
        ElseIf bolSameAsNative Then
            strVerdict = "프록시 Geometry는 부품 좌표"
        Else
            strVerdict = "판정 불가 (어느 쪽과도 다름)"
        End If

        Return "프록시" & FormatTriple(dblProxy) & " | 정의×T" & FormatTriple(dblNativeTransformed) & " | 정의 원값" & FormatTriple(dblNative) & " → " & strVerdict
    End Function

    ' 마지막 요약 · 리프마다 한 줄
    Private Sub WriteSummarySection(ByVal objText As System.Text.StringBuilder, ByVal colLeaves As System.Collections.Generic.List(Of Inventor.ComponentOccurrence), ByVal colLeafKeys As System.Collections.Generic.List(Of String), ByVal colStats As System.Collections.Generic.Dictionary(Of String, DefinitionStats))
        objText.AppendLine("==== 요약 (리프마다 · 정의 기준 숫자) ====")
        objText.AppendLine("부품명 | HoleFeature 수 | 원통면 수 | 원형 edge 수 | 원형 안쪽 루프 수")
        For intIndex As Integer = 0 To colLeaves.Count - 1
            Dim objStats As DefinitionStats = colStats(colLeafKeys(intIndex))
            objText.AppendLine(colLeaves(intIndex).Name &
                               " | " & objStats.strHoleFeatureCount &
                               " | " & objStats.intCylinderFaceCount.ToString() &
                               " | " & objStats.intCircleEdgeCount.ToString() &
                               " | " & objStats.intCircularInnerLoopCount.ToString())
        Next
    End Sub

    ' 정의 식별 · 문서 전체 이름 (모델 상태 포함) · 실패하면 Occurrence 이름
    Private Function GetDefinitionKey(ByVal invLeaf As Inventor.ComponentOccurrence) As String
        Try
            Return CType(invLeaf.Definition.Document, Inventor.Document).FullDocumentName ' ComponentDefinition.Document (추정)
        Catch
            Return invLeaf.Name
        End Try
    End Function

    ' 경로 A의 이름 경로와 맞추기 위한 "상위/…/리프" 문자열 · 실패하면 Occurrence 이름
    Private Function GetOccurrencePath(ByVal invLeaf As Inventor.ComponentOccurrence) As String
        Try
            Dim colNames As New System.Collections.Generic.List(Of String)()
            For Each invStep As Inventor.ComponentOccurrence In invLeaf.OccurrencePath ' ComponentOccurrence.OccurrencePath (추정)
                colNames.Add(invStep.Name)
            Next
            Return String.Join("/", colNames)
        Catch
            Return invLeaf.Name
        End Try
    End Function

    ' 점 × 4x4 행렬 · Matrix.Cell(행, 열) 1부터 (추정)
    Private Function TransformPoint(ByVal dblPoint As Double(), ByVal invMatrix As Inventor.Matrix) As Double()
        Dim dblResult(2) As Double
        For intRow As Integer = 1 To 3
            dblResult(intRow - 1) = invMatrix.Cell(intRow, 1) * dblPoint(0) +
                                    invMatrix.Cell(intRow, 2) * dblPoint(1) +
                                    invMatrix.Cell(intRow, 3) * dblPoint(2) +
                                    invMatrix.Cell(intRow, 4)
        Next
        Return dblResult
    End Function

    Private Function PointToArray(ByVal invPoint As Inventor.Point) As Double()
        Return New Double() {invPoint.X, invPoint.Y, invPoint.Z}
    End Function

    Private Function MaxAbsDifference(ByVal dblFirst As Double(), ByVal dblSecond As Double()) As Double
        Dim dblMax As Double = 0
        For intIndex As Integer = 0 To 2
            dblMax = System.Math.Max(dblMax, System.Math.Abs(dblFirst(intIndex) - dblSecond(intIndex)))
        Next
        Return dblMax
    End Function

    Private Function FormatBox(ByVal invBox As Inventor.Box) As String
        ' Box.MinPoint / MaxPoint (추정)
        Return "Min" & FormatTriple(PointToArray(invBox.MinPoint)) & " Max" & FormatTriple(PointToArray(invBox.MaxPoint))
    End Function

    Private Function FormatTriple(ByVal dblValues As Double()) As String
        Return "(" & FormatDouble(dblValues(0)) & ", " & FormatDouble(dblValues(1)) & ", " & FormatDouble(dblValues(2)) & ")"
    End Function

    Private Function FormatDouble(ByVal dblValue As Double) As String
        Return dblValue.ToString("R", System.Globalization.CultureInfo.InvariantCulture)
    End Function

    ' 속성 하나 읽기 · 실패하면 "읽기 실패: 메시지"를 돌려주고 계속
    Private Function ReadSafe(ByVal objReader As System.Func(Of String)) As String
        Try
            Return objReader()
        Catch objException As System.Exception
            Return "읽기 실패: " & objException.Message
        End Try
    End Function

End Module
