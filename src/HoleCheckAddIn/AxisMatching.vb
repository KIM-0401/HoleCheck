Option Strict On
Option Explicit On

' 원통면 축 기반 1차 매칭 · A·B 병렬 작업의 접점 파일
' 결과에는 숫자와 문자열만 담는다 (COM 객체 금지 · decisions.md "W05 · 결과 구조에는 숫자와 문자열만 담음")
' 구멍/축 외면 구분 · 지름 판정 · 중복 제거 · 배타 매칭은 하지 않는다 (W05 결정)
' W06 · 수집과 짝 계산 사이에 피처 필터를 둔다 · HoleFeature 조상이 있는 면만, 카운터보어 바깥 링을 뺀 뒤 짝 계산에 넣는다
'       짝 계산(평행 판정 · 점-직선 거리 · 후보 범위 · 정렬 공차)은 W05 그대로

' 결과 한 줄 · 길이는 전부 mm, 좌표는 조립품 좌표
' 그리드 바인딩(DataGridView)이 공개 속성만 읽으므로 자동 구현 속성으로 둔다
Public Class AxisMatchResult

    Public Const strErrorKindOk As String = "OK"
    Public Const strErrorKindPositionMismatch As String = "위치 불일치"

    Public Property strErrorKind As String = ""
    Public Property dblAxisDistance As Double
    Public Property dblDiameterA As Double
    Public Property dblDiameterB As Double
    Public Property strOccurrenceNameA As String = ""
    Public Property strOccurrenceNameB As String = ""
    Public Property dblTargetX As Double
    Public Property dblTargetY As Double
    Public Property dblTargetZ As Double

End Class

Public Module AxisMatching

    ' 두 축 방향이 이 각도 이내면 평행으로 본다 (반대 방향 포함)
    Private Const dblParallelToleranceDegrees As Double = 0.1

    ' API 길이 단위 cm → 내부 mm (W04 결정)
    Private Const dblCentimeterToMillimeter As Double = 10.0

    ' 진단 출력 · 같은 occurrence 안 구멍 간격에서 이 거리(mm) 이하는 동축으로 보고 제외
    Private Const dblCoaxialDistanceThreshold As Double = 0.001

    ' W06 · CreatedByFeature를 깊이 0으로 보고 ParentFeatures를 따라 올라가는 최대 깊이
    Private Const intMaxFeatureDepth As Integer = 6

    ' W06 · 원통면의 피처 판정 상태
    Private Const strFeatureStateHole As String = "홀"
    Private Const strFeatureStateNotHole As String = "홀 아님"
    Private Const strFeatureStateReadFailed As String = "피처 읽기 실패"
    Private Const strFeatureStateOuterRing As String = "바깥 링 제외"

    ' 원통면 하나의 축 · 숫자만 보관 (mm · 조립품 좌표)
    Private Class CylinderAxis
        Public intLeafIndex As Integer
        Public strOccurrenceName As String
        Public dblPointX As Double
        Public dblPointY As Double
        Public dblPointZ As Double
        Public dblDirectionX As Double
        Public dblDirectionY As Double
        Public dblDirectionZ As Double
        Public dblRadius As Double
        ' W06 · 피처 판정 결과 · 문자열만 보관
        Public strFeatureState As String = ""
        Public strHoleFeatureName As String = ""
        Public strFeaturePath As String = ""
        ' W06 · 색 표시용 Face 프록시 · 이 모듈 밖으로는 호출자가 넘긴 목록으로만 내보내고 AxisMatchResult에는 넣지 않는다
        Public invFace As Inventor.Face
    End Class

    ' 리프 하나의 수집 결과 · 진단 출력용
    Private Class LeafSummary
        Public strOccurrenceName As String
        Public strDefinitionName As String
        Public intCylinderCount As Integer
        Public intFailedFaceCount As Integer
        Public strLeafError As String = ""
    End Class

    ' 진단 출력용 카운터
    Private Class MatchCounters
        Public intCylinderCount As Integer
        Public intFailedFaceCount As Integer
        Public intComparisonCount As Integer
        Public intNotParallelCount As Integer
        Public intOutOfRangeCount As Integer
        Public intOkCount As Integer
        Public intPositionMismatchCount As Integer
        ' W06 · 피처 필터
        Public intHoleCount As Integer
        Public intNotHoleCount As Integer
        Public intFeatureReadFailedCount As Integer
        Public intOuterRingCount As Integer
    End Class

    ' dblCandidateRange · 후보 범위(mm) · 축간 거리가 이보다 크면 짝 아님
    ' dblAlignmentTolerance · 정렬 공차(mm) · 축간 거리가 이하이면 OK, 초과하면서 후보 범위 이내이면 위치 불일치
    Public Function RunAxisMatching(ByVal invAssemblyDocument As Inventor.AssemblyDocument, ByVal dblCandidateRange As Double, ByVal dblAlignmentTolerance As Double) As System.Collections.Generic.List(Of AxisMatchResult)
        Dim strDiagnosticPath As String = ""
        Return RunAxisMatching(invAssemblyDocument, dblCandidateRange, dblAlignmentTolerance,
                               New System.Collections.Generic.List(Of Inventor.Face)(),
                               New System.Collections.Generic.List(Of Inventor.Face)(),
                               strDiagnosticPath)
    End Function

    ' W06 · 색 표시용 오버로드 · 호출자가 만든 빈 목록 두 개를 채운다 (모듈 수준 보관 없음)
    ' colHoleFaces · 비교 투입 면 (홀 확정이면서 바깥 링이 아닌 면)
    ' colOuterRingFaces · 카운터보어 바깥 링으로 제외된 면
    ' strDiagnosticPath · 저장한 진단 txt 경로 · 호출자가 맨 아래에 줄을 덧붙일 때 쓴다
    Public Function RunAxisMatching(ByVal invAssemblyDocument As Inventor.AssemblyDocument, ByVal dblCandidateRange As Double, ByVal dblAlignmentTolerance As Double, ByVal colHoleFaces As System.Collections.Generic.List(Of Inventor.Face), ByVal colOuterRingFaces As System.Collections.Generic.List(Of Inventor.Face), ByRef strDiagnosticPath As String) As System.Collections.Generic.List(Of AxisMatchResult)
        Dim objCounters As New MatchCounters()
        Dim colLeafSummaries As New System.Collections.Generic.List(Of LeafSummary)()

        Dim colAxes As System.Collections.Generic.List(Of CylinderAxis) = CollectCylinderAxes(invAssemblyDocument, colLeafSummaries, objCounters)
        Dim colHoleAxes As System.Collections.Generic.List(Of CylinderAxis) = SelectHoleAxes(colAxes, objCounters)
        Dim colResults As System.Collections.Generic.List(Of AxisMatchResult) = MatchAxes(colHoleAxes, dblCandidateRange, dblAlignmentTolerance, objCounters)

        For Each objAxis As CylinderAxis In colHoleAxes
            colHoleFaces.Add(objAxis.invFace)
        Next
        For Each objAxis As CylinderAxis In colAxes
            If objAxis.strFeatureState = strFeatureStateOuterRing Then
                colOuterRingFaces.Add(objAxis.invFace)
            End If
        Next

        ' W05 임시 진단 출력 · W06 피처 판정 열 추가
        strDiagnosticPath = WriteDiagnostics(invAssemblyDocument, dblCandidateRange, dblAlignmentTolerance, colLeafSummaries, colAxes, colHoleAxes, colResults, objCounters)

        Return colResults
    End Function

    ' 1단계 · 수집 · AllLeafOccurrences 순회 · 리프 Transformation과 프록시 Geometry에 변환을 곱하지 않는다 (W04 결정)
    Private Function CollectCylinderAxes(ByVal invAssemblyDocument As Inventor.AssemblyDocument, ByVal colLeafSummaries As System.Collections.Generic.List(Of LeafSummary), ByVal objCounters As MatchCounters) As System.Collections.Generic.List(Of CylinderAxis)
        Dim colAxes As New System.Collections.Generic.List(Of CylinderAxis)()
        Dim intLeafIndex As Integer = 0

        For Each invLeaf As Inventor.ComponentOccurrence In invAssemblyDocument.ComponentDefinition.Occurrences.AllLeafOccurrences()
            Dim objSummary As New LeafSummary()
            objSummary.strOccurrenceName = invLeaf.Name
            objSummary.strDefinitionName = GetDefinitionName(invLeaf)
            colLeafSummaries.Add(objSummary)

            Try
                For Each invBody As Inventor.SurfaceBody In invLeaf.SurfaceBodies
                    For Each invFace As Inventor.Face In invBody.Faces
                        Try
                            If invFace.SurfaceType <> Inventor.SurfaceTypeEnum.kCylinderSurface Then
                                Continue For
                            End If

                            Dim invCylinder As Inventor.Cylinder = CType(invFace.Geometry, Inventor.Cylinder)
                            Dim objAxis As New CylinderAxis()
                            objAxis.intLeafIndex = intLeafIndex
                            objAxis.strOccurrenceName = objSummary.strOccurrenceName
                            objAxis.dblPointX = invCylinder.BasePoint.X * dblCentimeterToMillimeter
                            objAxis.dblPointY = invCylinder.BasePoint.Y * dblCentimeterToMillimeter
                            objAxis.dblPointZ = invCylinder.BasePoint.Z * dblCentimeterToMillimeter
                            objAxis.dblDirectionX = invCylinder.AxisVector.X
                            objAxis.dblDirectionY = invCylinder.AxisVector.Y
                            objAxis.dblDirectionZ = invCylinder.AxisVector.Z
                            objAxis.dblRadius = invCylinder.Radius * dblCentimeterToMillimeter

                            objAxis.invFace = invFace
                            ClassifyFaceFeature(invFace, objAxis)

                            colAxes.Add(objAxis)
                            objSummary.intCylinderCount += 1
                        Catch
                            objSummary.intFailedFaceCount += 1
                        End Try
                    Next
                Next
            Catch objException As System.Exception
                objSummary.strLeafError = objException.Message
            End Try

            objCounters.intCylinderCount += objSummary.intCylinderCount
            objCounters.intFailedFaceCount += objSummary.intFailedFaceCount
            intLeafIndex += 1
        Next

        Return colAxes
    End Function

    ' W06 · 면을 만든 피처에서 HoleFeature 조상을 찾아 objAxis에 상태·이름·경로를 적는다
    ' 읽기 실패는 면 수집 실패와 따로 센다 (원통면 자체는 수집됨)
    Private Sub ClassifyFaceFeature(ByVal invFace As Inventor.Face, ByVal objAxis As CylinderAxis)
        Try
            Dim colPath As New System.Collections.Generic.List(Of String)()
            Dim colVisited As New System.Collections.Generic.HashSet(Of String)()
            Dim invHoleFeature As Inventor.HoleFeature = FindHoleAncestor(invFace.CreatedByFeature, 0, colVisited, colPath)

            objAxis.strFeaturePath = String.Join(" > ", colPath)
            If invHoleFeature Is Nothing Then
                objAxis.strFeatureState = strFeatureStateNotHole
            Else
                objAxis.strFeatureState = strFeatureStateHole
                objAxis.strHoleFeatureName = invHoleFeature.Name
            End If
        Catch objException As System.Exception
            objAxis.strFeatureState = strFeatureStateReadFailed
            objAxis.strFeaturePath = objException.Message
        End Try
    End Sub

    ' W06 · 재귀 · 깊이 상한 intMaxFeatureDepth · 방문한 피처 이름을 기록해 순환을 막는다
    ' 피처 이름은 한 부품 문서 안에서 고유하므로 방문 키로 쓴다
    Private Function FindHoleAncestor(ByVal invFeature As Inventor.PartFeature, ByVal intDepth As Integer, ByVal colVisited As System.Collections.Generic.HashSet(Of String), ByVal colPath As System.Collections.Generic.List(Of String)) As Inventor.HoleFeature
        If invFeature Is Nothing Then
            colPath.Add("(피처 없음)")
            Return Nothing
        End If

        Dim strFeatureName As String = invFeature.Name
        If Not colVisited.Add(strFeatureName) Then
            colPath.Add("(순환 " & strFeatureName & ")")
            Return Nothing
        End If
        colPath.Add(invFeature.Type.ToString() & ":" & strFeatureName)

        If TypeOf invFeature Is Inventor.HoleFeature Then
            Return CType(invFeature, Inventor.HoleFeature)
        End If

        If intDepth >= intMaxFeatureDepth Then
            colPath.Add("(깊이 상한 " & intMaxFeatureDepth.ToString() & ")")
            Return Nothing
        End If

        Dim invParentFeatures As Inventor.ObjectCollection = GetParentFeatures(invFeature)
        If invParentFeatures Is Nothing Then
            Return Nothing
        End If

        For Each objParent As Object In invParentFeatures
            If Not TypeOf objParent Is Inventor.PartFeature Then
                Continue For
            End If
            Dim invHoleFeature As Inventor.HoleFeature = FindHoleAncestor(CType(objParent, Inventor.PartFeature), intDepth + 1, colVisited, colPath)
            If invHoleFeature IsNot Nothing Then
                Return invHoleFeature
            End If
        Next

        Return Nothing
    End Function

    ' W06 · ParentFeatures는 PartFeature 공통 멤버가 아니라 패턴·미러 피처에만 있다 · 그 외 피처는 Nothing
    '       SketchDrivenPatternFeature는 피처에 ParentFeatures가 없고 Definition(SketchDrivenPatternDefinition)에 있다
    Private Function GetParentFeatures(ByVal invFeature As Inventor.PartFeature) As Inventor.ObjectCollection
        If TypeOf invFeature Is Inventor.RectangularPatternFeature Then
            Return CType(invFeature, Inventor.RectangularPatternFeature).ParentFeatures
        End If
        If TypeOf invFeature Is Inventor.CircularPatternFeature Then
            Return CType(invFeature, Inventor.CircularPatternFeature).ParentFeatures
        End If
        If TypeOf invFeature Is Inventor.MirrorFeature Then
            Return CType(invFeature, Inventor.MirrorFeature).ParentFeatures
        End If
        If TypeOf invFeature Is Inventor.SketchDrivenPatternFeature Then
            Return CType(invFeature, Inventor.SketchDrivenPatternFeature).Definition.ParentFeatures
        End If
        Return Nothing
    End Function

    ' W06 · 홀로 확정된 면만 남기고, 같은 occurrence · 같은 HoleFeature · 축 일치(평행 · 축간 거리 dblCoaxialDistanceThreshold 이하) 묶음에서
    '       최소 반지름 하나만 남긴다 · 나머지는 카운터보어 바깥 링으로 보고 짝 계산에서 뺀다
    Private Function SelectHoleAxes(ByVal colAxes As System.Collections.Generic.List(Of CylinderAxis), ByVal objCounters As MatchCounters) As System.Collections.Generic.List(Of CylinderAxis)
        Dim colSelected As New System.Collections.Generic.List(Of CylinderAxis)()
        Dim dblParallelCosine As Double = System.Math.Cos(dblParallelToleranceDegrees * System.Math.PI / 180.0)

        For Each objAxis As CylinderAxis In colAxes
            If objAxis.strFeatureState <> strFeatureStateHole Then
                Continue For
            End If

            Dim intGroupIndex As Integer = -1
            For intIndex As Integer = 0 To colSelected.Count - 1
                Dim objKept As CylinderAxis = colSelected(intIndex)
                If objKept.intLeafIndex <> objAxis.intLeafIndex OrElse objKept.strHoleFeatureName <> objAxis.strHoleFeatureName Then
                    Continue For
                End If

                Dim dblDot As Double = objKept.dblDirectionX * objAxis.dblDirectionX +
                                       objKept.dblDirectionY * objAxis.dblDirectionY +
                                       objKept.dblDirectionZ * objAxis.dblDirectionZ
                If System.Math.Abs(dblDot) < dblParallelCosine Then
                    Continue For
                End If

                If GetPointToLineDistance(objAxis, objKept) <= dblCoaxialDistanceThreshold Then
                    intGroupIndex = intIndex
                    Exit For
                End If
            Next

            If intGroupIndex < 0 Then
                colSelected.Add(objAxis)
            ElseIf objAxis.dblRadius < colSelected(intGroupIndex).dblRadius Then
                colSelected(intGroupIndex).strFeatureState = strFeatureStateOuterRing
                colSelected(intGroupIndex) = objAxis
            Else
                objAxis.strFeatureState = strFeatureStateOuterRing
            End If
        Next

        For Each objAxis As CylinderAxis In colAxes
            Select Case objAxis.strFeatureState
                Case strFeatureStateHole
                    objCounters.intHoleCount += 1
                Case strFeatureStateOuterRing
                    objCounters.intOuterRingCount += 1
                Case strFeatureStateNotHole
                    objCounters.intNotHoleCount += 1
                Case strFeatureStateReadFailed
                    objCounters.intFeatureReadFailedCount += 1
            End Select
        Next

        Return colSelected
    End Function

    ' 2단계 · 짝 계산 · 서로 다른 리프끼리 전수 비교 · 중복 제거·배타 매칭 없음
    Private Function MatchAxes(ByVal colAxes As System.Collections.Generic.List(Of CylinderAxis), ByVal dblCandidateRange As Double, ByVal dblAlignmentTolerance As Double, ByVal objCounters As MatchCounters) As System.Collections.Generic.List(Of AxisMatchResult)
        Dim colResults As New System.Collections.Generic.List(Of AxisMatchResult)()
        Dim dblParallelCosine As Double = System.Math.Cos(dblParallelToleranceDegrees * System.Math.PI / 180.0)

        For intIndexA As Integer = 0 To colAxes.Count - 2
            Dim objAxisA As CylinderAxis = colAxes(intIndexA)

            For intIndexB As Integer = intIndexA + 1 To colAxes.Count - 1
                Dim objAxisB As CylinderAxis = colAxes(intIndexB)
                If objAxisA.intLeafIndex = objAxisB.intLeafIndex Then
                    Continue For
                End If

                objCounters.intComparisonCount += 1

                ' 평행 판정 · 내적 절댓값 (반대 방향도 평행)
                Dim dblDot As Double = objAxisA.dblDirectionX * objAxisB.dblDirectionX +
                                       objAxisA.dblDirectionY * objAxisB.dblDirectionY +
                                       objAxisA.dblDirectionZ * objAxisB.dblDirectionZ
                If System.Math.Abs(dblDot) < dblParallelCosine Then
                    objCounters.intNotParallelCount += 1
                    Continue For
                End If

                Dim dblAxisDistance As Double = GetPointToLineDistance(objAxisA, objAxisB)
                If dblAxisDistance > dblCandidateRange Then
                    objCounters.intOutOfRangeCount += 1
                    Continue For
                End If

                Dim objResult As New AxisMatchResult()
                If dblAxisDistance <= dblAlignmentTolerance Then
                    objResult.strErrorKind = AxisMatchResult.strErrorKindOk
                    objCounters.intOkCount += 1
                Else
                    objResult.strErrorKind = AxisMatchResult.strErrorKindPositionMismatch
                    objCounters.intPositionMismatchCount += 1
                End If
                objResult.dblAxisDistance = dblAxisDistance
                objResult.dblDiameterA = objAxisA.dblRadius * 2.0
                objResult.dblDiameterB = objAxisB.dblRadius * 2.0
                objResult.strOccurrenceNameA = objAxisA.strOccurrenceName
                objResult.strOccurrenceNameB = objAxisB.strOccurrenceName
                ' 이동 목표점 · 두 축 위 점의 중점
                objResult.dblTargetX = (objAxisA.dblPointX + objAxisB.dblPointX) / 2.0
                objResult.dblTargetY = (objAxisA.dblPointY + objAxisB.dblPointY) / 2.0
                objResult.dblTargetZ = (objAxisA.dblPointZ + objAxisB.dblPointZ) / 2.0

                colResults.Add(objResult)
            Next
        Next

        Return colResults
    End Function

    ' A 축 위의 점에서 B 축 직선까지 거리 · |(PA - PB) × uB| (uB는 단위벡터) · skew line 일반 공식은 쓰지 않는다
    Private Function GetPointToLineDistance(ByVal objAxisA As CylinderAxis, ByVal objAxisB As CylinderAxis) As Double
        Dim dblDeltaX As Double = objAxisA.dblPointX - objAxisB.dblPointX
        Dim dblDeltaY As Double = objAxisA.dblPointY - objAxisB.dblPointY
        Dim dblDeltaZ As Double = objAxisA.dblPointZ - objAxisB.dblPointZ

        Dim dblCrossX As Double = dblDeltaY * objAxisB.dblDirectionZ - dblDeltaZ * objAxisB.dblDirectionY
        Dim dblCrossY As Double = dblDeltaZ * objAxisB.dblDirectionX - dblDeltaX * objAxisB.dblDirectionZ
        Dim dblCrossZ As Double = dblDeltaX * objAxisB.dblDirectionY - dblDeltaY * objAxisB.dblDirectionX

        Return System.Math.Sqrt(dblCrossX * dblCrossX + dblCrossY * dblCrossY + dblCrossZ * dblCrossZ)
    End Function

    Private Function GetDefinitionName(ByVal invLeaf As Inventor.ComponentOccurrence) As String
        Try
            Return CType(invLeaf.Definition.Document, Inventor.Document).DisplayName
        Catch objException As System.Exception
            Return "읽기 실패: " & objException.Message
        End Try
    End Function

    ' 3단계 · W05 임시 진단 출력 · ApiProbe와 같은 방식으로 바탕화면에 저장 · 저장 경로를 돌려준다
    Private Function WriteDiagnostics(ByVal invAssemblyDocument As Inventor.AssemblyDocument, ByVal dblCandidateRange As Double, ByVal dblAlignmentTolerance As Double, ByVal colLeafSummaries As System.Collections.Generic.List(Of LeafSummary), ByVal colAxes As System.Collections.Generic.List(Of CylinderAxis), ByVal colHoleAxes As System.Collections.Generic.List(Of CylinderAxis), ByVal colResults As System.Collections.Generic.List(Of AxisMatchResult), ByVal objCounters As MatchCounters) As String
        Dim objText As New System.Text.StringBuilder()
        objText.AppendLine("단위 mm (API cm ×10) · 좌표는 조립품 좌표")
        objText.AppendLine("문서: " & invAssemblyDocument.FullFileName)
        objText.AppendLine("시각: " & System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        objText.AppendLine("판정 재료: 피처 기반 (W06) · HoleFeature 조상이 있는 원통면 · 카운터보어 바깥 링 제외 · 재귀 깊이 상한 " & intMaxFeatureDepth.ToString())
        objText.AppendLine("후보 범위 " & FormatDouble(dblCandidateRange) & " mm · 정렬 공차 " & FormatDouble(dblAlignmentTolerance) & " mm · 평행 허용각 " & FormatDouble(dblParallelToleranceDegrees) & " 도")
        objText.AppendLine()

        objText.AppendLine("==== occurrence별 원통면 수 ====")
        objText.AppendLine(JoinTab("번호", "occurrence", "부품 정의", "원통면", "읽기 실패 면", "홀 확정", "바깥 링 제외", "비교 투입", "리프 오류"))
        For intIndex As Integer = 0 To colLeafSummaries.Count - 1
            Dim objSummary As LeafSummary = colLeafSummaries(intIndex)
            Dim intOuterRingCount As Integer = 0
            Dim intSelectedCount As Integer = 0
            For Each objAxis As CylinderAxis In colAxes
                If objAxis.intLeafIndex <> intIndex Then
                    Continue For
                End If
                If objAxis.strFeatureState = strFeatureStateOuterRing Then
                    intOuterRingCount += 1
                ElseIf objAxis.strFeatureState = strFeatureStateHole Then
                    intSelectedCount += 1
                End If
            Next
            objText.AppendLine(JoinTab(intIndex.ToString(),
                                       objSummary.strOccurrenceName,
                                       objSummary.strDefinitionName,
                                       objSummary.intCylinderCount.ToString(),
                                       objSummary.intFailedFaceCount.ToString(),
                                       (intSelectedCount + intOuterRingCount).ToString(),
                                       intOuterRingCount.ToString(),
                                       intSelectedCount.ToString(),
                                       objSummary.strLeafError))
        Next
        objText.AppendLine()

        WriteCylinderList(objText, colAxes)
        WriteInnerSpacing(objText, colLeafSummaries, colAxes)

        objText.AppendLine("==== 결과 전체 (" & colResults.Count.ToString() & "줄) ====")
        objText.AppendLine(JoinTab("strErrorKind", "dblAxisDistance", "dblDiameterA", "dblDiameterB", "strOccurrenceNameA", "strOccurrenceNameB", "dblTargetX", "dblTargetY", "dblTargetZ"))
        For Each objResult As AxisMatchResult In colResults
            objText.AppendLine(JoinTab(objResult.strErrorKind,
                                       FormatDouble(objResult.dblAxisDistance),
                                       FormatDouble(objResult.dblDiameterA),
                                       FormatDouble(objResult.dblDiameterB),
                                       objResult.strOccurrenceNameA,
                                       objResult.strOccurrenceNameB,
                                       FormatDouble(objResult.dblTargetX),
                                       FormatDouble(objResult.dblTargetY),
                                       FormatDouble(objResult.dblTargetZ)))
        Next
        objText.AppendLine()

        objText.AppendLine("==== 카운터 ====")
        objText.AppendLine("총 원통면: " & objCounters.intCylinderCount.ToString())
        objText.AppendLine("읽기 실패 면 (건너뜀): " & objCounters.intFailedFaceCount.ToString())
        objText.AppendLine("피처 판정 · 홀 확정: " & (objCounters.intHoleCount + objCounters.intOuterRingCount).ToString())
        objText.AppendLine("피처 판정 · 홀 아님 (제외): " & objCounters.intNotHoleCount.ToString())
        objText.AppendLine("피처 판정 · 읽기 실패 (제외): " & objCounters.intFeatureReadFailedCount.ToString())
        objText.AppendLine("카운터보어 바깥 링 제외: " & objCounters.intOuterRingCount.ToString())
        objText.AppendLine("비교 투입 원통면: " & colHoleAxes.Count.ToString())
        objText.AppendLine("비교 횟수 (서로 다른 occurrence 쌍): " & objCounters.intComparisonCount.ToString())
        objText.AppendLine("평행 아님 제외: " & objCounters.intNotParallelCount.ToString())
        objText.AppendLine("후보 범위 밖: " & objCounters.intOutOfRangeCount.ToString())
        objText.AppendLine("OK: " & objCounters.intOkCount.ToString())
        objText.AppendLine("위치 불일치: " & objCounters.intPositionMismatchCount.ToString())

        Dim strDocumentName As String = System.IO.Path.GetFileNameWithoutExtension(invAssemblyDocument.DisplayName)
        For Each chrInvalid As Char In System.IO.Path.GetInvalidFileNameChars()
            strDocumentName = strDocumentName.Replace(chrInvalid, "_"c)
        Next
        Dim strFileName As String = "HoleCheck_Feature_" & strDocumentName & "_" & System.DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".txt"
        Dim strPath As String = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.DesktopDirectory), strFileName)

        ' BOM 포함 UTF-8 · 메모장·PowerShell에서 한글이 깨지지 않게
        System.IO.File.WriteAllText(strPath, objText.ToString(), New System.Text.UTF8Encoding(True))

        System.Windows.Forms.MessageBox.Show("저장됨" & System.Environment.NewLine & strPath, "Feature Matching")
        Return strPath
    End Function

    ' 진단 · 원통면 전체 목록
    Private Sub WriteCylinderList(ByVal objText As System.Text.StringBuilder, ByVal colAxes As System.Collections.Generic.List(Of CylinderAxis))
        objText.AppendLine("==== 원통면 전체 목록 (" & colAxes.Count.ToString() & "개) ====")
        objText.AppendLine(JoinTab("리프 번호", "occurrence", "지름", "점X", "점Y", "점Z", "방향X", "방향Y", "방향Z", "피처 판정", "HoleFeature", "피처 경로"))
        For Each objAxis As CylinderAxis In colAxes
            objText.AppendLine(JoinTab(objAxis.intLeafIndex.ToString(),
                                       objAxis.strOccurrenceName,
                                       FormatDouble(objAxis.dblRadius * 2.0),
                                       FormatDouble(objAxis.dblPointX),
                                       FormatDouble(objAxis.dblPointY),
                                       FormatDouble(objAxis.dblPointZ),
                                       FormatDouble(objAxis.dblDirectionX),
                                       FormatDouble(objAxis.dblDirectionY),
                                       FormatDouble(objAxis.dblDirectionZ),
                                       objAxis.strFeatureState,
                                       objAxis.strHoleFeatureName,
                                       objAxis.strFeaturePath))
        Next
        objText.AppendLine()
    End Sub

    ' 진단 · 같은 occurrence 안에서 평행한 원통면끼리 축간 거리 최소값 · 출력만 하고 판정에는 쓰지 않는다
    Private Sub WriteInnerSpacing(ByVal objText As System.Text.StringBuilder, ByVal colLeafSummaries As System.Collections.Generic.List(Of LeafSummary), ByVal colAxes As System.Collections.Generic.List(Of CylinderAxis))
        Dim dblParallelCosine As Double = System.Math.Cos(dblParallelToleranceDegrees * System.Math.PI / 180.0)
        Dim dblLeafMinimum(colLeafSummaries.Count - 1) As Double
        For intIndex As Integer = 0 To dblLeafMinimum.Length - 1
            dblLeafMinimum(intIndex) = System.Double.MaxValue
        Next

        For intIndexA As Integer = 0 To colAxes.Count - 2
            Dim objAxisA As CylinderAxis = colAxes(intIndexA)
            For intIndexB As Integer = intIndexA + 1 To colAxes.Count - 1
                Dim objAxisB As CylinderAxis = colAxes(intIndexB)
                If objAxisA.intLeafIndex <> objAxisB.intLeafIndex Then
                    Continue For
                End If

                Dim dblDot As Double = objAxisA.dblDirectionX * objAxisB.dblDirectionX +
                                       objAxisA.dblDirectionY * objAxisB.dblDirectionY +
                                       objAxisA.dblDirectionZ * objAxisB.dblDirectionZ
                If System.Math.Abs(dblDot) < dblParallelCosine Then
                    Continue For
                End If

                Dim dblAxisDistance As Double = GetPointToLineDistance(objAxisA, objAxisB)
                If dblAxisDistance <= dblCoaxialDistanceThreshold Then
                    Continue For
                End If

                If dblAxisDistance < dblLeafMinimum(objAxisA.intLeafIndex) Then
                    dblLeafMinimum(objAxisA.intLeafIndex) = dblAxisDistance
                End If
            Next
        Next

        objText.AppendLine("==== 같은 occurrence 안의 구멍 간격 (평행 · 동축 " & FormatDouble(dblCoaxialDistanceThreshold) & " mm 이하 제외 · 판정에 쓰지 않음) ====")
        objText.AppendLine(JoinTab("리프 번호", "occurrence", "최소 축간 거리"))
        Dim dblOverallMinimum As Double = System.Double.MaxValue
        Dim intOverallIndex As Integer = -1
        For intIndex As Integer = 0 To dblLeafMinimum.Length - 1
            Dim strMinimum As String = "없음"
            If dblLeafMinimum(intIndex) < System.Double.MaxValue Then
                strMinimum = FormatDouble(dblLeafMinimum(intIndex))
                If dblLeafMinimum(intIndex) < dblOverallMinimum Then
                    dblOverallMinimum = dblLeafMinimum(intIndex)
                    intOverallIndex = intIndex
                End If
            End If
            objText.AppendLine(JoinTab(intIndex.ToString(), colLeafSummaries(intIndex).strOccurrenceName, strMinimum))
        Next

        If intOverallIndex >= 0 Then
            objText.AppendLine("전체 최소: " & FormatDouble(dblOverallMinimum) & " mm · " & colLeafSummaries(intOverallIndex).strOccurrenceName & " (리프 번호 " & intOverallIndex.ToString() & ")")
        Else
            objText.AppendLine("전체 최소: 없음")
        End If
        objText.AppendLine()
    End Sub

    Private Function JoinTab(ParamArray strValues As String()) As String
        Return String.Join(System.Convert.ToChar(9).ToString(), strValues)
    End Function

    Private Function FormatDouble(ByVal dblValue As Double) As String
        Return dblValue.ToString("R", System.Globalization.CultureInfo.InvariantCulture)
    End Function

End Module
