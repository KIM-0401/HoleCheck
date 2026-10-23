Option Strict On
Option Explicit On

' W07 7-2 · 실행 진입 · 짝 계산 · 진단 출력 (AxisMatching.vb에서 이동)
' 원통면 축 기반 1차 매칭 · A·B 병렬 작업의 접점 파일
' 결과에는 숫자와 문자열만 담는다 (COM 객체 금지 · decisions.md "W05 · 결과 구조에는 숫자와 문자열만 담음")
' 구멍/축 외면 구분 · 지름 판정 · 중복 제거 · 배타 매칭은 하지 않는다 (W05 결정)
' W06 · 수집과 짝 계산 사이에 피처 필터를 둔다 · HoleFeature 조상이 있는 면만, 카운터보어 바깥 링을 뺀 뒤 짝 계산에 넣는다
'       짝 계산(평행 판정 · 점-직선 거리 · 후보 범위 · 정렬 공차)은 W05 그대로
Public Module MatchingEngine

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

    ' W07 7-5 · EdgeLoop 수집 실행 · 축 경로(RunAxisMatching) 옆에 둔다 · 쌍별 거리는 아직 계산하지 않는다 (7-6)
    ' strDiagnosticPath · 저장한 Loop 진단 txt 경로
    Public Function RunLoopCollection(ByVal invAssemblyDocument As Inventor.AssemblyDocument, ByRef strDiagnosticPath As String) As System.Collections.Generic.List(Of HoleCluster)
        Dim objCounters As New LoopCounters()
        Dim colLeafSummaries As New System.Collections.Generic.List(Of LoopLeafSummary)()

        Dim colClusters As System.Collections.Generic.List(Of HoleCluster) = CollectHoleClusters(invAssemblyDocument, colLeafSummaries, objCounters)

        strDiagnosticPath = WriteLoopDiagnostics(invAssemblyDocument, colLeafSummaries, colClusters, objCounters)

        Return colClusters
    End Function

    ' W07 7-5 · Loop 진단 txt · Feature 진단과 같은 폴더(바탕화면) · 같은 인코딩(BOM 포함 UTF-8) · 같은 이름 규칙(접두만 Loop)
    ' 저장 알림 메시지박스는 띄우지 않는다 · 폼 상태 라벨에 표시
    Private Function WriteLoopDiagnostics(ByVal invAssemblyDocument As Inventor.AssemblyDocument, ByVal colLeafSummaries As System.Collections.Generic.List(Of LoopLeafSummary), ByVal colClusters As System.Collections.Generic.List(Of HoleCluster), ByVal objCounters As LoopCounters) As String
        Dim objText As New System.Text.StringBuilder()
        objText.AppendLine("단위 mm (API cm ×10) · 좌표는 조립품 좌표 · Transformation 은 곱하지 않음")
        objText.AppendLine("문서: " & invAssemblyDocument.FullFileName)
        objText.AppendLine("시각: " & System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        objText.AppendLine("수집: 평면 face · 안쪽 EdgeLoop · 모든 edge 가 kCircleCurve · 동심 묶음 기준 " & FormatDouble(dblConcentricClusterThreshold) & " mm (같은 리프 안 · 묶음 중심은 첫 원 중심)")
        objText.AppendLine()

        objText.AppendLine("==== 카운터 ====")
        objText.AppendLine("리프 수: " & objCounters.intLeafCount.ToString())
        objText.AppendLine("평면 face 수: " & objCounters.intPlaneFaceCount.ToString())
        objText.AppendLine("안쪽 루프 수: " & objCounters.intInnerLoopCount.ToString())
        objText.AppendLine("채택 루프 수: " & objCounters.intAcceptedLoopCount.ToString())
        objText.AppendLine("탈락 · 원호 포함: " & objCounters.intRejectedArcCount.ToString())
        objText.AppendLine("탈락 · 기타 곡선 포함: " & objCounters.intRejectedOtherCurveCount.ToString())
        objText.AppendLine("탈락 · 읽기 실패: " & objCounters.intReadFailedCount.ToString())
        objText.AppendLine("edge 2개 이상 채택 루프 수: " & objCounters.intMultiEdgeAcceptedLoopCount.ToString())
        objText.AppendLine("묶음 수: " & objCounters.intClusterCount.ToString())
        objText.AppendLine("지름 2개 이상 묶음 수: " & objCounters.intMultiDiameterClusterCount.ToString())
        objText.AppendLine()

        objText.AppendLine("==== 리프별 ====")
        objText.AppendLine(JoinTab("번호", "occurrence", "부품 정의", "채택 원", "묶음", "지름 2개 이상 묶음", "리프 오류"))
        For intIndex As Integer = 0 To colLeafSummaries.Count - 1
            Dim objSummary As LoopLeafSummary = colLeafSummaries(intIndex)
            objText.AppendLine(JoinTab(intIndex.ToString(),
                                       objSummary.strOccurrenceName,
                                       objSummary.strDefinitionName,
                                       objSummary.intAcceptedCircleCount.ToString(),
                                       objSummary.intClusterCount.ToString(),
                                       objSummary.intMultiDiameterClusterCount.ToString(),
                                       objSummary.strLeafError))
        Next
        objText.AppendLine()

        objText.AppendLine("==== 묶음 목록 (" & colClusters.Count.ToString() & "개) ====")
        objText.AppendLine(JoinTab("리프 번호", "occurrence", "묶음 번호", "중심X", "중심Y", "중심Z", "지름 (오름차순)"))
        Dim intPreviousLeafIndex As Integer = -1
        Dim intClusterNumber As Integer = 0
        For Each objCluster As HoleCluster In colClusters
            ' 묶음 번호는 리프 안에서 0부터 · colClusters 는 리프 순서대로 이어 붙어 있다
            If objCluster.intLeafIndex <> intPreviousLeafIndex Then
                intPreviousLeafIndex = objCluster.intLeafIndex
                intClusterNumber = 0
            End If

            Dim colDiameterTexts As New System.Collections.Generic.List(Of String)()
            For Each dblDiameter As Double In objCluster.colDiameters
                colDiameterTexts.Add(FormatDouble(dblDiameter))
            Next

            objText.AppendLine(JoinTab(objCluster.intLeafIndex.ToString(),
                                       objCluster.strOccurrenceName,
                                       intClusterNumber.ToString(),
                                       FormatDouble(objCluster.dblCenterX),
                                       FormatDouble(objCluster.dblCenterY),
                                       FormatDouble(objCluster.dblCenterZ),
                                       String.Join(" / ", colDiameterTexts)))
            intClusterNumber += 1
        Next

        Dim strDocumentName As String = System.IO.Path.GetFileNameWithoutExtension(invAssemblyDocument.DisplayName)
        For Each chrInvalid As Char In System.IO.Path.GetInvalidFileNameChars()
            strDocumentName = strDocumentName.Replace(chrInvalid, "_"c)
        Next
        Dim strFileName As String = "HoleCheck_Loop_" & strDocumentName & "_" & System.DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".txt"
        Dim strPath As String = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.DesktopDirectory), strFileName)

        ' BOM 포함 UTF-8 · 메모장·PowerShell에서 한글이 깨지지 않게
        System.IO.File.WriteAllText(strPath, objText.ToString(), New System.Text.UTF8Encoding(True))

        Return strPath
    End Function

End Module
