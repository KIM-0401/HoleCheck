Option Strict On
Option Explicit On

' 디자이너 파일 없이 코드로 만든 폼 · W05 원통면 축 기반 1차 매칭 결과 표시
' 입력 2개(후보 범위 · 정렬 공차) → RunAxisMatching → 4열 그리드 · 더블클릭 시 Camera.Fit 후 목표점으로 이동
' W06 · 실행 후 HighlightSet 색 표시 (홀 면 초록 · 바깥 링 주황)
Public Class HoleCheckForm
    Inherits System.Windows.Forms.Form

    Private invApplication As Inventor.Application

    ' 그리드 줄 번호(0부터) = 이 목록의 인덱스 · 정렬을 막아 순서를 유지한다
    Private colResults As New System.Collections.Generic.List(Of AxisMatchResult)()

    Private objCandidateRangeInput As System.Windows.Forms.NumericUpDown
    Private objAlignmentToleranceInput As System.Windows.Forms.NumericUpDown
    Private objRunButton As System.Windows.Forms.Button
    Private objStatusLabel As System.Windows.Forms.Label
    Private objResultGrid As System.Windows.Forms.DataGridView
    Private objSelectionLabel As System.Windows.Forms.Label

    ' W06 · 하이라이트 세트 수명 확인용 진단 · 함수가 끝나도 참조를 유지한다
    ' 실행할 때마다 새 세트로 덮어쓰기만 하고 이전 세트는 정리하지 않는다 (누적 · 문서 전환 처리는 10주차)
    Private invHoleSet As Inventor.HighlightSet
    Private invOuterRingSet As Inventor.HighlightSet

    Public Sub New(ByVal invInventorApplication As Inventor.Application)
        MyBase.New()

        invApplication = invInventorApplication

        Me.Text = "HoleCheck"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Size = New System.Drawing.Size(640, 480)

        BuildControls()

        AddHandler objRunButton.Click, AddressOf OnRunClick
        AddHandler objResultGrid.SelectionChanged, AddressOf OnGridSelectionChanged
        AddHandler objResultGrid.CellDoubleClick, AddressOf OnGridCellDoubleClick
    End Sub

    Private Sub BuildControls()
        ' 입력 줄
        Dim objInputPanel As New System.Windows.Forms.FlowLayoutPanel()
        objInputPanel.Dock = System.Windows.Forms.DockStyle.Top
        objInputPanel.Height = 36
        objInputPanel.Padding = New System.Windows.Forms.Padding(4)

        objCandidateRangeInput = CreateNumberInput(dblDefaultCandidateRange)
        objAlignmentToleranceInput = CreateNumberInput(dblDefaultAlignmentTolerance)

        objRunButton = New System.Windows.Forms.Button()
        objRunButton.Text = "실행"
        objRunButton.AutoSize = True

        objInputPanel.Controls.Add(CreateInputLabel("후보 범위(mm)"))
        objInputPanel.Controls.Add(objCandidateRangeInput)
        objInputPanel.Controls.Add(CreateInputLabel("정렬 공차(mm)"))
        objInputPanel.Controls.Add(objAlignmentToleranceInput)
        objInputPanel.Controls.Add(objRunButton)

        ' 상태 한 줄
        objStatusLabel = New System.Windows.Forms.Label()
        objStatusLabel.Dock = System.Windows.Forms.DockStyle.Top
        objStatusLabel.Height = 24
        objStatusLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        objStatusLabel.Text = "실행 전"

        ' 그리드 4열
        objResultGrid = New System.Windows.Forms.DataGridView()
        objResultGrid.Dock = System.Windows.Forms.DockStyle.Fill
        objResultGrid.ReadOnly = True
        objResultGrid.AllowUserToAddRows = False
        objResultGrid.AllowUserToDeleteRows = False
        objResultGrid.RowHeadersVisible = False
        objResultGrid.MultiSelect = False
        objResultGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        objResultGrid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        objResultGrid.Columns.Add("colNumber", "번호")
        objResultGrid.Columns.Add("colErrorKind", "오류 종류")
        objResultGrid.Columns.Add("colValue1", "값1 · 축간 거리(mm)")
        objResultGrid.Columns.Add("colValue2", "값2 · A지름 / B지름(mm)")
        For Each objColumn As System.Windows.Forms.DataGridViewColumn In objResultGrid.Columns
            objColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Next

        ' 선택 줄의 occurrence 쌍
        objSelectionLabel = New System.Windows.Forms.Label()
        objSelectionLabel.Dock = System.Windows.Forms.DockStyle.Bottom
        objSelectionLabel.Height = 24
        objSelectionLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        objSelectionLabel.Text = ""

        ' Dock은 나중에 추가한 컨트롤부터 자리를 잡는다 · Fill을 먼저 넣는다
        Me.Controls.Add(objResultGrid)
        Me.Controls.Add(objSelectionLabel)
        Me.Controls.Add(objStatusLabel)
        Me.Controls.Add(objInputPanel)
    End Sub

    Private Function CreateNumberInput(ByVal dblInitialValue As Double) As System.Windows.Forms.NumericUpDown
        Dim objInput As New System.Windows.Forms.NumericUpDown()
        objInput.DecimalPlaces = 3
        objInput.Increment = 0.1D
        objInput.Minimum = 0D
        objInput.Maximum = 1000D
        objInput.Value = CDec(dblInitialValue)
        objInput.Width = 80
        Return objInput
    End Function

    Private Function CreateInputLabel(ByVal strText As String) As System.Windows.Forms.Label
        Dim objLabel As New System.Windows.Forms.Label()
        objLabel.Text = strText
        objLabel.AutoSize = True
        objLabel.Padding = New System.Windows.Forms.Padding(0, 6, 0, 0)
        Return objLabel
    End Function

    Private Sub OnRunClick(ByVal objSender As Object, ByVal objArgs As System.EventArgs)
        Try
            Dim dblCandidateRange As Double = CDbl(objCandidateRangeInput.Value)
            Dim dblAlignmentTolerance As Double = CDbl(objAlignmentToleranceInput.Value)

            If dblCandidateRange <= 0 OrElse dblAlignmentTolerance <= 0 Then
                System.Windows.Forms.MessageBox.Show("후보 범위와 정렬 공차는 0보다 커야 합니다.", "HoleCheck")
                Return
            End If
            If dblAlignmentTolerance >= dblCandidateRange Then
                System.Windows.Forms.MessageBox.Show("정렬 공차는 후보 범위보다 작아야 합니다.", "HoleCheck")
                Return
            End If

            Dim invAssemblyDocument As Inventor.AssemblyDocument = TryCast(invApplication.ActiveDocument, Inventor.AssemblyDocument)
            If invAssemblyDocument Is Nothing Then
                System.Windows.Forms.MessageBox.Show("조립품(.iam) 문서를 활성화하고 실행하세요.", "HoleCheck")
                Return
            End If

            Dim colHoleFaces As New System.Collections.Generic.List(Of Inventor.Face)()
            Dim colOuterRingFaces As New System.Collections.Generic.List(Of Inventor.Face)()
            Dim strDiagnosticPath As String = ""
            colResults = MatchingEngine.RunAxisMatching(invAssemblyDocument, dblCandidateRange, dblAlignmentTolerance, colHoleFaces, colOuterRingFaces, strDiagnosticPath)
            FillGrid()

            HighlightFaces(colHoleFaces, colOuterRingFaces, strDiagnosticPath)

            ' W07 7-5 · EdgeLoop 수집을 축 경로 뒤에 한 번 더 실행 · 실패해도 위의 그리드 · 색 표시는 이미 끝났다
            Dim strLoopDiagnosticPath As String = ""
            Dim colHoleClusters As System.Collections.Generic.List(Of HoleCluster) = MatchingEngine.RunLoopCollection(invAssemblyDocument, strLoopDiagnosticPath)
            objStatusLabel.Text = objStatusLabel.Text & " · 홀 묶음 " & colHoleClusters.Count.ToString() & " · Loop txt 저장됨"
        Catch objException As System.Exception
            System.Windows.Forms.MessageBox.Show(objException.Message & System.Environment.NewLine & System.Environment.NewLine & objException.StackTrace, "HoleCheck · 실행 오류")
        End Try
    End Sub

    ' W06 · 색 표시 · 홀 면 초록 · 바깥 링 주황 · 실행할 때마다 활성 문서에 HighlightSet 2개를 새로 만든다
    ' 세트는 폼 필드에 덮어써 보관하고, 이전 세트는 지우지 않는다 (누적 정리는 W06 범위 밖) · 모델은 바꾸지 않는다
    ' 결과 "색 표시 · 성공 N · 실패 N" 과 세트 진단 2줄은 진단 txt 맨 아래(카운터 끝)에 덧붙인다
    Private Sub HighlightFaces(ByVal colHoleFaces As System.Collections.Generic.List(Of Inventor.Face), ByVal colOuterRingFaces As System.Collections.Generic.List(Of Inventor.Face), ByVal strDiagnosticPath As String)
        Dim invAssemblyDocument As Inventor.AssemblyDocument = TryCast(invApplication.ActiveDocument, Inventor.AssemblyDocument)
        If invAssemblyDocument Is Nothing Then
            Return
        End If

        Dim intSuccessCount As Integer = 0
        Dim intFailureCount As Integer = 0

        invHoleSet = invAssemblyDocument.HighlightSets.Add()
        invHoleSet.Color = invApplication.TransientObjects.CreateColor(0, 255, 0)
        AddFacesToHighlightSet(invHoleSet, colHoleFaces, intSuccessCount, intFailureCount)
        Dim strHoleCount As String = ReadHighlightSetCount(invHoleSet)
        Dim strHoleTypeName As String = ReadFirstItemTypeName(invHoleSet)

        invOuterRingSet = invAssemblyDocument.HighlightSets.Add()
        invOuterRingSet.Color = invApplication.TransientObjects.CreateColor(255, 128, 0)
        AddFacesToHighlightSet(invOuterRingSet, colOuterRingFaces, intSuccessCount, intFailureCount)
        Dim strOuterRingCount As String = ReadHighlightSetCount(invOuterRingSet)
        Dim strOuterRingTypeName As String = ReadFirstItemTypeName(invOuterRingSet)

        If strDiagnosticPath <> "" Then
            Dim objText As New System.Text.StringBuilder()
            objText.AppendLine("색 표시 · 성공 " & intSuccessCount.ToString() & " · 실패 " & intFailureCount.ToString())
            objText.AppendLine("색 표시 진단 · 홀 세트 Count " & strHoleCount & " / 시도 " & colHoleFaces.Count.ToString() &
                               " · 바깥 링 세트 Count " & strOuterRingCount & " / 시도 " & colOuterRingFaces.Count.ToString())

            ' 첫 아이템 TypeName · 읽을 수 있는 세트만 적는다
            Dim colTypeNameParts As New System.Collections.Generic.List(Of String)()
            If strHoleTypeName <> "" Then
                colTypeNameParts.Add("홀 세트 " & strHoleTypeName)
            End If
            If strOuterRingTypeName <> "" Then
                colTypeNameParts.Add("바깥 링 세트 " & strOuterRingTypeName)
            End If
            If colTypeNameParts.Count > 0 Then
                objText.AppendLine("색 표시 진단 · 첫 아이템 TypeName · " & String.Join(" · ", colTypeNameParts))
            End If

            ' 기존 파일 끝에 덧붙이므로 BOM 없이 쓴다
            System.IO.File.AppendAllText(strDiagnosticPath, objText.ToString(), New System.Text.UTF8Encoding(False))
        End If
    End Sub

    ' W06 진단 · HighlightSet.Count (Interop: Count As Int32) · 실패하면 예외 메시지
    Private Function ReadHighlightSetCount(ByVal invHighlightSet As Inventor.HighlightSet) As String
        Try
            Return invHighlightSet.Count.ToString()
        Catch objException As System.Exception
            Return "읽기 실패(" & objException.Message & ")"
        End Try
    End Function

    ' W06 진단 · HighlightSet.Item(Index As Int32) As Object 의 첫 아이템 · 읽지 못하면 빈 문자열
    ' 인덱스 1부터 (Inventor 컬렉션 관례 · 추정)
    Private Function ReadFirstItemTypeName(ByVal invHighlightSet As Inventor.HighlightSet) As String
        Try
            If invHighlightSet.Count < 1 Then
                Return ""
            End If
            Return Microsoft.VisualBasic.Information.TypeName(invHighlightSet.Item(1))
        Catch
            Return ""
        End Try
    End Function

    ' 면 하나가 실패해도 나머지는 계속 넣는다
    Private Sub AddFacesToHighlightSet(ByVal invHighlightSet As Inventor.HighlightSet, ByVal colFaces As System.Collections.Generic.List(Of Inventor.Face), ByRef intSuccessCount As Integer, ByRef intFailureCount As Integer)
        For Each invFace As Inventor.Face In colFaces
            Try
                invHighlightSet.AddItem(invFace)
                intSuccessCount += 1
            Catch
                intFailureCount += 1
            End Try
        Next
    End Sub

    ' OK 줄도 전부 표시
    Private Sub FillGrid()
        objResultGrid.Rows.Clear()
        objSelectionLabel.Text = ""

        Dim intOkCount As Integer = 0
        Dim intPositionMismatchCount As Integer = 0

        For intIndex As Integer = 0 To colResults.Count - 1
            Dim objResult As AxisMatchResult = colResults(intIndex)
            objResultGrid.Rows.Add((intIndex + 1).ToString(),
                                objResult.strErrorKind,
                                FormatMillimeter(objResult.dblAxisDistance),
                                FormatMillimeter(objResult.dblDiameterA) & " / " & FormatMillimeter(objResult.dblDiameterB))

            If objResult.strErrorKind = AxisMatchResult.strErrorKindOk Then
                intOkCount += 1
            ElseIf objResult.strErrorKind = AxisMatchResult.strErrorKindPositionMismatch Then
                intPositionMismatchCount += 1
            End If
        Next

        objStatusLabel.Text = "전체 " & colResults.Count.ToString() & "줄 · OK " & intOkCount.ToString() & " · 위치 불일치 " & intPositionMismatchCount.ToString()
    End Sub

    Private Sub OnGridSelectionChanged(ByVal objSender As Object, ByVal objArgs As System.EventArgs)
        Dim objResult As AxisMatchResult = GetResultAt(If(objResultGrid.CurrentRow Is Nothing, -1, objResultGrid.CurrentRow.Index))
        If objResult Is Nothing Then
            objSelectionLabel.Text = ""
            Return
        End If
        objSelectionLabel.Text = objResult.strOccurrenceNameA & " ↔ " & objResult.strOccurrenceNameB
    End Sub

    ' 더블클릭 · 계획서 5주차대로 Camera.Fit 후 목표점 지정 · Eye는 계산하지 않는다
    Private Sub OnGridCellDoubleClick(ByVal objSender As Object, ByVal objArgs As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            Dim objResult As AxisMatchResult = GetResultAt(objArgs.RowIndex)
            If objResult Is Nothing Then
                Return
            End If

            Zoom.MoveCameraToTarget(invApplication, objResult.dblTargetX, objResult.dblTargetY, objResult.dblTargetZ)
        Catch objException As System.Exception
            System.Windows.Forms.MessageBox.Show(objException.Message & System.Environment.NewLine & System.Environment.NewLine & objException.StackTrace, "HoleCheck · 이동 오류")
        End Try
    End Sub

    Private Function GetResultAt(ByVal intRowIndex As Integer) As AxisMatchResult
        If intRowIndex < 0 OrElse intRowIndex >= colResults.Count Then
            Return Nothing
        End If
        Return colResults(intRowIndex)
    End Function

    Private Function FormatMillimeter(ByVal dblValue As Double) As String
        Return dblValue.ToString("0.000", System.Globalization.CultureInfo.InvariantCulture)
    End Function

End Class
