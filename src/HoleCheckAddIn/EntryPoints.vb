Option Strict On
Option Explicit On

Imports System.Runtime.InteropServices
Imports Inventor

' Inventor가 .addin 매니페스트의 ClassId로 찾아 생성하는 진입 클래스
' GUID는 HoleCheckAddIn.addin 의 ClassId·ClientId 와 같아야 한다
' 네임스페이스는 프로젝트 RootNamespace(HoleCheckAddIn)를 그대로 쓴다
<ComVisible(True)>
<GuidAttribute("475D83A4-5AE7-414F-ADED-296ED976490A")>
Public Class StandardAddInServer
    Implements Inventor.ApplicationAddInServer

    ' 리본·버튼 정의의 ClientId · decisions.md "W04 · Add-in GUID 확정"
    Private Const strClientId As String = "{475D83A4-5AE7-414F-ADED-296ED976490A}"

    Private Const strRunInternalName As String = "HoleCheck_Run"
    Private Const strApiProbeInternalName As String = "HoleCheck_ApiProbe"
    Private Const strTabInternalName As String = "HoleCheck_Tab"
    Private Const strPanelInternalName As String = "HoleCheck_Panel"

    Private invApplication As Inventor.Application
    Private invRunButton As Inventor.ButtonDefinition
    Private invApiProbeButton As Inventor.ButtonDefinition

    Public Sub Activate(ByVal AddInSiteObject As Inventor.ApplicationAddInSite, ByVal FirstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate
        Try
            invApplication = AddInSiteObject.Application

            invRunButton = GetOrAddButtonDefinition(strRunInternalName, "Hole Check", "조립품 구멍 짝 검사 창을 연다")
            invApiProbeButton = GetOrAddButtonDefinition(strApiProbeInternalName, "API Probe", "임시 진단 · API 실측 결과를 텍스트 파일로 출력")

            AddHandler invRunButton.OnExecute, AddressOf OnRunExecute
            AddHandler invApiProbeButton.OnExecute, AddressOf OnApiProbeExecute

            ' Assembly 리본에만 만든다 (Part 리본 아님)
            Dim invRibbon As Inventor.Ribbon = invApplication.UserInterfaceManager.Ribbons.Item("Assembly")
            Dim invTab As Inventor.RibbonTab = GetOrAddTab(invRibbon)
            Dim invPanel As Inventor.RibbonPanel = GetOrAddPanel(invTab)

            AddButtonIfMissing(invPanel, invRunButton)
            AddButtonIfMissing(invPanel, invApiProbeButton)
        Catch objException As System.Exception
            System.Windows.Forms.MessageBox.Show(objException.Message & System.Environment.NewLine & System.Environment.NewLine & objException.StackTrace, "HoleCheck · Activate 오류")
        End Try
    End Sub

    Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate
        ' Inventor 종료 중에는 Delete가 실패할 수 있어 예외를 삼킨다 · 참조 해제는 반드시 진행
        If invRunButton IsNot Nothing Then
            RemoveHandler invRunButton.OnExecute, AddressOf OnRunExecute
            Try
                invRunButton.Delete()
            Catch
            End Try
            invRunButton = Nothing
        End If

        If invApiProbeButton IsNot Nothing Then
            RemoveHandler invApiProbeButton.OnExecute, AddressOf OnApiProbeExecute
            Try
                invApiProbeButton.Delete()
            Catch
            End Try
            invApiProbeButton = Nothing
        End If

        invApplication = Nothing

        System.GC.Collect()
        System.GC.WaitForPendingFinalizers()
    End Sub

    ' 구식 명령 방식용 · 사용하지 않음 (2027 시그니처 변경 여부는 빌드 오류로 확인 필요)
    Public Sub ExecuteCommand(ByVal CommandID As Integer) Implements Inventor.ApplicationAddInServer.ExecuteCommand
    End Sub

    ' 외부에 노출할 자동화 객체 없음
    Public ReadOnly Property Automation() As Object Implements Inventor.ApplicationAddInServer.Automation
        Get
            Return Nothing
        End Get
    End Property

    Private Sub OnRunExecute(ByVal Context As Inventor.NameValueMap)
        Try
            Dim invAssemblyDocument As Inventor.AssemblyDocument = GetActiveAssemblyDocument()
            If invAssemblyDocument Is Nothing Then
                Return
            End If

            Dim objForm As New HoleCheckForm(invApplication)
            objForm.Show()
        Catch objException As System.Exception
            System.Windows.Forms.MessageBox.Show(objException.Message & System.Environment.NewLine & System.Environment.NewLine & objException.StackTrace, "HoleCheck · Hole Check 오류")
        End Try
    End Sub

    Private Sub OnApiProbeExecute(ByVal Context As Inventor.NameValueMap)
        Try
            Dim invAssemblyDocument As Inventor.AssemblyDocument = GetActiveAssemblyDocument()
            If invAssemblyDocument Is Nothing Then
                Return
            End If

            ApiProbe.Run(invAssemblyDocument)
        Catch objException As System.Exception
            System.Windows.Forms.MessageBox.Show(objException.Message & System.Environment.NewLine & System.Environment.NewLine & objException.StackTrace, "HoleCheck · API Probe 오류")
        End Try
    End Sub

    ' 활성 문서가 조립품이 아니면 안내 후 Nothing
    Private Function GetActiveAssemblyDocument() As Inventor.AssemblyDocument
        Dim invAssemblyDocument As Inventor.AssemblyDocument = TryCast(invApplication.ActiveDocument, Inventor.AssemblyDocument)
        If invAssemblyDocument Is Nothing Then
            System.Windows.Forms.MessageBox.Show("조립품(.iam) 문서를 열고 실행하세요.", "HoleCheck")
        End If
        Return invAssemblyDocument
    End Function

    ' 비정상 종료로 Deactivate가 안 불렸을 때 같은 이름의 정의가 남아 있을 수 있어 먼저 찾는다
    Private Function GetOrAddButtonDefinition(ByVal strInternalName As String, ByVal strDisplayName As String, ByVal strDescription As String) As Inventor.ButtonDefinition
        Dim invControlDefinitions As Inventor.ControlDefinitions = invApplication.CommandManager.ControlDefinitions

        Try
            Return CType(invControlDefinitions.Item(strInternalName), Inventor.ButtonDefinition)
        Catch
            ' Item은 없으면 예외를 던진다 · 아래에서 새로 만든다
        End Try

        ' 인자 순서(2026 이전 기준) · DisplayName, InternalName, Classification, ClientId,
        '   DescriptionText, ToolTipText, StandardIcon, LargeIcon, ButtonDisplay
        ' 2027에서 아이콘 인자가 Object · 설명 인자 이름이 DescriptionText 라고 전달받음.
        ' 순서·타입이 바뀌었는지는 확인하지 못함 → 빌드 오류로 확인 필요
        Return invControlDefinitions.AddButtonDefinition(strDisplayName, strInternalName, Inventor.CommandTypesEnum.kQueryOnlyCmdType, strClientId, strDescription, strDisplayName, Nothing, Nothing, Inventor.ButtonDisplayEnum.kAlwaysDisplayText)
    End Function

    Private Function GetOrAddTab(ByVal invRibbon As Inventor.Ribbon) As Inventor.RibbonTab
        For Each invTab As Inventor.RibbonTab In invRibbon.RibbonTabs
            If invTab.InternalName = strTabInternalName Then
                Return invTab
            End If
        Next

        ' 인자 · DisplayName, InternalName, ClientId (위치 지정 인자는 생략) · 빌드 오류로 확인 필요
        Return invRibbon.RibbonTabs.Add("HoleCheck", strTabInternalName, strClientId)
    End Function

    Private Function GetOrAddPanel(ByVal invTab As Inventor.RibbonTab) As Inventor.RibbonPanel
        For Each invPanel As Inventor.RibbonPanel In invTab.RibbonPanels
            If invPanel.InternalName = strPanelInternalName Then
                Return invPanel
            End If
        Next

        ' 인자 · DisplayName, InternalName, ClientId (위치 지정 인자는 생략) · 빌드 오류로 확인 필요
        Return invTab.RibbonPanels.Add("검사", strPanelInternalName, strClientId)
    End Function

    Private Sub AddButtonIfMissing(ByVal invPanel As Inventor.RibbonPanel, ByVal invButton As Inventor.ButtonDefinition)
        For Each invControl As Inventor.CommandControl In invPanel.CommandControls
            If invControl.InternalName = invButton.InternalName Then
                Return
            End If
        Next

        ' 인자 · ButtonDefinition, UseLargeIcon, ShowText · 빌드 오류로 확인 필요
        invPanel.CommandControls.AddButton(invButton, True, True)
    End Sub

End Class
