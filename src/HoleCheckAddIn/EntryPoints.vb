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

    Private invApplication As Inventor.Application

    Public Sub Activate(ByVal AddInSiteObject As Inventor.ApplicationAddInSite, ByVal FirstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate
        Try
            invApplication = AddInSiteObject.Application

            System.Windows.Forms.MessageBox.Show("HoleCheck 로드됨", "HoleCheck")
        Catch objException As System.Exception
            System.Windows.Forms.MessageBox.Show(objException.Message & System.Environment.NewLine & System.Environment.NewLine & objException.StackTrace, "HoleCheck · Activate 오류")
        End Try
    End Sub

    Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate
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

End Class
