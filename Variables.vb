Imports TemplateDB

Module Variables
    Public OverClass As OverClass
    Public Const ReportPath As String = "M:\VOLUNTEER SCREENING SERVICES\Systems\Resource_Scheduling_System\Reports\"
    Private Const TablePath As String = "M:\VOLUNTEER SCREENING SERVICES\Systems\Resource_Scheduling_System\Backend.accdb"
    Private Const PWord As String = "RetroRetro*1"
    Private Const Connect2 As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & TablePath & "';Jet OLEDB:Database Password=" & PWord
    Private Const UserTable As String = "[Users]"
    Private Const UserField As String = "Username"
    Private Const LockTable As String = "[Locker]"
    Private Const AuditTable As String = "[Audit]"
    Private Contact As String = "Michael Gray"
    Public Const SolutionName As String = "Resource Scheduling System"
    Public PickCohort As Long

    Public Function GetTheConnection() As String
        GetTheConnection = Connect2
    End Function


    Public Sub StartUp()

        OverClass = New TemplateDB.OverClass
        OverClass.SetPrivate(UserTable,
                           UserField,
                           Contact,
                           Connect2,
                           AuditTable)

        OverClass.LoginCheck()

        OverClass.AddAllDataItem(Form1)

        For Each ctl In OverClass.DataItemCollection
            If (TypeOf ctl Is Button) Then
                Dim But As Button = ctl
                AddHandler But.Click, AddressOf ButtonSpecifics
            End If
        Next


    End Sub

End Module
