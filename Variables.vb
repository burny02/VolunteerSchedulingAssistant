Imports TemplateDB
Module Variables
    Public Central As TemplateDB.Template
    Private Const TablePath As String = "M:\VOLUNTEER SCREENING SERVICES\DavidBurnside\\Scheduler\Backend2.accdb"
    Private Const PWord As String = "Crypto*Dave02"
    Private Const Connect2 As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & TablePath & ";Jet OLEDB:Database Password=" & PWord
    Private Const UserTable As String = "[Users]"
    Private Const UserField As String = "Username"
    Private Const LockTable As String = "[Locker]"
    Private Const ActiveUsersTable As String = "[ActiveUsers]"
    Private Contact As String = "Michael Gray"
    Public Const SolutionName As String = "Volunteer Scheduling Assistant"


    Public Sub StartUpCentral()

        Central = New TemplateDB.Template
        Central.SetPrivate(UserTable, _
                           UserField, _
                           LockTable, _
                           Contact, _
                           Connect2,
                           ActiveUsersTable)
    End Sub
End Module
