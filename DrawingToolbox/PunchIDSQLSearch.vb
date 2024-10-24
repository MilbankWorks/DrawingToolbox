Imports System.Collections.Generic
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.Threading
Imports System.Ling
Imports System.IO

Public Class PunchIDSQLSearch
    Protected connectionString As String = "Data Source= GAR-EPDMDB01; Initial Catalog=PDM_Milbank;User ID =epdm_viewer;password=epdm_viewer1!" 'connection string including info about connection
    Protected objReader As System.IO.StreamReader
    Protected reader As SqlClient.SqlDataReader
    Protected conn As SqlClient.SqlConnection
    Protected cmd As SqlCommand
    Private results As String

    Public myList0 As New List(Of String)()
    Public myList1 As New List(Of String)()
    Public myList2 As New List(Of Int16)()

    Public Sub BuildTable()
        conn = New System.Data.SqlClient.SqlConnection(connectionString)
        cmd = conn.CreateCommand
        'cmd.CommandText = $"exec ENG_Reporting.dbo.PDM_Design_Library"
        cmd.CommandText = $"SELECT * FROM ENG_Reporting.dbo.DesignLibraryPunchID ORDER BY UniqueDesc Desc"
        conn.Open()
        reader = cmd.ExecuteReader()

        Dim iter As Integer
        iter = 0

        Do While reader.Read()
            myList0.Add(reader.GetValue(0))

            Try
                myList1.Add(reader.GetValue(1))
            Catch ex As Exception
                myList1.Add("")
            End Try

            myList2.Add(reader.GetValue(2))
            'Debug.Print(reader.GetValue(0) & "|" & reader.GetValue(1) & "|" & reader.GetValue(2))
            'Console.WriteLine(myList0 & "|" & myList1 & "|" & myList2)
        Loop
        reader.Close()
        conn.Close()
    End Sub

    Public Function TableLookUp(InputString)
        Dim Outcome As Boolean
        Outcome = False
        Dim OutputString As String
        OutputString = ""
        For Each PunchID In myList0
            'Console.WriteLine(InputString & " | " & PunchID & " | " & myList2(myList0.IndexOf(PunchID)) & " | " & InputString.Equals(PunchID))
            If InputString.Equals(PunchID) And myList2(myList0.IndexOf(PunchID)) = 1 Then
                OutputString = If(myList1(myList0.IndexOf(PunchID)) = "", "NO DESCRIPTION", myList1(myList0.IndexOf(PunchID)))
                'Console.WriteLine("Result found: " & OutputString)
                Outcome = True
                Exit For
            ElseIf InputString.Equals(PunchID) And myList2(myList0.IndexOf(PunchID)) > 1 Then
                OutputString = "INVALID DESCRIPTION"
                'Console.WriteLine("Result found but is not unique and invalid.")
                Outcome = True
                Exit For
            End If
        Next
        'Debug.Print(InputString & "Input")
        'Debug.Print(OutputString & "Output")

        Return OutputString
    End Function
End Class
