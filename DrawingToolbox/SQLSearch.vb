Imports System.Collections.Generic
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.Threading
Imports System.Linq

''' <summary>
''' Class to execute a SQL Search and return the results in an ArrayList(rows) of Dictionarys(columns)
''' </summary>
Public Class SQLSearch
    Public Sub New()

    End Sub
    Public Sub New(headstring As String, file As String)
        SearchFile = file
        Me.HeadString = headstring
        CommandText = headstring & vbCrLf & SearchFile & vbCrLf & TailString
        conn = New SqlClient.SqlConnection(connectionString)
    End Sub

    Public Sub New(file As String)
        SearchFile = file
        CommandText = HeadString & vbCrLf & SearchFile & vbCrLf & TailString
        conn = New SqlClient.SqlConnection(connectionString)
    End Sub

    Public Sub New(headstring As String, file As String, tailstring As String)
        SearchFile = file
        Me.HeadString = headstring
        Me.TailString = tailstring
        CommandText = headstring & vbCrLf & SearchFile & vbCrLf & tailstring
        conn = New SqlClient.SqlConnection(connectionString)
    End Sub

    Public Sub New(commandText As List(Of String))
        Me.CommandText = ConcatArray(commandText, vbCrLf)
        conn = New SqlClient.SqlConnection(connectionString)
        'Debug.Print(Me.CommandText)
    End Sub

    Public Overridable Function RunQuery() As List(Of Dictionary(Of String, String))
        RunQuery = New List(Of Dictionary(Of String, String))
        'Dim obj As T
        reader = InternalQuery()
        If reader Is Nothing Then
            Return Nothing
        End If
        If Not reader.HasRows Then Exit Function
        Do While reader.Read()
            Dim Line As Dictionary(Of String, String) = New Dictionary(Of String, String)
            For i As Integer = 0 To reader.FieldCount - 1
                Line.Add(reader.GetName(i), reader(i))
            Next
            RunQuery.Add(Line)
        Loop
        conn.Close()
        searchResults = RunQuery
        Return RunQuery
    End Function

    Protected Function InternalQuery() As SqlDataReader

        cmd = New SqlClient.SqlCommand(CommandText, conn) With {
            .CommandType = CommandType.Text,
            .CommandTimeout = 120
        }

        'Debug.Print(CommandText)
        'Debug.Print("")
        Do While conn.State <> ConnectionState.Open
            conn.Open()
        Loop

        '
        Try
            InternalQuery = cmd.ExecuteReader(CommandBehavior.CloseConnection)
        Catch ex As Exception
            Debug.Print(CommandText)
            Debug.Print(ex.Message)
            Debug.Print("")
            Debug.Print(ex.ToString())
            Debug.Print("")
            searchErrored = True
            Return Nothing
        End Try

        'conn.Close()
    End Function

    Public Function GetRow(index As Integer) As Dictionary(Of String, String)
        If index < searchResults.Count Then
            Return (searchResults(index))
        Else
            Throw New IndexOutOfRangeException
        End If
    End Function

    Public Function GetRowsFromColumns(names As List(Of String)) As List(Of Dictionary(Of String, String))
        Dim ret As New List(Of Dictionary(Of String, String))

        For i As Integer = 0 To searchResults.Count - 1 Step 1
            ret.Add(New Dictionary(Of String, String))
            For Each name In names
                ret(i).Add(name, searchResults(i)(name))
            Next
        Next
        Return ret
    End Function

    Public Function GetColumn(index As String) As List(Of String)
        GetColumn = New List(Of String)
        For Each list As Dictionary(Of String, String) In searchResults
            GetColumn.Add(list(index))
        Next
    End Function

    Public Function GetColumns(names As List(Of String)) As List(Of List(Of String))
        GetColumns = New List(Of List(Of String))
        For Each name In names
            GetColumns.Add(GetColumn(name))
        Next
    End Function

    Public Sub setTimeOut(input As Integer)
        _CmdTimeout = input
    End Sub

    Public ReadOnly Property CmdTimeout As Integer = 120

    Protected connectionString As String = "Data Source=gar-epdmdb01;Initial Catalog=PDM_Milbank;User ID=epdm_viewer;password=epdm_viewer1!" 'Integrated Security=True" ' CONNECTION STRING IS HERE
    Protected objReader As System.IO.StreamReader
    Protected reader As SqlClient.SqlDataReader
    Protected conn As SqlConnection
    Protected CommandText As String
    Protected cmd As SqlCommand

    Protected searchResults As List(Of Dictionary(Of String, String)) = New List(Of Dictionary(Of String, String))

    Protected SearchFile As String
    Protected HeadString As String = Nothing
    Protected TailString As String = Nothing

    Public Shared searchErrored As Boolean = False
End Class

