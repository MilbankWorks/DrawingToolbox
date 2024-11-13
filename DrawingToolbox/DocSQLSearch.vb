Imports System.Collections.Generic
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.Threading
Imports System.Ling
Imports System.IO

Public Class DocSQLSearch
    Protected connectionString As String = "Data Source= GAR-EPDMDB01; Initial Catalog=PDM_Milbank;User ID =epdm_viewer;password=epdm_viewer1!" 'connection string including info about connection
    Protected objReader As System.IO.StreamReader
    Protected reader As SqlClient.SqlDataReader
    Protected conn As SqlClient.SqlConnection
    Protected cmd As SqlCommand

    Public Function QuickerTableLookup(InputString)
        'Generate default values for results
        Dim conisioLink As String = "Not Found"
        Dim FileNameExt As String = "Not Found"
        Dim FilePath As String = "Not Found"

        Dim docID As New List(Of Int32)()
        Dim docName As New List(Of String)()
        Dim projID As New List(Of Int16)()
        Dim projPath As New List(Of String)()

        conn = New System.Data.SqlClient.SqlConnection(connectionString)
        cmd = conn.CreateCommand

        'This query searches for the 1st drawing that has the name matches the PartNo InputString of the active model configuration
        Dim DocSQLSearchQuery As String
        DocSQLSearchQuery = "--This is the DrawingToolbox script to check for existing drawing
            SELECT TOP 1 DOC.DocumentID
            ,DOC.Filename
            ,PRJ.ProjectID
            ,CONCAT('C:\PDM_Milbank',PRJ.[Path]) [FullPath]
            --,CONCAT('conisio://PDM_Milbank/open?projectid=',PRJ.ProjectID,'&documentid=',DOC.DocumentID,'&objecttype=1') [DocLink]
            FROM PDM_Milbank.dbo.Documents DOC
            INNER JOIN(SELECT DocumentID, ProjectID FROM PDM_Milbank.dbo.DocumentsInProjects) DIP ON DOC.DocumentID = DIP.DocumentID
            INNER JOIN(SELECT ProjectID, [Path] FROM PDM_Milbank.dbo.Projects) PRJ ON DIP.ProjectID = PRJ.ProjectID
            WHERE ObjectTypeID = 1
            AND Deleted <> 1
            AND ExtensionID = 3
            AND REPLACE(DOC.Filename,'.SLDDRW','') = "
        DocSQLSearchQuery = DocSQLSearchQuery + "'" + CStr(CInt(Left(InputString, 7))) + "'"
        cmd.CommandText = DocSQLSearchQuery
        conn.Open()
        reader = cmd.ExecuteReader()
        Do While reader.Read()
            docID.Add(reader.GetValue(0))
            docName.Add(reader.GetValue(1))
            projID.Add(reader.GetValue(2))
            projPath.Add(reader.GetValue(3))
        Loop
        reader.Close()
        conn.Close()

        'If search result isn't empty, generate conisio link to quickly open the drawing and overwrite default values with search result
        If docID.Count <> 0 Then
            conisioLink = "conisio://PDM_Milbank/open?projectid=" & projID(0) & "&documentID=" & docID(0) & "&objecttype=1"
            FileNameExt = docName(0)
            FilePath = projPath(0)
        End If

        'Return results
        Dim output = New String() {conisioLink, FileNameExt, FilePath}
        Return output

    End Function
End Class
