Imports EPDM.Interop.epdm
Imports SolidWorks.Interop.sldworks

Public Class FindPartNo

    Public swFactory As Factories.SWFactory
    Public swApp As SldWorks
    Private swError As Integer
    Private ActiveDoc As ModelDoc2
    Private swConfig As Configuration
    Private swPropMgr As CustomPropertyManager
    Private PNVal As String

    Public Function FindPN()

        swFactory = New Factories.SWFactory
        swApp = swFactory.GetSwAppFromExisting() 'Get Active SOLIDWORKS application
        If swApp Is Nothing Then Exit Function
        ActiveDoc = swApp.ActiveDoc
        If ActiveDoc Is Nothing Then Exit Function
        If ActiveDoc.GetType <> 1 And ActiveDoc.GetType <> 2 Then 'If active document is not swDocPart or swDocAssembly then exit
            Exit Function
        End If

        swPropMgr = ActiveDoc.GetActiveConfiguration.CustomPropertyManager
        Dim DocSQLSearch As New DocSQLSearch
        'DocSQLSearch.BuildTable()

        swPropMgr.Get6("PartNo", False, Nothing, PNVal, True, True) 'Get PartNo custom property for active configuration

        Dim output() As String
        Dim conisioLink As String
        Dim FileName As String
        Dim FilePath As String
        conisioLink = "Not Found"
        If Len(PNVal) = 7 And IsNumeric(PNVal) Then 'PartNo must be 7 digit and numeric number
            'output = DocSQLSearch.TableLookUp(PNVal)
            output = DocSQLSearch.QuickerTableLookup(PNVal)
            conisioLink = output(0)
            FileName = output(1)
            FilePath = output(2)
        End If

        If conisioLink <> "Not Found" Then 'Give user option to open drawing
            If MsgBox("Found drawing " + vbCrLf + vbCrLf + FileName + vbCrLf + vbCrLf + "at location " + vbCrLf + vbCrLf + FilePath + vbCrLf + vbCrLf + "Do you want to open this drawing?", vbYesNo, "Check Exist Active Config") = vbYes Then
                System.Diagnostics.Process.Start(conisioLink)
            End If
        Else
            MsgBox("No drawing with matching name found.")
        End If
        Return conisioLink
    End Function
End Class
