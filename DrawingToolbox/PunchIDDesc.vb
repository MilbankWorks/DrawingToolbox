Imports EPDM.Interop.epdm
Imports SolidWorks.Interop.swconst
Imports SolidWorks.Interop.sldworks

Public Class PunchIDDesc
    Public swFactory As Factories.SWFactory
    Public swApp As SldWorks


    Public Sub main()

        swFactory = New Factories.SWFactory
        swApp = swFactory.GetSwAppFromExisting()

        Dim ActiveDoc As ModelDoc2
        Dim DrawingDoc As DrawingDoc
        Dim tableAnn As TableAnnotation
        Dim punchTableAnn As PunchTableAnnotation
        Dim swView As View
        Dim PunchIDSQLSearch As New PunchIDSQLSearch

        ActiveDoc = swApp.ActiveDoc
        If ActiveDoc.GetType <> 3 Then Exit Sub 'swDocDrawing 

        DrawingDoc = ActiveDoc

        PunchIDSQLSearch.BuildTable()

        DrawingDoc.ActivateSheet(DrawingDoc.GetSheetNames(0))
        swView = DrawingDoc.GetFirstView()
        Do While Not swView Is Nothing
            tableAnn = swView.GetFirstTableAnnotation
            Do While Not tableAnn Is Nothing
                If tableAnn.Title = "Punch Table" Then
                    punchTableAnn = tableAnn
                    Exit Do
                End If
                tableAnn = tableAnn.GetNext
            Loop
            swView = swView.GetNextView
        Loop

        If punchTableAnn Is Nothing Then
            Exit Sub
        End If

        tableAnn = punchTableAnn

        If tableAnn.Text2(0, 2, True) <> "DESCRIPTION" Then
            tableAnn.InsertColumn2(2, tableAnn.ColumnCount - 1, "DESCRIPTION", 0) 'Add Desc Column
            tableAnn.SetColumnType2(tableAnn.ColumnCount - 2, 0, True) 'Force Desc type to Custom
            tableAnn.Text2(0, tableAnn.ColumnCount - 2, True) = "DESCRIPTION" 'Reset Desc title

            Dim colDescWidth As Double 'This is put here as a safeguard. Ultimately the column size will be driven by the PunchTableFormatUpdate VBA macro.
            colDescWidth = 0.065725 'in meters
            tableAnn.SetLockColumnWidth(tableAnn.ColumnCount - 2, False)
            tableAnn.SetColumnWidth(tableAnn.ColumnCount - 2, colDescWidth, 0) 'Resize Desc
            tableAnn.SetLockColumnWidth(tableAnn.ColumnCount - 2, True)
        End If

        For iter = 1 To tableAnn.RowCount
            tableAnn.Text2(iter, tableAnn.ColumnCount - 2, True) = PunchIDSQLSearch.TableLookUp(tableAnn.Text2(iter, 1, True)) 'Write Desc to the row
        Next
    End Sub
End Class
