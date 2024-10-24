Imports SolidWorks.Interop.swconst
Imports SolidWorks.Interop.sldworks
Imports EPDM.Interop.epdm

Public Class DrawingToolbox

    Public swFactory As Factories.SWFactory
    Public swApp As SldWorks
    Private PunchIDDesc As PunchIDDesc
    Private FindPartNo As FindPartNo
    Private swError As Integer
    Private macroRootPath As String
    Private macroModelNonFlat As String, moduleModelNonFlat As String, procedureModelNonFlat As String
    Private macroModelFlat As String, moduleModelFlat As String, procedureModelFlat As String
    Private macroNonFlatModelItem As String, moduleNonFlatModelItem As String, procedureNonFlatModelItem As String
    Private macroNonFlatBOMTable As String, moduleNonFlatBOMTable As String, procedureNonFlatBOMTable As String
    Private macroNonFlatBOMTableFormat As String, moduleNonFlatBOMTableFormat As String, procedureNonFlatBOMTableFormat As String
    Private macroNonFlatAutoBalloon As String, moduleNonFlatAutoBalloon As String, procedureNonFlatAutoBalloon As String
    Private macroNonFlatRevTable As String, moduleNonFlatRevTable As String, procedureNonFlatRevTable As String
    Private macroNonFlatRevTableFormat As String, moduleNonFlatRevTableFormat As String, procedureNonFlatRevTableFormat As String
    Private macroNonFlatReloadSheetFormat As String, moduleNonFlatReloadSheetFormat As String, procedureNonFlatReloadSheetFormat As String
    Private macroNonFlatPDF As String, moduleNonFlatPDF As String, procedureNonFlatPDF As String
    Private macroFlatPunchTable As String, moduleFlatPunchTable As String, procedureFlatPunchTable As String
    Private macroFlatPunchTableFormat As String, moduleFlatPunchTableFormat As String, procedureFlatPunchTableFormat As String
    Private macroFlatOrdDim As String, moduleFlatOrdDim As String, procedureFlatOrdDim As String
    Private macroFlatRevTable As String, moduleFlatRevTable As String, procedureFlatRevTable As String
    Private macroFlatRevTableFormat As String, moduleFlatRevTableFormat As String, procedureFlatRevTableFormat As String
    Private macroFlatReloadSheetFormat As String, moduleFlatReloadSheetFormat As String, procedureFlatReloadSheetFormat As String
    Private macroFlatPDF As String, moduleFlatPDF As String, procedureFlatPDF As String

    Private Sub btnSelectAll1_Click(sender As Object, e As EventArgs) Handles btnSelectAll1.Click
        If chNonFlatModelItem.Checked = False And chNonFlatBOMTable.Checked = False And chNonFlatAutoBalloon.Checked = False And chNonFlatRevTable.Checked = False And chNonFlatReloadSheetFormat.Checked = False And chNonFlatPDF.Checked = False Then
            'Radio buttons seemingly have a bug. If radio button is checked, and a different checkbox event is set to clear the radio button, checkbox will have to be clicked twice to actually set as True
            'Radio buttons require checkbox.checked = True twice to actually change
            chNonFlatModelItem.Checked = True
            chNonFlatBOMTable.Checked = True
            chNonFlatAutoBalloon.Checked = True
            chNonFlatRevTable.Checked = True
            chNonFlatReloadSheetFormat.Checked = True
            chNonFlatPDF.Checked = True
        Else
            chNonFlatModelItem.Checked = False
            chNonFlatBOMTable.Checked = False
            chNonFlatAutoBalloon.Checked = False
            chNonFlatRevTable.Checked = False
            chNonFlatReloadSheetFormat.Checked = False
            chNonFlatPDF.Checked = False
        End If
        chModelCreateNonFlat.Checked = False
        chModelCreateFlat.Checked = False

        chFlatPunchTable.Checked = False
        chFlatOrdDim.Checked = False
        chFlatRevTable.Checked = False
        chFlatReloadSheetFormat.Checked = False
        chFlatPDF.Checked = False
    End Sub

    Private Sub btnSelectAll2_Click(sender As Object, e As EventArgs) Handles btnSelectAll2.Click
        If chFlatPunchTable.Checked = False And chFlatOrdDim.Checked = False And chFlatRevTable.Checked = False And chFlatReloadSheetFormat.Checked = False And chFlatPDF.Checked = False Then
            'Radio buttons seemingly have a bug. If radio button is checked, and a different checkbox event is set to clear the radio button, checkbox will have to be clicked twice to actually set as True
            'Radio buttons require checkbox.checked = True twice to actually change
            chFlatPunchTable.Checked = True
            chFlatOrdDim.Checked = True
            chFlatRevTable.Checked = True
            chFlatReloadSheetFormat.Checked = True
            chFlatPDF.Checked = True
        Else
            chFlatPunchTable.Checked = False
            chFlatOrdDim.Checked = False
            chFlatRevTable.Checked = False
            chFlatReloadSheetFormat.Checked = False
            chFlatPDF.Checked = False
        End If
        chModelCreateNonFlat.Checked = False
        chModelCreateFlat.Checked = False

        chNonFlatModelItem.Checked = False
        chNonFlatBOMTable.Checked = False
        chNonFlatAutoBalloon.Checked = False
        chNonFlatRevTable.Checked = False
        chNonFlatReloadSheetFormat.Checked = False
        chNonFlatPDF.Checked = False

    End Sub

    Private Sub chModelNonFlat_CheckedChanged(sender As Object, e As EventArgs) Handles chModelCreateNonFlat.CheckedChanged
        If chModelCreateNonFlat.Checked Then
            chModelCreateFlat.Checked = False

            chNonFlatModelItem.Checked = False
            chNonFlatBOMTable.Checked = False
            chNonFlatAutoBalloon.Checked = False
            chNonFlatRevTable.Checked = False
            chNonFlatReloadSheetFormat.Checked = False
            chNonFlatPDF.Checked = False

            chFlatPunchTable.Checked = False
            chFlatOrdDim.Checked = False
            chFlatRevTable.Checked = False
            chFlatReloadSheetFormat.Checked = False
            chFlatPDF.Checked = False
        End If
    End Sub
    Private Sub chModelFlat_CheckedChanged(sender As Object, e As EventArgs) Handles chModelCreateFlat.CheckedChanged
        If chModelCreateFlat.Checked Then
            chModelCreateNonFlat.Checked = False

            chNonFlatModelItem.Checked = False
            chNonFlatBOMTable.Checked = False
            chNonFlatAutoBalloon.Checked = False
            chNonFlatRevTable.Checked = False
            chNonFlatReloadSheetFormat.Checked = False
            chNonFlatPDF.Checked = False

            chFlatPunchTable.Checked = False
            chFlatOrdDim.Checked = False
            chFlatRevTable.Checked = False
            chFlatReloadSheetFormat.Checked = False
            chFlatPDF.Checked = False
        End If
    End Sub
    Private Sub chNonFlatModelItem_CheckedChanged(sender As Object, e As EventArgs) Handles chNonFlatModelItem.CheckedChanged
        If chNonFlatModelItem.Checked Then
            BehaviorDrawingNonFlat()
        End If
    End Sub
    Private Sub chNonFlatBOMTable_CheckedChanged(sender As Object, e As EventArgs) Handles chNonFlatBOMTable.CheckedChanged
        If chNonFlatBOMTable.Checked Then
            BehaviorDrawingNonFlat()
        End If
    End Sub
    Private Sub chNonFlatAutoBalloon_CheckedChanged(sender As Object, e As EventArgs) Handles chNonFlatAutoBalloon.CheckedChanged
        If chNonFlatAutoBalloon.Checked Then
            BehaviorDrawingNonFlat()
        End If
    End Sub
    Private Sub chNonFlatRevTable_CheckedChanged(sender As Object, e As EventArgs) Handles chNonFlatRevTable.CheckedChanged
        If chNonFlatRevTable.Checked Then
            BehaviorDrawingNonFlat()
        End If
    End Sub
    Private Sub chNonFlatReloadSheetFormat_CheckedChanged(sender As Object, e As EventArgs) Handles chNonFlatReloadSheetFormat.CheckedChanged
        If chNonFlatReloadSheetFormat.Checked Then
            BehaviorDrawingNonFlat()
        End If
    End Sub

    Private Sub chNonFlatPDF_CheckedChanged(sender As Object, e As EventArgs) Handles chNonFlatPDF.CheckedChanged
        If chNonFlatPDF.Checked Then
            BehaviorDrawingNonFlat()
        End If
    End Sub
    Private Sub chFlatPunchTable_CheckedChanged(sender As Object, e As EventArgs) Handles chFlatPunchTable.CheckedChanged
        If chFlatPunchTable.Checked Then
            BehaviorDrawingFlat()
        End If
    End Sub
    Private Sub chFlatOrdDim_CheckedChanged(sender As Object, e As EventArgs) Handles chFlatOrdDim.CheckedChanged
        If chFlatOrdDim.Checked Then
            BehaviorDrawingFlat()
        End If
    End Sub
    Private Sub chFlatRevTable_CheckedChanged(sender As Object, e As EventArgs) Handles chFlatRevTable.CheckedChanged
        If chFlatRevTable.Checked Then
            BehaviorDrawingFlat()
        End If
    End Sub
    Private Sub chFlatReloadSheetFormat_CheckedChanged(sender As Object, e As EventArgs) Handles chFlatReloadSheetFormat.CheckedChanged
        If chFlatReloadSheetFormat.Checked Then
            BehaviorDrawingFlat()
        End If
    End Sub

    Private Sub chFlatPDF_CheckedChanged(sender As Object, e As EventArgs) Handles chFlatPDF.CheckedChanged
        If chFlatPDF.Checked Then
            BehaviorDrawingFlat()
        End If
    End Sub
    Private Sub btnCheckExists_Click(sender As Object, e As EventArgs) Handles btnCheckExists.Click
        tbxStatus.Text = "Running"
        FindPartNo = New FindPartNo
        FindPartNo.FindPN()
        tbxStatus.Text = "Ready for next command"
    End Sub

    Private Sub btnRunMacros_Click(sender As Object, e As EventArgs) Handles btnRunMacros.Click

        tbxStatus.Text = "Running"
        lblDisplayText.Text = ""

        macroRootPath = "C:\PDM_Milbank\SolidWorks Library\Macros\Drawings\"

        macroModelNonFlat = "Create_NonFlat.swp"
        moduleModelNonFlat = "Create_NonFlat1"
        procedureModelNonFlat = "main"

        macroNonFlatModelItem = "NonFlat_ModelItems.swp"
        moduleNonFlatModelItem = "NonFlat_ModelItems1"
        procedureNonFlatModelItem = "main"

        macroNonFlatBOMTable = "NonFlat_BOMTable.swp"
        moduleNonFlatBOMTable = "NonFlat_BOMTable1"
        procedureNonFlatBOMTable = "main"

        macroNonFlatBOMTableFormat = "NonFlat_BOMTableFormat.swp"
        moduleNonFlatBOMTableFormat = "NonFlat_BOMTableFormat1"
        procedureNonFlatBOMTableFormat = "main"

        macroNonFlatAutoBalloon = "NonFlat_AutoBalloon.swp"
        moduleNonFlatAutoBalloon = "NonFlat_AutoBalloon1"
        procedureNonFlatAutoBalloon = "main"

        macroNonFlatRevTable = "Drw_RevTableUpdate.swp"
        moduleNonFlatRevTable = "Drw_RevTableUpdate1"
        procedureNonFlatRevTable = "main"

        macroNonFlatRevTableFormat = "Drw_RevTableFormat.swp"
        moduleNonFlatRevTableFormat = "Drw_RevTableFormat1"
        procedureNonFlatRevTableFormat = "main"

        macroNonFlatReloadSheetFormat = "Drw_SheetFormat.swp"
        moduleNonFlatReloadSheetFormat = "Drw_SheetFormat1"
        procedureNonFlatReloadSheetFormat = "main"

        macroNonFlatPDF = "Drw_PDF.swp"
        moduleNonFlatPDF = "Drw_PDF1"
        procedureNonFlatPDF = "main"

        macroModelFlat = "Create_Flat.swp"
        moduleModelFlat = "Create_Flat1"
        procedureModelFlat = "main"

        macroFlatPunchTable = "Flat_PunchTable.swp"
        moduleFlatPunchTable = "Flat_PunchTable1"
        procedureFlatPunchTable = "main"

        macroFlatPunchTableFormat = "Flat_PunchTableFormat.swp"
        moduleFlatPunchTableFormat = "Flat_PunchTableFormat1"
        procedureFlatPunchTableFormat = "main"

        macroFlatOrdDim = "Flat_OrdinateDims.swp"
        moduleFlatOrdDim = "Flat_OrdinateDims1"
        procedureFlatOrdDim = "main"

        macroFlatRevTable = macroNonFlatRevTable
        moduleFlatRevTable = moduleNonFlatRevTable
        procedureFlatRevTable = procedureNonFlatRevTable

        macroFlatRevTableFormat = macroNonFlatRevTableFormat
        moduleFlatRevTableFormat = moduleNonFlatRevTableFormat
        procedureFlatRevTableFormat = procedureNonFlatRevTableFormat

        macroFlatPDF = macroNonFlatPDF
        moduleFlatPDF = moduleNonFlatPDF
        procedureFlatPDF = procedureNonFlatPDF

        macroFlatReloadSheetFormat = macroNonFlatReloadSheetFormat
        moduleFlatReloadSheetFormat = moduleNonFlatReloadSheetFormat
        procedureFlatReloadSheetFormat = procedureNonFlatPDF

        swFactory = New Factories.SWFactory
        swApp = swFactory.GetSwAppFromExisting()

        If chModelCreateNonFlat.Checked Then
            swApp.RunMacro2(macroRootPath + macroModelNonFlat, moduleModelNonFlat, procedureModelNonFlat, 1, swError)
            lblDisplayText.Text = lblDisplayText.Text + chModelCreateNonFlat.Text + vbCrLf
        End If
        If chNonFlatModelItem.Checked Then
            swApp.RunMacro2(macroRootPath + macroNonFlatModelItem, moduleNonFlatModelItem, procedureNonFlatModelItem, 1, swError)
            lblDisplayText.Text = lblDisplayText.Text + chNonFlatModelItem.Text + vbCrLf
        End If
        If chNonFlatBOMTable.Checked Then
            swApp.RunMacro2(macroRootPath + macroNonFlatBOMTable, moduleNonFlatBOMTable, procedureNonFlatBOMTable, 1, swError)
            swApp.RunMacro2(macroRootPath + macroNonFlatBOMTableFormat, moduleNonFlatBOMTableFormat, procedureNonFlatBOMTableFormat, 1, swError)
            lblDisplayText.Text = lblDisplayText.Text + chNonFlatBOMTable.Text + vbCrLf
        End If
        If chNonFlatAutoBalloon.Checked Then
            swApp.RunMacro2(macroRootPath + macroNonFlatAutoBalloon, moduleNonFlatAutoBalloon, procedureNonFlatAutoBalloon, 1, swError)
            lblDisplayText.Text = lblDisplayText.Text + chNonFlatAutoBalloon.Text + vbCrLf
        End If
        If chNonFlatRevTable.Checked Then
            swApp.RunMacro2(macroRootPath + macroNonFlatRevTable, moduleNonFlatRevTable, procedureNonFlatRevTable, 1, swError)
            swApp.RunMacro2(macroRootPath + macroNonFlatRevTableFormat, moduleNonFlatRevTableFormat, procedureNonFlatRevTableFormat, 1, swError)
            lblDisplayText.Text = lblDisplayText.Text + chNonFlatRevTable.Text + vbCrLf
        End If
        If chNonFlatReloadSheetFormat.Checked Then
            swApp.RunMacro2(macroRootPath + macroNonFlatReloadSheetFormat, moduleNonFlatReloadSheetFormat, procedureNonFlatReloadSheetFormat, 1, swError)
            lblDisplayText.Text = lblDisplayText.Text + chNonFlatReloadSheetFormat.Text + vbCrLf
        End If
        If chNonFlatPDF.Checked Then
            swApp.RunMacro2(macroRootPath + macroNonFlatPDF, moduleNonFlatPDF, procedureNonFlatPDF, 1, swError)
            lblDisplayText.Text = lblDisplayText.Text + chNonFlatPDF.Text + vbCrLf
        End If
        If chModelCreateFlat.Checked Then
            swApp.RunMacro2(macroRootPath + macroModelFlat, moduleModelFlat, procedureModelFlat, 1, swError)
            lblDisplayText.Text = lblDisplayText.Text + chModelCreateFlat.Text + vbCrLf
        End If
        If chFlatPunchTable.Checked Then
            swApp.RunMacro2(macroRootPath + macroFlatPunchTable, moduleFlatPunchTable, procedureFlatPunchTable, 1, swError)
            PunchIDDesc = New PunchIDDesc
            PunchIDDesc.main()
            swApp.RunMacro2(macroRootPath + macroFlatPunchTableFormat, moduleFlatPunchTableFormat, procedureFlatPunchTableFormat, 1, swError)
            lblDisplayText.Text = lblDisplayText.Text + chFlatPunchTable.Text + vbCrLf
        End If
        If chFlatOrdDim.Checked Then
            swApp.RunMacro2(macroRootPath + macroFlatOrdDim, moduleFlatOrdDim, procedureFlatOrdDim, 1, swError)
            lblDisplayText.Text = lblDisplayText.Text + chFlatOrdDim.Text + vbCrLf
        End If
        If chFlatRevTable.Checked Then
            swApp.RunMacro2(macroRootPath + macroFlatRevTable, moduleFlatRevTable, procedureFlatRevTable, 1, swError)
            swApp.RunMacro2(macroRootPath + macroFlatRevTableFormat, moduleFlatRevTableFormat, procedureFlatRevTableFormat, 1, swError)
            lblDisplayText.Text = lblDisplayText.Text + chFlatRevTable.Text + vbCrLf
        End If
        If chFlatReloadSheetFormat.Checked Then
            swApp.RunMacro2(macroRootPath + macroFlatReloadSheetFormat, moduleFlatReloadSheetFormat, procedureFlatReloadSheetFormat, 1, swError)
            lblDisplayText.Text = lblDisplayText.Text + chFlatReloadSheetFormat.Text + vbCrLf
        End If
        If chFlatPDF.Checked Then
            swApp.RunMacro2(macroRootPath + macroFlatPDF, moduleFlatPDF, procedureFlatPDF, 1, swError)
            lblDisplayText.Text = lblDisplayText.Text + chFlatPDF.Text + vbCrLf
        End If
        tbxStatus.Text = "Ready for next command"
    End Sub

    Private Sub BehaviorDrawingNonFlat()
        chModelCreateNonFlat.Checked = False
        chModelCreateFlat.Checked = False

        chFlatPunchTable.Checked = False
        chFlatOrdDim.Checked = False
        chFlatRevTable.Checked = False
        chFlatReloadSheetFormat.Checked = False
        chFlatPDF.Checked = False
    End Sub

    Private Sub BehaviorDrawingFlat()
        chModelCreateNonFlat.Checked = False
        chModelCreateFlat.Checked = False

        chNonFlatModelItem.Checked = False
        chNonFlatBOMTable.Checked = False
        chNonFlatAutoBalloon.Checked = False
        chNonFlatRevTable.Checked = False
        chNonFlatReloadSheetFormat.Checked = False
        chNonFlatPDF.Checked = False
    End Sub
End Class
