﻿Imports SolidWorks.Interop.swconst
Imports SolidWorks.Interop.sldworks
Imports EPDM.Interop.epdm

Public Class DrawingToolbox

    Public swFactory As Factories.SWFactory
    Public swApp As SldWorks
    Public swVault As New EdmVault5
    Private PunchIDDesc As PunchIDDesc
    Private FindPartNo As FindPartNo
    Private swError As Integer

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
        chCreateNonFlat.Checked = False
        chCreateFlat.Checked = False

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
        chCreateNonFlat.Checked = False
        chCreateFlat.Checked = False

        chNonFlatModelItem.Checked = False
        chNonFlatBOMTable.Checked = False
        chNonFlatAutoBalloon.Checked = False
        chNonFlatRevTable.Checked = False
        chNonFlatReloadSheetFormat.Checked = False
        chNonFlatPDF.Checked = False

    End Sub

    Private Sub chModelNonFlat_CheckedChanged(sender As Object, e As EventArgs) Handles chCreateNonFlat.CheckedChanged
        If chCreateNonFlat.Checked Then
            chCreateFlat.Checked = False

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
    Private Sub chModelFlat_CheckedChanged(sender As Object, e As EventArgs) Handles chCreateFlat.CheckedChanged
        If chCreateFlat.Checked Then
            chCreateNonFlat.Checked = False

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

        'These are now in the MacroClass.vb
        'Dim macroRootPath As String = "C:\PDM_Milbank\SolidWorks Library\Macros\Drawings\"
        'Dim macroProcedure As String = "main"

        If Not swVault.IsLoggedIn Then swVault.LoginAuto("PDM_Milbank", 0)
        Dim mCreate_NonFlat As New MacroClass(swVault, "Create_NonFlat.swp", "Create_NonFlat1")
        Dim mCreate_Flat As New MacroClass(swVault, "Create_Flat.swp", "Create_Flat1")
        Dim mDrw_PDF As New MacroClass(swVault, "Drw_PDF.swp", "Drw_PDF1")
        Dim mDrw_RevTableFormat As New MacroClass(swVault, "Drw_RevTableFormat.swp", "Drw_RevTableFormat1")
        Dim mDrw_RevTableUpdate As New MacroClass(swVault, "Drw_RevTableUpdate.swp", "Drw_RevTableUpdate1")
        Dim mDrw_SheetFormat As New MacroClass(swVault, "Drw_SheetFormat.swp", "Drw_SheetFormat1")
        Dim mNonFlat_AutoBalloon As New MacroClass(swVault, "NonFlat_AutoBalloon.swp", "NonFlat_AutoBalloon1")
        Dim mNonFlat_BOMTable As New MacroClass(swVault, "NonFlat_BOMTable.swp", "NonFlat_BOMTable1")
        Dim mNonFlat_BOMTableFormat As New MacroClass(swVault, "NonFlat_BOMTableFormat.swp", "NonFlat_BOMTableFormat1")
        Dim mNonFlat_ModelItems As New MacroClass(swVault, "NonFlat_ModelItems.swp", "NonFlat_ModelItems1")
        Dim mFlat_OrdinateDims As New MacroClass(swVault, "Flat_OrdinateDims.swp", "Flat_OrdinateDims1")
        Dim mFlat_PunchTable As New MacroClass(swVault, "Flat_PunchTable.swp", "Flat_PunchTable1")
        Dim mFlat_PunchTableFormat As New MacroClass(swVault, "Flat_PunchTableFormat.swp", "Flat_PunchTableFormat1")

        swFactory = New Factories.SWFactory
        swApp = swFactory.GetSwAppFromExisting()

        If chCreateNonFlat.Checked Then
            swApp.RunMacro2(mCreate_NonFlat.RootPath + mCreate_NonFlat.Name, mCreate_NonFlat.Module, mCreate_NonFlat.Procedure, 1, swError)
            lblDisplayText.Text = lblDisplayText.Text + chCreateNonFlat.Text + vbCrLf
        End If
        If chNonFlatModelItem.Checked Then
            swApp.RunMacro2(mNonFlat_ModelItems.RootPath + mNonFlat_ModelItems.Name, mNonFlat_ModelItems.Module, mNonFlat_ModelItems.Procedure, 1, swError)
            lblDisplayText.Text = lblDisplayText.Text + chNonFlatModelItem.Text + vbCrLf
        End If
        If chNonFlatBOMTable.Checked Then
            swApp.RunMacro2(mNonFlat_BOMTable.RootPath + mNonFlat_BOMTable.Name, mNonFlat_BOMTable.Module, mNonFlat_BOMTable.Procedure, 1, swError)
            swApp.RunMacro2(mNonFlat_BOMTableFormat.RootPath + mNonFlat_BOMTableFormat.Name, mNonFlat_BOMTableFormat.Module, mNonFlat_BOMTableFormat.Procedure, 1, swError)
            lblDisplayText.Text = lblDisplayText.Text + chNonFlatBOMTable.Text + vbCrLf
        End If
        If chNonFlatAutoBalloon.Checked Then
            swApp.RunMacro2(mNonFlat_AutoBalloon.RootPath + mNonFlat_AutoBalloon.Name, mNonFlat_AutoBalloon.Module, mNonFlat_AutoBalloon.Procedure, 1, swError)
            lblDisplayText.Text = lblDisplayText.Text + chNonFlatAutoBalloon.Text + vbCrLf
        End If
        If chNonFlatRevTable.Checked Then
            swApp.RunMacro2(mDrw_RevTableUpdate.RootPath + mDrw_RevTableUpdate.Name, mDrw_RevTableUpdate.Module, mDrw_RevTableUpdate.Procedure, 1, swError)
            swApp.RunMacro2(mDrw_RevTableFormat.RootPath + mDrw_RevTableFormat.Name, mDrw_RevTableFormat.Module, mDrw_RevTableFormat.Procedure, 1, swError)
            lblDisplayText.Text = lblDisplayText.Text + chNonFlatRevTable.Text + vbCrLf
        End If
        If chNonFlatReloadSheetFormat.Checked Then
            swApp.RunMacro2(mDrw_SheetFormat.RootPath + mDrw_SheetFormat.Name, mDrw_SheetFormat.Module, mDrw_SheetFormat.Procedure, 1, swError)
            lblDisplayText.Text = lblDisplayText.Text + chNonFlatReloadSheetFormat.Text + vbCrLf
        End If
        If chNonFlatPDF.Checked Then
            swApp.RunMacro2(mDrw_PDF.RootPath + mDrw_PDF.Name, mDrw_PDF.Module, mDrw_PDF.Procedure, 1, swError)
            lblDisplayText.Text = lblDisplayText.Text + chNonFlatPDF.Text + vbCrLf
        End If
        If chCreateFlat.Checked Then
            swApp.RunMacro2(mCreate_Flat.RootPath + mCreate_Flat.Name, mCreate_Flat.Module, mCreate_Flat.Procedure, 1, swError)
            lblDisplayText.Text = lblDisplayText.Text + chCreateFlat.Text + vbCrLf
        End If
        If chFlatPunchTable.Checked Then
            swApp.RunMacro2(mFlat_PunchTable.RootPath + mFlat_PunchTable.Name, mFlat_PunchTable.Module, mFlat_PunchTable.Procedure, 1, swError)
            PunchIDDesc = New PunchIDDesc
            PunchIDDesc.main()
            swApp.RunMacro2(mFlat_PunchTableFormat.RootPath + mFlat_PunchTableFormat.Name, mFlat_PunchTableFormat.Module, mFlat_PunchTableFormat.Procedure, 1, swError)
            lblDisplayText.Text = lblDisplayText.Text + chFlatPunchTable.Text + vbCrLf
        End If
        If chFlatOrdDim.Checked Then
            swApp.RunMacro2(mFlat_OrdinateDims.RootPath + mFlat_OrdinateDims.Name, mFlat_OrdinateDims.Module, mFlat_OrdinateDims.Procedure, 1, swError)
            lblDisplayText.Text = lblDisplayText.Text + chFlatOrdDim.Text + vbCrLf
        End If
        If chFlatRevTable.Checked Then
            swApp.RunMacro2(mDrw_RevTableUpdate.RootPath + mDrw_RevTableUpdate.Name, mDrw_RevTableUpdate.Module, mDrw_RevTableUpdate.Procedure, 1, swError)
            swApp.RunMacro2(mDrw_RevTableFormat.RootPath + mDrw_RevTableFormat.Name, mDrw_RevTableFormat.Module, mDrw_RevTableFormat.Procedure, 1, swError)
            lblDisplayText.Text = lblDisplayText.Text + chFlatRevTable.Text + vbCrLf
        End If
        If chFlatReloadSheetFormat.Checked Then
            swApp.RunMacro2(mDrw_SheetFormat.RootPath + mDrw_SheetFormat.Name, mDrw_SheetFormat.Module, mDrw_SheetFormat.Procedure, 1, swError)
            lblDisplayText.Text = lblDisplayText.Text + chFlatReloadSheetFormat.Text + vbCrLf
        End If
        If chFlatPDF.Checked Then
            swApp.RunMacro2(mDrw_PDF.RootPath + mDrw_PDF.Name, mDrw_PDF.Module, mDrw_PDF.Procedure, 1, swError)
            lblDisplayText.Text = lblDisplayText.Text + chFlatPDF.Text + vbCrLf
        End If
        tbxStatus.Text = "Ready for next command"
    End Sub

    Private Sub BehaviorDrawingNonFlat()
        chCreateNonFlat.Checked = False
        chCreateFlat.Checked = False

        chFlatPunchTable.Checked = False
        chFlatOrdDim.Checked = False
        chFlatRevTable.Checked = False
        chFlatReloadSheetFormat.Checked = False
        chFlatPDF.Checked = False
    End Sub

    Private Sub BehaviorDrawingFlat()
        chCreateNonFlat.Checked = False
        chCreateFlat.Checked = False

        chNonFlatModelItem.Checked = False
        chNonFlatBOMTable.Checked = False
        chNonFlatAutoBalloon.Checked = False
        chNonFlatRevTable.Checked = False
        chNonFlatReloadSheetFormat.Checked = False
        chNonFlatPDF.Checked = False
    End Sub
End Class
