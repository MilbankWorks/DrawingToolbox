<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class DrawingToolbox
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.btnRunMacros = New System.Windows.Forms.Button()
        Me.chCreateNonFlat = New System.Windows.Forms.CheckBox()
        Me.chNonFlatModelItem = New System.Windows.Forms.CheckBox()
        Me.chNonFlatBOMTable = New System.Windows.Forms.CheckBox()
        Me.chNonFlatAutoBalloon = New System.Windows.Forms.CheckBox()
        Me.lblDisplayText = New System.Windows.Forms.Label()
        Me.chNonFlatRevTable = New System.Windows.Forms.CheckBox()
        Me.chCreateFlat = New System.Windows.Forms.CheckBox()
        Me.chFlatPunchTable = New System.Windows.Forms.CheckBox()
        Me.chFlatOrdDim = New System.Windows.Forms.CheckBox()
        Me.chFlatRevTable = New System.Windows.Forms.CheckBox()
        Me.lblProcedureList = New System.Windows.Forms.Label()
        Me.lblVersionText = New System.Windows.Forms.Label()
        Me.lblExplanation1 = New System.Windows.Forms.Label()
        Me.grbxNonFlat = New System.Windows.Forms.GroupBox()
        Me.chNonFlatReloadSheetFormat = New System.Windows.Forms.CheckBox()
        Me.chNonFlatPDF = New System.Windows.Forms.CheckBox()
        Me.btnSelectAll1 = New System.Windows.Forms.Button()
        Me.grbxFlat = New System.Windows.Forms.GroupBox()
        Me.chFlatReloadSheetFormat = New System.Windows.Forms.CheckBox()
        Me.chFlatPDF = New System.Windows.Forms.CheckBox()
        Me.btnSelectAll2 = New System.Windows.Forms.Button()
        Me.lblExplanation2 = New System.Windows.Forms.Label()
        Me.btnCheckExists = New System.Windows.Forms.Button()
        Me.grbxCreate = New System.Windows.Forms.GroupBox()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.tbxStatus = New System.Windows.Forms.TextBox()
        Me.grbxNonFlat.SuspendLayout()
        Me.grbxFlat.SuspendLayout()
        Me.grbxCreate.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnRunMacros
        '
        Me.btnRunMacros.Location = New System.Drawing.Point(491, 12)
        Me.btnRunMacros.Name = "btnRunMacros"
        Me.btnRunMacros.Size = New System.Drawing.Size(158, 56)
        Me.btnRunMacros.TabIndex = 0
        Me.btnRunMacros.Text = "Run Selected Macros"
        Me.btnRunMacros.UseVisualStyleBackColor = True
        '
        'chCreateNonFlat
        '
        Me.chCreateNonFlat.AutoSize = True
        Me.chCreateNonFlat.Location = New System.Drawing.Point(16, 62)
        Me.chCreateNonFlat.Name = "chCreateNonFlat"
        Me.chCreateNonFlat.Size = New System.Drawing.Size(186, 17)
        Me.chCreateNonFlat.TabIndex = 1
        Me.chCreateNonFlat.Text = "Create Formed/Assembly Drawing"
        Me.chCreateNonFlat.UseVisualStyleBackColor = True
        '
        'chNonFlatModelItem
        '
        Me.chNonFlatModelItem.AutoSize = True
        Me.chNonFlatModelItem.Location = New System.Drawing.Point(16, 19)
        Me.chNonFlatModelItem.Name = "chNonFlatModelItem"
        Me.chNonFlatModelItem.Size = New System.Drawing.Size(87, 17)
        Me.chNonFlatModelItem.TabIndex = 2
        Me.chNonFlatModelItem.Text = "Model Items*"
        Me.chNonFlatModelItem.UseVisualStyleBackColor = True
        '
        'chNonFlatBOMTable
        '
        Me.chNonFlatBOMTable.AutoSize = True
        Me.chNonFlatBOMTable.Location = New System.Drawing.Point(16, 41)
        Me.chNonFlatBOMTable.Name = "chNonFlatBOMTable"
        Me.chNonFlatBOMTable.Size = New System.Drawing.Size(80, 17)
        Me.chNonFlatBOMTable.TabIndex = 3
        Me.chNonFlatBOMTable.Text = "BOM Table"
        Me.chNonFlatBOMTable.UseVisualStyleBackColor = True
        '
        'chNonFlatAutoBalloon
        '
        Me.chNonFlatAutoBalloon.AutoSize = True
        Me.chNonFlatAutoBalloon.Location = New System.Drawing.Point(16, 62)
        Me.chNonFlatAutoBalloon.Name = "chNonFlatAutoBalloon"
        Me.chNonFlatAutoBalloon.Size = New System.Drawing.Size(86, 17)
        Me.chNonFlatAutoBalloon.TabIndex = 4
        Me.chNonFlatAutoBalloon.Text = "Auto Balloon"
        Me.chNonFlatAutoBalloon.UseVisualStyleBackColor = True
        '
        'lblDisplayText
        '
        Me.lblDisplayText.AutoSize = True
        Me.lblDisplayText.Location = New System.Drawing.Point(488, 165)
        Me.lblDisplayText.Name = "lblDisplayText"
        Me.lblDisplayText.Size = New System.Drawing.Size(27, 13)
        Me.lblDisplayText.TabIndex = 5
        Me.lblDisplayText.Text = "N/A"
        '
        'chNonFlatRevTable
        '
        Me.chNonFlatRevTable.AutoSize = True
        Me.chNonFlatRevTable.Location = New System.Drawing.Point(16, 85)
        Me.chNonFlatRevTable.Name = "chNonFlatRevTable"
        Me.chNonFlatRevTable.Size = New System.Drawing.Size(76, 17)
        Me.chNonFlatRevTable.TabIndex = 6
        Me.chNonFlatRevTable.Text = "Rev Table"
        Me.chNonFlatRevTable.UseVisualStyleBackColor = True
        '
        'chCreateFlat
        '
        Me.chCreateFlat.AutoSize = True
        Me.chCreateFlat.Location = New System.Drawing.Point(16, 85)
        Me.chCreateFlat.Name = "chCreateFlat"
        Me.chCreateFlat.Size = New System.Drawing.Size(119, 17)
        Me.chCreateFlat.TabIndex = 7
        Me.chCreateFlat.Text = "Create Flat Drawing"
        Me.chCreateFlat.UseVisualStyleBackColor = True
        '
        'chFlatPunchTable
        '
        Me.chFlatPunchTable.AutoSize = True
        Me.chFlatPunchTable.Location = New System.Drawing.Point(16, 19)
        Me.chFlatPunchTable.Name = "chFlatPunchTable"
        Me.chFlatPunchTable.Size = New System.Drawing.Size(91, 17)
        Me.chFlatPunchTable.TabIndex = 8
        Me.chFlatPunchTable.Text = "Punch Table*"
        Me.chFlatPunchTable.UseVisualStyleBackColor = True
        '
        'chFlatOrdDim
        '
        Me.chFlatOrdDim.AutoSize = True
        Me.chFlatOrdDim.Location = New System.Drawing.Point(16, 42)
        Me.chFlatOrdDim.Name = "chFlatOrdDim"
        Me.chFlatOrdDim.Size = New System.Drawing.Size(127, 17)
        Me.chFlatOrdDim.TabIndex = 9
        Me.chFlatOrdDim.Text = "Ordinate Dimensions*"
        Me.chFlatOrdDim.UseVisualStyleBackColor = True
        '
        'chFlatRevTable
        '
        Me.chFlatRevTable.AutoSize = True
        Me.chFlatRevTable.Location = New System.Drawing.Point(16, 65)
        Me.chFlatRevTable.Name = "chFlatRevTable"
        Me.chFlatRevTable.Size = New System.Drawing.Size(76, 17)
        Me.chFlatRevTable.TabIndex = 10
        Me.chFlatRevTable.Text = "Rev Table"
        Me.chFlatRevTable.UseVisualStyleBackColor = True
        '
        'lblProcedureList
        '
        Me.lblProcedureList.AutoSize = True
        Me.lblProcedureList.Location = New System.Drawing.Point(488, 152)
        Me.lblProcedureList.Name = "lblProcedureList"
        Me.lblProcedureList.Size = New System.Drawing.Size(98, 13)
        Me.lblProcedureList.TabIndex = 11
        Me.lblProcedureList.Text = "Macros Completed:"
        '
        'lblVersionText
        '
        Me.lblVersionText.AutoSize = True
        Me.lblVersionText.Location = New System.Drawing.Point(580, 333)
        Me.lblVersionText.Name = "lblVersionText"
        Me.lblVersionText.Size = New System.Drawing.Size(69, 13)
        Me.lblVersionText.TabIndex = 12
        Me.lblVersionText.Text = "Version 0.9.2"
        '
        'lblExplanation1
        '
        Me.lblExplanation1.AutoSize = True
        Me.lblExplanation1.Location = New System.Drawing.Point(9, 333)
        Me.lblExplanation1.Name = "lblExplanation1"
        Me.lblExplanation1.Size = New System.Drawing.Size(351, 13)
        Me.lblExplanation1.TabIndex = 13
        Me.lblExplanation1.Text = "Macros marked by a * can take a long time based on complexity of model"
        '
        'grbxNonFlat
        '
        Me.grbxNonFlat.BackColor = System.Drawing.SystemColors.Info
        Me.grbxNonFlat.Controls.Add(Me.chNonFlatReloadSheetFormat)
        Me.grbxNonFlat.Controls.Add(Me.chNonFlatPDF)
        Me.grbxNonFlat.Controls.Add(Me.btnSelectAll1)
        Me.grbxNonFlat.Controls.Add(Me.chNonFlatModelItem)
        Me.grbxNonFlat.Controls.Add(Me.chNonFlatBOMTable)
        Me.grbxNonFlat.Controls.Add(Me.chNonFlatAutoBalloon)
        Me.grbxNonFlat.Controls.Add(Me.chNonFlatRevTable)
        Me.grbxNonFlat.Location = New System.Drawing.Point(229, 12)
        Me.grbxNonFlat.Name = "grbxNonFlat"
        Me.grbxNonFlat.Size = New System.Drawing.Size(253, 159)
        Me.grbxNonFlat.TabIndex = 16
        Me.grbxNonFlat.TabStop = False
        Me.grbxNonFlat.Text = "Edit Active Formed/Assembly Drawing"
        '
        'chNonFlatReloadSheetFormat
        '
        Me.chNonFlatReloadSheetFormat.AutoSize = True
        Me.chNonFlatReloadSheetFormat.Location = New System.Drawing.Point(16, 108)
        Me.chNonFlatReloadSheetFormat.Name = "chNonFlatReloadSheetFormat"
        Me.chNonFlatReloadSheetFormat.Size = New System.Drawing.Size(126, 17)
        Me.chNonFlatReloadSheetFormat.TabIndex = 26
        Me.chNonFlatReloadSheetFormat.Text = "Reload Sheet Format"
        Me.chNonFlatReloadSheetFormat.UseVisualStyleBackColor = True
        '
        'chNonFlatPDF
        '
        Me.chNonFlatPDF.AutoSize = True
        Me.chNonFlatPDF.Location = New System.Drawing.Point(16, 131)
        Me.chNonFlatPDF.Name = "chNonFlatPDF"
        Me.chNonFlatPDF.Size = New System.Drawing.Size(121, 17)
        Me.chNonFlatPDF.TabIndex = 24
        Me.chNonFlatPDF.Text = "Create/Update PDF"
        Me.chNonFlatPDF.UseVisualStyleBackColor = True
        '
        'btnSelectAll1
        '
        Me.btnSelectAll1.Location = New System.Drawing.Point(149, 19)
        Me.btnSelectAll1.Name = "btnSelectAll1"
        Me.btnSelectAll1.Size = New System.Drawing.Size(89, 37)
        Me.btnSelectAll1.TabIndex = 23
        Me.btnSelectAll1.Text = "(De)Select All 1"
        Me.btnSelectAll1.UseVisualStyleBackColor = True
        '
        'grbxFlat
        '
        Me.grbxFlat.BackColor = System.Drawing.SystemColors.Info
        Me.grbxFlat.Controls.Add(Me.chFlatReloadSheetFormat)
        Me.grbxFlat.Controls.Add(Me.chFlatPDF)
        Me.grbxFlat.Controls.Add(Me.btnSelectAll2)
        Me.grbxFlat.Controls.Add(Me.chFlatPunchTable)
        Me.grbxFlat.Controls.Add(Me.chFlatOrdDim)
        Me.grbxFlat.Controls.Add(Me.chFlatRevTable)
        Me.grbxFlat.Location = New System.Drawing.Point(229, 177)
        Me.grbxFlat.Name = "grbxFlat"
        Me.grbxFlat.Size = New System.Drawing.Size(253, 139)
        Me.grbxFlat.TabIndex = 17
        Me.grbxFlat.TabStop = False
        Me.grbxFlat.Text = "Edit Active Flat Drawing"
        '
        'chFlatReloadSheetFormat
        '
        Me.chFlatReloadSheetFormat.AutoSize = True
        Me.chFlatReloadSheetFormat.Location = New System.Drawing.Point(16, 88)
        Me.chFlatReloadSheetFormat.Name = "chFlatReloadSheetFormat"
        Me.chFlatReloadSheetFormat.Size = New System.Drawing.Size(126, 17)
        Me.chFlatReloadSheetFormat.TabIndex = 27
        Me.chFlatReloadSheetFormat.Text = "Reload Sheet Format"
        Me.chFlatReloadSheetFormat.UseVisualStyleBackColor = True
        '
        'chFlatPDF
        '
        Me.chFlatPDF.AutoSize = True
        Me.chFlatPDF.Location = New System.Drawing.Point(16, 111)
        Me.chFlatPDF.Name = "chFlatPDF"
        Me.chFlatPDF.Size = New System.Drawing.Size(121, 17)
        Me.chFlatPDF.TabIndex = 23
        Me.chFlatPDF.Text = "Create/Update PDF"
        Me.chFlatPDF.UseVisualStyleBackColor = True
        '
        'btnSelectAll2
        '
        Me.btnSelectAll2.Location = New System.Drawing.Point(149, 19)
        Me.btnSelectAll2.Name = "btnSelectAll2"
        Me.btnSelectAll2.Size = New System.Drawing.Size(89, 40)
        Me.btnSelectAll2.TabIndex = 22
        Me.btnSelectAll2.Text = "(De)Select All 2"
        Me.btnSelectAll2.UseVisualStyleBackColor = True
        '
        'lblExplanation2
        '
        Me.lblExplanation2.AutoSize = True
        Me.lblExplanation2.ForeColor = System.Drawing.Color.Red
        Me.lblExplanation2.Location = New System.Drawing.Point(9, 320)
        Me.lblExplanation2.Name = "lblExplanation2"
        Me.lblExplanation2.Size = New System.Drawing.Size(165, 13)
        Me.lblExplanation2.TabIndex = 18
        Me.lblExplanation2.Text = "Cross box selection is not allowed"
        '
        'btnCheckExists
        '
        Me.btnCheckExists.Location = New System.Drawing.Point(16, 19)
        Me.btnCheckExists.Name = "btnCheckExists"
        Me.btnCheckExists.Size = New System.Drawing.Size(178, 37)
        Me.btnCheckExists.TabIndex = 19
        Me.btnCheckExists.Text = "Check If Drawing Exists"
        Me.btnCheckExists.UseVisualStyleBackColor = True
        '
        'grbxCreate
        '
        Me.grbxCreate.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.grbxCreate.Controls.Add(Me.btnCheckExists)
        Me.grbxCreate.Controls.Add(Me.chCreateNonFlat)
        Me.grbxCreate.Controls.Add(Me.chCreateFlat)
        Me.grbxCreate.Location = New System.Drawing.Point(12, 12)
        Me.grbxCreate.Name = "grbxCreate"
        Me.grbxCreate.Size = New System.Drawing.Size(211, 114)
        Me.grbxCreate.TabIndex = 21
        Me.grbxCreate.TabStop = False
        Me.grbxCreate.Text = "Create Drawing For Active Config"
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(488, 91)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(37, 13)
        Me.lblStatus.TabIndex = 24
        Me.lblStatus.Text = "Status"
        '
        'tbxStatus
        '
        Me.tbxStatus.Location = New System.Drawing.Point(491, 106)
        Me.tbxStatus.Name = "tbxStatus"
        Me.tbxStatus.Size = New System.Drawing.Size(158, 20)
        Me.tbxStatus.TabIndex = 25
        Me.tbxStatus.Text = "Ready for next command"
        '
        'DrawingToolbox
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(661, 353)
        Me.Controls.Add(Me.tbxStatus)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.grbxCreate)
        Me.Controls.Add(Me.lblExplanation2)
        Me.Controls.Add(Me.grbxFlat)
        Me.Controls.Add(Me.grbxNonFlat)
        Me.Controls.Add(Me.lblExplanation1)
        Me.Controls.Add(Me.lblVersionText)
        Me.Controls.Add(Me.lblProcedureList)
        Me.Controls.Add(Me.lblDisplayText)
        Me.Controls.Add(Me.btnRunMacros)
        Me.Name = "DrawingToolbox"
        Me.Text = "DrawingToolbox"
        Me.grbxNonFlat.ResumeLayout(False)
        Me.grbxNonFlat.PerformLayout()
        Me.grbxFlat.ResumeLayout(False)
        Me.grbxFlat.PerformLayout()
        Me.grbxCreate.ResumeLayout(False)
        Me.grbxCreate.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnRunMacros As Button
    Friend WithEvents chCreateNonFlat As CheckBox
    Friend WithEvents chNonFlatModelItem As CheckBox
    Friend WithEvents chNonFlatBOMTable As CheckBox
    Friend WithEvents chNonFlatAutoBalloon As CheckBox
    Friend WithEvents lblDisplayText As Label
    Friend WithEvents chNonFlatRevTable As CheckBox
    Friend WithEvents chCreateFlat As CheckBox
    Friend WithEvents chFlatPunchTable As CheckBox
    Friend WithEvents chFlatOrdDim As CheckBox
    Friend WithEvents chFlatRevTable As CheckBox
    Friend WithEvents lblProcedureList As Label
    Friend WithEvents lblVersionText As Label
    Friend WithEvents lblExplanation1 As Label
    Friend WithEvents grbxNonFlat As GroupBox
    Friend WithEvents grbxFlat As GroupBox
    Friend WithEvents btnCheckExists As Button
    Friend WithEvents grbxCreate As GroupBox
    Friend WithEvents btnSelectAll1 As Button
    Friend WithEvents btnSelectAll2 As Button
    Friend WithEvents lblExplanation2 As Label
    Friend WithEvents lblStatus As Label
    Friend WithEvents tbxStatus As TextBox
    Friend WithEvents chNonFlatPDF As CheckBox
    Friend WithEvents chFlatPDF As CheckBox
    Friend WithEvents chNonFlatReloadSheetFormat As CheckBox
    Friend WithEvents chFlatReloadSheetFormat As CheckBox
End Class
