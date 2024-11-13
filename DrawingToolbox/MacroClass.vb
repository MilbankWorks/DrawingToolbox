Imports EPDM.Interop.epdm

Public Class MacroClass
    Public Property RootPath As String
    Public Property Name As String
    Public Property [Module] As String
    Public Property Procedure As String

    Private swVault As New EdmVault5
    Private swFile As IEdmFile12
    Private swFolder As IEdmFolder10


    Public Sub New(swVault As EdmVault5, macroName As String, macroModule As String, Optional macroRootPath As String = "C:\PDM_Milbank\SolidWorks Library\Macros\Drawings\", Optional macroProcedure As String = "main")
        If Not swVault.IsLoggedIn Then swVault.LoginAuto("PDM_Milbank", 0)
        RootPath = macroRootPath
        Name = macroName
        [Module] = macroModule
        Procedure = macroProcedure

        'Get latest of macros when loading up form
        swFile = swVault.GetFileFromPath(RootPath + Name, swFolder)
        If swFile.CurrentVersion <> swFile.GetLocalVersionNo(swFolder.ID) Then
            If Not swFile.IsLocked Then
                swFile.GetFileCopy(0)
            End If
        End If

    End Sub

End Class
