Imports SolidWorks.Interop.swconst
Imports SolidWorks.Interop.sldworks
Imports EPDM.Interop.epdm
Imports System.IO
Imports System.Reflection
Imports System.Linq.Expressions
Imports System.Runtime.InteropServices

Public Class Delegator


    Public swFactory As Factories.SWFactory


    Public Sub New(swApp)
        Me.New()
        Me.swApp = swApp
    End Sub

    Public Sub New()
        'AddHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf MyResolveEventHandler
        Try
            swVault = New EdmVault5()
            If Not swVault.IsLoggedIn Then
                swVault.LoginAuto("PDM_Milbank", 0)
            End If
            swFactory = New Factories.SWFactory
            GetSWApp()
            Dim userMgr As IEdmUserMgr10 = swVault
            swUserName = userMgr.GetLoggedInUser.Name
        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try

    End Sub

    Public Sub GetSWApp()
        swApp = swFactory.GetSwAppFromExisting()
    End Sub

    Public Sub NextSWInstance()
        swApp = swFactory.GetNextSWInstance()
    End Sub

End Class
