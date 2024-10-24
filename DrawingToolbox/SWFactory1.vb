Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Imports System
Imports System.Diagnostics
Imports System.Threading.Tasks
Imports System.IO
Imports System.Runtime.InteropServices.ComTypes
Imports System.Linq
Imports System.ComponentModel
Imports System.Runtime.CompilerServices

Namespace Factories
    Public Module Factories
        <DllImport("ole32.dll")>
        Private Function CreateBindCtx(ByVal reserved As UInteger, <Out> ByRef ppbc As IBindCtx) As Integer
        End Function
        Public Class SWFactory


            Private swFileType As swDocumentTypes_e
            Private swApp As SldWorks
            Private nStart As Double
            Private lastSWPrcID As Integer
            Private _swPrcID As Integer = 0

            Public Property swPrcID As Integer
                Get
                    Return _swPrcID
                End Get
                Private Set(value As Integer)
                    _swPrcID = value
                End Set
            End Property


            Public Sub New()

            End Sub

            Public Function GetNextSWInstance() As ISldWorks
                lastSWPrcID = _swPrcID
                swApp = GetSwAppFromExisting(lastSWPrcID)
                If swApp IsNot Nothing Then
                    Return swApp
                Else
                    Return GetSwAppFromExisting()
                End If

            End Function

            Function StartSW() As ISldWorks

                Const apppath As String = "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\Sldworks.exe"
                Dim app As ISldWorks = Nothing
                app = GetSwAppFromExisting()
                If app Is Nothing Then
                    Dim prc = Process.Start(apppath)
                    nStart = DateAndTime.Timer

                    Do While app Is Nothing
                        If DateAndTime.Timer - nStart > 300 Then Throw New TimeoutException()
                        app = TryCast(GetSwAppFromProcess(prc.Id), SolidWorks.Interop.sldworks.ISldWorks)
                    Loop
                    Threading.Thread.Sleep(4000)
                    swPrcID = prc.Id
                    'Debug.Print(swPrcID)
                End If

                'wait for solidworks to actually start
                'Threading.Thread.Sleep(5000)

                Return app
            End Function

            Private Function GetSwAppFromProcess(ByVal processId As Integer) As ISldWorks

                Dim monikerName = "SolidWorks_PID_" & processId.ToString()

                Dim context As IBindCtx = Nothing
                Dim rot As IRunningObjectTable = Nothing
                Dim monikers As IEnumMoniker = Nothing

                Try
                    CreateBindCtx(0, context)

                    context.GetRunningObjectTable(rot)
                    rot.EnumRunning(monikers)

                    Dim moniker = New IMoniker(0) {}

                    While monikers.[Next](1, moniker, IntPtr.Zero) = 0

                        Dim curMoniker = moniker.First()
                        Dim name As String = Nothing

                        If curMoniker IsNot Nothing Then

                            Try
                                curMoniker.GetDisplayName(context, Nothing, name)
                            Catch ex As UnauthorizedAccessException
                            End Try

                        End If

                        If String.Equals(monikerName, name, StringComparison.CurrentCultureIgnoreCase) Then
                            Dim app As Object = Nothing
                            rot.GetObject(curMoniker, app)
                            Return TryCast(app, ISldWorks)
                        End If

                    End While

                Finally

                    If monikers IsNot Nothing Then
                        Marshal.ReleaseComObject(monikers)
                    End If

                    If rot IsNot Nothing Then
                        Marshal.ReleaseComObject(rot)
                    End If

                    If context IsNot Nothing Then
                        Marshal.ReleaseComObject(context)
                    End If
                End Try

                Return Nothing

            End Function

            Public Function GetSwAppFromExisting(Optional prcID As Object = Nothing) As ISldWorks

                Dim foundSW As Boolean = False
                Dim monikerName = "SolidWorks_PID_"

                Dim context As IBindCtx = Nothing
                Dim rot As IRunningObjectTable = Nothing
                Dim monikers As IEnumMoniker = Nothing

                Try

                    CreateBindCtx(0, context)

                    context.GetRunningObjectTable(rot)
                    rot.EnumRunning(monikers)

                    Dim moniker = New IMoniker(0) {}

                    While monikers.[Next](1, moniker, IntPtr.Zero) = 0

                        Dim curMoniker = moniker.First()
                        Dim name As String = Nothing

                        If curMoniker IsNot Nothing Then

                            Try
                                curMoniker.GetDisplayName(context, Nothing, name)
                            Catch ex As UnauthorizedAccessException
                            End Try

                        End If
                        'Debug.Print(name)
                        If String.Equals(Left(monikerName, 15), Left(name, 15), StringComparison.CurrentCultureIgnoreCase) Then
                            'Debug.Print(prcID)
                            'Debug.Print(Right(name, Len(name) - 15))
                            'Debug.Print(vbCrLf)
                            If prcID Is Nothing Then
                                'Debug.Print("Moniker Name: " & name)
                                'Debug.Print(swPrcID)
                                swPrcID = Right(name, Len(name) - 15)
                                Dim app As Object = Nothing
                                rot.GetObject(curMoniker, app)
                                Return TryCast(app, ISldWorks)
                            ElseIf foundSW = False Then
                                If lastSWPrcID = Right(name, Len(name) - 15) Then
                                    foundSW = True
                                End If
                            Else
                                swPrcID = Right(name, Len(name) - 15)
                                Dim app As Object = Nothing
                                rot.GetObject(curMoniker, app)
                                Return TryCast(app, ISldWorks)
                            End If



                        End If

                    End While

                Finally

                    If monikers IsNot Nothing Then
                        Marshal.ReleaseComObject(monikers)
                    End If

                    If rot IsNot Nothing Then
                        Marshal.ReleaseComObject(rot)
                    End If

                    If context IsNot Nothing Then
                        Marshal.ReleaseComObject(context)
                    End If
                End Try

                Return Nothing

            End Function

            Public Function getRunningSW() As ISldWorks
                Return GetSwAppFromProcess(swPrcID)
            End Function

            Public Sub KillSW()
                Dim Proc() = Process.GetProcesses()
                For Each item In Proc
                    If item.Id = swPrcID Then
                        Try
                            item.Kill()
                        Catch
                            System.Threading.Thread.Sleep(1000)
                        End Try
                        Exit For
                    End If
                Next
                swPrcID = 0
            End Sub


        End Class
    End Module

End Namespace
