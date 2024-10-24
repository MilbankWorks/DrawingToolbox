Imports System.Text
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst

Module Util

    Public Function ConcatArray(array As List(Of String), Optional delimiter As String = (","), Optional padOption As PadOpt = PadOpt.None, Optional padChar As String = "") As String
        'delimiter = delimiter & vbCrLf
        Dim i As Integer = 0
        Dim szBuilder As StringBuilder = New StringBuilder

        'Array should be presorted, or an flag can be added for it
        'array.Sort()

        Select Case padOption
            Case PadOpt.Parantheses
                szBuilder.Append("(")
            Case PadOpt.Character
                szBuilder.Append(padChar)
        End Select

        For Each sz In array
            szBuilder.Append(sz)
            szBuilder.Append(delimiter)
        Next
        Dim endpad As Integer = Len(delimiter)

        If array.Count > 0 Then
            szBuilder.Remove(szBuilder.Length - endpad, endpad)
        End If

        Select Case padOption
            Case PadOpt.Parantheses
                szBuilder.Append(")")
            Case PadOpt.Character
                szBuilder.Append(padChar)
        End Select

        'Debug.Print(szBuilder.ToString())
        Return szBuilder.ToString()
    End Function

    Enum PadOpt
        None
        Parantheses
        Character

    End Enum

    Public Function GetDocType(Name As String) As swDocumentTypes_e
        Dim nDocType As swDocumentTypes_e = swDocumentTypes_e.swDocNONE
        If Name.EndsWith("SLDPRT") Or Name.EndsWith("sldprt") Then
            nDocType = swDocumentTypes_e.swDocPART
        ElseIf Name.EndsWith("SLDASM") Or Name.EndsWith("sldasm") Then
            nDocType = swDocumentTypes_e.swDocASSEMBLY
        ElseIf Name.EndsWith("SLDDRW") Or Name.EndsWith("slddrw") Then
            nDocType = swDocumentTypes_e.swDocDRAWING
        End If
        Return nDocType
    End Function

End Module
