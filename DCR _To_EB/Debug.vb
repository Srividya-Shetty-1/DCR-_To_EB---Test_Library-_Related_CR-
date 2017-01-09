Imports System.IO
Module Debug

    Dim debug_mode As Boolean = 0

    Public Function Check_debug()
        If File.Exists(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\enable_dcr_logging.txt") Then
            debug_mode = 1
            Message("Found enable_dcr_logging.txt. Logging Enabled")
        End If
    End Function

    Public Function Message(msg As String)
        If (debug_mode) Then
            MessageBox.Show(msg)
        End If
    End Function

End Module
