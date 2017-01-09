Imports System.IO
Imports System.Xml

Public Class Syn_Main
    Dim Shared_path As String
    Public Sub main()

        Try
            Dim cls As New SYN_Input_validation
            cls.checking(Shared_drive())
            MessageBox.Show("Excel validation Successfully Completed")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        Try
            Dim SingOn As New Syn_Sign_On
            SingOn.Sign_On()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Application.Exit()
        End Try

        Try
            Dim Syn_Input As New Syn_Input_File
            Dim Local_File_Path = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\DCR_Capture\"
            Syn_Input.Read_Input_File(Local_File_Path)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Application.Exit()
        End Try
    End Sub
    Public Function Shared_drive()
        Dim xmldoc As New XmlDocument
        Dim xmlnode As XmlNodeList
        Dim exePath As String = System.Windows.Forms.Application.ExecutablePath
        Dim fileName As String = Path.GetDirectoryName(exePath)
        Dim resourcepath = fileName & "\Resources\"
        Dim config As String = resourcepath & "Path_Config.xml"
        xmldoc.Load(config)
        xmlnode = xmldoc.GetElementsByTagName("SharedLocation")
        Shared_path = xmlnode(0).Attributes.ItemOf("Path").Value
        Return Shared_path
    End Function

End Class
