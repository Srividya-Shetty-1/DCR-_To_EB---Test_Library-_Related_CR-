Imports System.IO
Imports System.Xml
Imports MetroFramework

Public Class Syn_Main
    Dim Shared_path As String
    Public Sub main()

        Debug.Check_debug()

        'Step 1
        Try
            Dim cls As New SYN_Input_validation
            cls.checking(Shared_drive())
        Catch ex As Exception
            'MetroMessageBox.Show(Form1, "Tool could not sign-on to EB system" & vbCrLf & "Please try again later", "EB Sign-On Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            MetroMessageBox.Show(Form1, ex.Message)
        End Try

        ''Step 2
        Try
            Form1.MetroLabel1.Text = Nothing
            Form1.MetroLabel1.Update()
            Form1.MetroLabel1.Text = "EB Sign On In Progress"
            Form1.MetroLabel1.Update()

            Dim SingOn As New Syn_Sign_On
            SingOn.Sign_On()
        Catch ex As Exception

            MetroMessageBox.Show(Form1, ex.Message)
            Application.Exit()
        End Try
        ''Step3
        Try
            Dim Syn_Input As New Syn_Input_File
            Dim Local_File_Path = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\DCR_Capture\"
            Syn_Input.Read_Input_File(Local_File_Path)
        Catch ex As Exception
            MetroMessageBox.Show(Form1, ex.Message)
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
