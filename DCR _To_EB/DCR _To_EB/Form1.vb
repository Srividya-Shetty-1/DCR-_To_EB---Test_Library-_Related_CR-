Imports System.IO

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Copy Username and Password - DCR and EB.XLSX File to User - Document Folder
        If Not File.Exists(Directory.GetCurrentDirectory & "\Path_Config.xml") Then
            My.Computer.FileSystem.WriteAllBytes(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\Username and Password - DCR and EB.xlsx", My.Resources.Username_and_Password___DCR_and_EB, True)
        End If
        'Main Class
        Dim main_clsaa As New Syn_Main
        main_clsaa.main()

    End Sub
End Class
