Imports System.Data.OleDb
Imports System.IO
Imports System.Xml
Imports MetroFramework

Public Class Form1
    Dim Shared_path As String
    Dim shared_ip_filepath As String
    Dim shared_ip_filename As String
    Dim new_shared_ip_filename As String
    Dim stppath As String = Shared_drive()
    Dim ip_filename As String
    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

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
    Function rename(filename As String)
        shared_ip_filepath = stppath & "\" & filename
        shared_ip_filename = Path.GetFileName(shared_ip_filepath)
        new_shared_ip_filename = shared_ip_filename.Replace("_1.xlsx", ".xlsx")
        My.Computer.FileSystem.RenameFile(stppath & "\" & shared_ip_filename, new_shared_ip_filename)
    End Function
    Function GetStringValue(ByVal value As Object) As String
        If value Is DBNull.Value Then
            GetStringValue = Nothing
        Else
            GetStringValue = value
        End If
    End Function
    Private Sub MetroButton1_Click(sender As Object, e As EventArgs) Handles MetroButton1.Click
        Try
            Dim result_conn As OleDbConnection
            Dim ds As New DataSet
            Dim record_count As Integer

            Dim Local_File_Path As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\DCR_Capture"

            Dim file As FileInfo
            Dim processed_filename As String

            Dim STATUS As String
            Dim EB_Val As String = "False"
            'Check if DCR_Capture folder exist in the local folder 

            If Directory.Exists(Local_File_Path) Then
                Dim dir As DirectoryInfo = New DirectoryInfo(Local_File_Path)
                Dim op_xlsx_files As FileInfo() = dir.GetFiles("*_1op.xlsx")
                Dim ip_xlsx_files As FileInfo() = dir.GetFiles("*_1.xlsx")

                If op_xlsx_files.Length > 0 Then
                    For Each file In op_xlsx_files
                        processed_filename = file.ToString  ' File in the local drive that has completed basic validation 
                        result_conn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Local_File_Path & "\" & processed_filename & ";extended properties=Excel 12.0 Xml;")
                        result_conn.Open()
                        Dim ws_selectcmd As New OleDbDataAdapter("select [Status] from [INPUT_CASES$] Where ([Sl No]<>'') ", result_conn)
                        ws_selectcmd.Fill(ds)
                        record_count = ds.Tables(0).Rows.Count
                        For i = 0 To record_count - 1
                            STATUS = GetStringValue(ds.Tables(0).Rows(i).ItemArray(0))
                            If STATUS = "Y" Then
                                EB_Val = "True"
                                Exit For
                            End If
                        Next
                        ds.Clear()
                        result_conn.Close()
                        If EB_Val = "False" Then

                            'Delete the output file and input file in the local drive 
                            ip_filename = processed_filename.Replace("_1op.xlsx", "_1.xlsx")
                            My.Computer.FileSystem.DeleteFile(Local_File_Path & "\" & processed_filename)
                            My.Computer.FileSystem.DeleteFile(Local_File_Path & "\" & ip_filename)

                            'Rename the input file in th shared drive
                            rename(ip_filename)
                        End If
                    Next

                ElseIf ip_xlsx_files.Length > 0 Then

                    'Delete the input file in the local drive
                    For Each file In ip_xlsx_files
                        Dim input_filename As String = file.ToString
                        My.Computer.FileSystem.DeleteFile(Local_File_Path & "\" & input_filename)

                        'Rename the input file in the shared drive
                        rename(input_filename)
                    Next
                End If

            End If

            Dim main_clsaa As New Syn_Main
            main_clsaa.main()
        Catch ex As Exception
            MetroMessageBox.Show(Me, ex.Message)
        End Try

    End Sub

    Private Sub MetroProgressBar1_Click(sender As Object, e As EventArgs) Handles MetroProgressBar1.Click

    End Sub
End Class
