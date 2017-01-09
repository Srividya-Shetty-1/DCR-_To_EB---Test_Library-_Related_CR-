Imports System.Data.OleDb
Imports System.IO

Public Class Syn_Dcr_Reportvb
    Sub reporting(outputfile As String, stppath As String, input_file As String)

        Dim result_conn As OleDbConnection
        Dim ds As New DataSet
        Dim dd As New DataSet
        Dim Exception_count, Pass_count As Integer
        Dim Output_Folder As String
        Dim time = DateTime.Now
        Dim st = time.ToString("yyyyMMdd")

        result_conn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & outputfile & ";extended properties=Excel 12.0 Xml;")
        result_conn.Open()

        Dim ws_selectcmd As New OleDbDataAdapter("select * from [Exception$]  Where ([Sl No]<>'')", result_conn)
        ws_selectcmd.Fill(ds)
        Exception_count = ds.Tables(0).Rows.Count

        Dim ws_selectcmd1 As New OleDbDataAdapter("select * from [Pass$] Where ([Sl No] <>'')", result_conn)

        ws_selectcmd1.Fill(dd)
        Pass_count = dd.Tables(0).Rows.Count

        Dim pass_update As New OleDbCommand("insert into [Summary$] ([Pass_Records],[Exception_Records]) values('" & Pass_count & "','" & Exception_count & "')", result_conn)
        pass_update.ExecuteNonQuery()
        result_conn.Close()

        'Dim new_outputfile As String
        Dim old_outputfile, new_outputfile As String
        old_outputfile = Path.GetFileName(outputfile)

        ''Rename the output file from _1 to _2 
        new_outputfile = old_outputfile.Replace("_1op.xlsx", "_2op.xlsx")
        My.Computer.FileSystem.RenameFile(outputfile, new_outputfile)

        'old_outputfilepath is documents/DCR_Capture 
        Dim old_outputfilepath As String = Path.GetDirectoryName(outputfile)
        Dim new_outputfilepath As String = old_outputfilepath & "\" & new_outputfile

        'Rename the input file in the local drive from _1 to _2 
        Dim new_inputfile As String = input_file.Replace("_1.xlsx", "_2.xlsx")
        My.Computer.FileSystem.RenameFile(old_outputfilepath & "\" & input_file, new_inputfile)

        'Create a date folder if it does not exist
        If Not Directory.Exists(old_outputfilepath & "\" & st) Then
            Directory.CreateDirectory(old_outputfilepath & "\" & st & "\ Input_Files")
            Directory.CreateDirectory(old_outputfilepath & "\" & st & "\ Output_Files")
        End If

        'Copy the input and the output file to the today's date folder from the local drive  
        My.Computer.FileSystem.MoveFile(old_outputfilepath & "\" & new_inputfile, old_outputfilepath & "\" & st & "\ Input_Files\" & new_inputfile)
        My.Computer.FileSystem.MoveFile(old_outputfilepath & "\" & new_outputfile, old_outputfilepath & "\" & st & "\ Output_Files\" & new_outputfile)


        'copy the processed file to the shared folder 
        Output_Folder = "Output" & st
        Dim Shared_Output_Folder_Path As String = stppath & "\" & Output_Folder
        If Not Directory.Exists(Shared_Output_Folder_Path) Then
            Directory.CreateDirectory(Shared_Output_Folder_Path)
        End If
        My.Computer.FileSystem.CopyFile(old_outputfilepath & "\" & st & "\ Output_Files\" & new_outputfile, Shared_Output_Folder_Path & "\" & new_outputfile)

    End Sub
End Class
