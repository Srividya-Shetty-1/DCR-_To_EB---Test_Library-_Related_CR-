Imports System.Data.OleDb
Imports System.IO
Imports Microsoft.Office.Interop

Public Class Syn_Dcr_Reportvb
    Sub reporting(outputfile As String, stppath As String, input_file As String)

        Dim result_conn As OleDbConnection
        Dim ds As New DataSet
        Dim dd As New DataSet
        Dim Exception_count, Pass_count As Integer
        Dim Output_Folder As String
        Dim Input_Folder As String
        Dim time = DateTime.Now
        Dim st = time.ToString("yyyyMMdd")
        Debug.Message("Reporting")
        result_conn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & outputfile & ";extended properties=Excel 12.0 Xml;")
        result_conn.Open()

        Dim ws_selectcmd As New OleDbDataAdapter("select * from [Exception$]  Where ([Sl No]<>'')", result_conn)
        ws_selectcmd.Fill(ds)
        Exception_count = ds.Tables(0).Rows.Count

        Dim ws_selectcmd1 As New OleDbDataAdapter("select * from [Pass$] Where ([Sl No] <>'')", result_conn)

        ws_selectcmd1.Fill(dd)
        Pass_count = dd.Tables(0).Rows.Count

        result_conn.Close()
        Try
            'update summary
            Call Updating_summary(outputfile, Pass_count, Exception_count)
        Catch ex As Exception

        End Try
        Try
            Exception_Border(outputfile, Exception_count)

        Catch ex As Exception
        End Try
        Try
            Pass_Border(outputfile, Pass_count)

        Catch ex As Exception
        End Try
        Try
            'Delete excelsheel
            Call Delete_sheet(outputfile)
        Catch ex As Exception
        End Try


        'Dim new_outputfile As String
        Dim old_outputfile, new_outputfile As String
        old_outputfile = Path.GetFileName(outputfile)

        ''Rename the output file from _1 to _2 
        ' MessageBox.Show("Renaming _1op to _2op")
        new_outputfile = old_outputfile.Replace("_1op.xlsx", "_2op.xlsx")
        My.Computer.FileSystem.RenameFile(outputfile, new_outputfile)

        'old_outputfilepath is documents/DCR_Capture 
        Dim old_outputfilepath As String = Path.GetDirectoryName(outputfile)
        Dim new_outputfilepath As String = old_outputfilepath & "\" & new_outputfile

        'Rename the input file in the local drive from _1 to _2 
        Debug.Message("Renaming _1 to _2")
        Dim new_inputfile As String = input_file.Replace("_1.xlsx", "_2.xlsx")

        Try
            My.Computer.FileSystem.RenameFile(old_outputfilepath & "\" & input_file, new_inputfile)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


        'Create a date folder if it does not exist
        Debug.Message("Checking and creating Input/Output folders")
        If Not Directory.Exists(old_outputfilepath & "\" & st) Then
            Directory.CreateDirectory(old_outputfilepath & "\" & st & "\ Input_Files")
            Directory.CreateDirectory(old_outputfilepath & "\" & st & "\ Output_Files")
        End If

        'Copy the input and the output file to the today's date folder from the local drive  
        Debug.Message("Moving Input/Output files to folder")
        My.Computer.FileSystem.MoveFile(old_outputfilepath & "\" & new_inputfile, old_outputfilepath & "\" & st & "\ Input_Files\" & new_inputfile)
        My.Computer.FileSystem.MoveFile(old_outputfilepath & "\" & new_outputfile, old_outputfilepath & "\" & st & "\ Output_Files\" & new_outputfile)


        Output_Folder = "Output" & st
        Input_Folder = "Input" & st
        Dim Backup_Folder_Path As String = stppath & "\" & "Backup"
        Dim Shared_Output_Folder_Path As String = Backup_Folder_Path & "\" & Output_Folder
        Dim Shared_Input_Folder_Path As String = Backup_Folder_Path & "\" & Input_Folder
        'Create a Backup folder
        If Not Directory.Exists(Backup_Folder_Path) Then
            Directory.CreateDirectory(Backup_Folder_Path)
        End If

        'copy the output processed file to the shared folder 
        If Not Directory.Exists(Shared_Output_Folder_Path) Then
            Directory.CreateDirectory(Shared_Output_Folder_Path)
        End If
        My.Computer.FileSystem.CopyFile(old_outputfilepath & "\" & st & "\ Output_Files\" & new_outputfile, Shared_Output_Folder_Path & "\" & new_outputfile)

	If Not Directory.Exists(Shared_Input_Folder_Path) Then
            Directory.CreateDirectory(Shared_Input_Folder_Path)
    End If
        My.Computer.FileSystem.MoveFile(stppath & "\" & input_file, Shared_Input_Folder_Path & "\" & input_file)
    End Sub

    Private Sub Updating_summary(excelpath As String, pass As String, exceptioncase As String)
        Dim xlapp As New Excel.Application
        Dim xlwb As Excel.Workbook = xlapp.Workbooks.Open(excelpath) 'Excel workbook to store the PDF table
        Dim xlsheet As Excel.Worksheet

        xlapp.Visible = False
        xlsheet = Nothing
        xlsheet = xlwb.Worksheets("Summary")
        xlsheet.Range("D9").Value = pass
        xlsheet.Range("D10").Value = exceptioncase
        Dim value1 As Integer = Integer.Parse(pass)
        Dim value2 As Integer = Integer.Parse(exceptioncase)
        xlsheet.Range("D7").Value = value1 + value2
        xlwb.Save()
        xlwb.Close()
        xlapp.Quit()
        xlwb = Nothing
        xlapp = Nothing

    End Sub
    Private Sub Exception_Border(excelpath As String, Exception_count As Integer)
        Dim Last_Exception_range As String
        Dim r1 As Integer
        r1 = Exception_count + 1
        Last_Exception_range = "R" & r1

        Dim xlapp As New Excel.Application
        Dim xlWorkBook As Excel.Workbook = xlapp.Workbooks.Open(excelpath)
        Dim xlWorkSheet As Excel.Worksheet
        Dim chartRange As Excel.Range
        xlWorkSheet = xlWorkBook.Sheets("EXCEPTION")
        chartRange = xlWorkSheet.Range("A2", Last_Exception_range)

        xlapp.Visible = False
        xlapp.DisplayAlerts = False

        chartRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous

        xlWorkBook.Save()
        xlWorkBook.Close()
        xlapp.Quit()
        xlWorkBook = Nothing
        xlapp = Nothing
    End Sub
    Private Sub Pass_Border(excelpath As String, Pass_count As Integer)
        Dim Last_Exception_range As String
        Dim r1 As Integer
        r1 = Pass_count + 1
        Last_Exception_range = "R" & r1

        Dim xlapp As New Excel.Application
        Dim xlWorkBook As Excel.Workbook = xlapp.Workbooks.Open(excelpath)
        Dim xlWorkSheet As Excel.Worksheet
        Dim chartRange As Excel.Range
        xlWorkSheet = xlWorkBook.Sheets("PASS")
        chartRange = xlWorkSheet.Range("A2", Last_Exception_range)

        xlapp.Visible = False
        xlapp.DisplayAlerts = False

        chartRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous

        xlWorkBook.Save()
        xlWorkBook.Close()
        xlapp.Quit()
        xlWorkBook = Nothing
        xlapp = Nothing
    End Sub
    Private Sub Delete_sheet(excelpath As String)
        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        xlApp.DisplayAlerts = False
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(excelpath, 0, False, 5, "", "", False, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", True, False, 0, True, False, False)
        Dim worksheets As Excel.Sheets = xlWorkBook.Worksheets
        worksheets(4).Delete()
        xlWorkBook.Save()
        xlWorkBook.Close()
        xlApp.Quit()
        xlWorkBook = Nothing
        xlApp = Nothing
    End Sub
End Class
