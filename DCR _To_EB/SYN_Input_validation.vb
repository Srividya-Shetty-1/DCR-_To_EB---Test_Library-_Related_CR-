Imports System
Imports System.Data
Imports System.IO
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports System.Xml
Imports System.Globalization
Imports Microsoft.VisualBasic.FileIO
Imports System.Text
Imports MetroFramework

Public Class SYN_Input_validation
    Dim result_conn As OleDbConnection
    Dim xmldoc As New XmlDocument
    Dim ds As New DataSet
    Public SLNO As String
    Public BILLNO As String
    Public POLNO As String
    Public CHKBNKCODE As String
    Public CHKBRNCODE As String
    Public CHKNO As String
    Public CHKDATE As String
    Public REF As String
    Public CURR As String
    Public RECV_AMT As String
    Public OPT As String
    Public PAID_AMT As String
    Public RECEIPT As String
    Public WRITE_OFF As String
    Public REFUND As String
    Public FUND As String
    Public BNK_ACCNT As String
    Dim output_filename As String
    Dim new_file As String

    Public Multi_Settlement_Case As Boolean = False
    'Array
    Dim Skip_Cheque_List As ArrayList = New ArrayList()
    Dim Skip_Cash_List As ArrayList = New ArrayList()
    Dim Multiple_Cheque_List As ArrayList = New ArrayList()
    Dim Multiple_Cash_List As ArrayList = New ArrayList()
    Public Rows_to_Delete As Stack = New Stack()

    'progress counter
    Dim counter As Integer = 1
    Private Field_Names() As String = {"SLNO", "BILLNO", "POLNO", "CHKBNKCODE", "CHKBRNCODE", "CHKNO", "CHKDATE", "REF", "CURR", "RECV_AMT", "OPT", "PAID_AMT", "RECEIPT", "WRITE_OFF", "REFUND", "FUND", "BNK_ACCNT"}
    Sub checking(stppath As String)

        Dim Dir_Folder As New IO.DirectoryInfo(stppath)
        Dim Dir_Files As IO.FileInfo() = Dir_Folder.GetFiles()
        Dim Dir_File_Info As IO.FileInfo

        'Loop for each excel file in the shared folder
        For Each Dir_File_Info In Dir_Folder.GetFiles("*.xlsx")
            Dim old_file As String
            Form1.MetroLabel1.Text = "Validation In Progress"
            Form1.MetroLabel1.Update()
            old_file = Dir_File_Info.ToString 'Get the file name 
            counter = 1
            Dim record_count As Integer
            Dim Local_Folder_Path As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\DCR_Capture"
            Dim SheetName As String = "Sheet1"
            If File.Exists(stppath + "\" + old_file) Then

                Dim last_character As String = "_1.xlsx"
                'Check if the file name already contains _1
                Dim last_character_check As Boolean = old_file.Contains(last_character)

                If last_character_check = False Then
                    Dim old_filepath As String = stppath

                    'Check if the file with the same name is already processed on that day 
                    If Not Check_File_Exists(stppath, old_file) Then


                        ' Rename the selected file by _1.xlsx
                        new_file = old_file.Replace(".xlsx", "_1.xlsx")
                        My.Computer.FileSystem.RenameFile(old_filepath & "\" & old_file, new_file)

                        Dim new_filepath As String = old_filepath & "\" & new_file

                        If Not Directory.Exists(Local_Folder_Path) Then   ' Check if the Local Folder exists 
                            Directory.CreateDirectory(Local_Folder_Path) ' Create the Local Folder
                        End If

                        'Copy the input file from shared folder to local folder
                        My.Computer.FileSystem.CopyFile(new_filepath, Local_Folder_Path & "\" & new_file)

                        'check if the config file exists in local folder 
                        If Not Directory.Exists(Local_Folder_Path & "\Configurable_Files") Then
                            My.Computer.FileSystem.CopyDirectory(old_filepath & "\Configurable_Files", Local_Folder_Path & "\Configurable_Files", True)
                        End If

                        Dim outputfile As String
                        outputfile = new_file.Replace("_1.xlsx", "_1op.xlsx")
                        output_filename = Local_Folder_Path & "\" & outputfile

                        If Not File.Exists(output_filename) Then
                            My.Computer.FileSystem.WriteAllBytes(output_filename, My.Resources.Template, False)
                        End If

                        Debug.Message("Opening input file")
                        result_conn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Local_Folder_Path & "\" & new_file & ";extended properties='Excel 12.0 Xml;IMEX=1;';")
                        result_conn.Open()
                        Dim ws_selectcmd As New OleDbDataAdapter("select * from [Sheet1$] Where Not ([Sl No] is NULL) ", result_conn)

                        ws_selectcmd.Fill(ds)
                        record_count = ds.Tables(0).Rows.Count

                        result_conn.Dispose()
                        result_conn.Close()
                        result_conn = Nothing
                        Debug.Message("Closing input file")

                        'progress bar update 
                        Form1.MetroProgressBar1.Visible = True
                        Form1.MetroProgressBar1.Minimum = 0
                        Form1.MetroProgressBar1.Maximum = record_count
                        Form1.MetroProgressBar1.Update()

                        Update_Multi_Lists()

                        For i = 0 To record_count - 1
                            Form1.MetroProgressBar1.Value = counter
                            Form1.MetroProgressBar1.Update()
                            If Validate_Fields(i) = True Then
                                passhandling(output_filename)
                            End If
                            counter = counter + 1
                        Next i


                        ds.Clear()

                    Else
                        MessageBox.Show("A File with the same name " + old_file + " has been processed today. Please change the filename and continue.")
                    End If
                Else
                    MessageBox.Show("The selected file Is being processed")

                End If
            End If
        Next
        'Form1.MetroProgressBar1.Visible = False
        Form1.MetroLabel1.Text = Nothing
        Form1.MetroLabel1.Update()
        Form1.MetroLabel1.Text = "Validation Completed"
        Form1.MetroLabel1.Update()
        Form1.MetroLabel1.Text = Nothing
        Form1.MetroLabel1.Update()
    End Sub

    Public Sub Update_Multi_Lists()

        Dim Chk_no, Ref_no As String

        'Determine if multiple cheque/cash case
        For i = 0 To ds.Tables(0).Rows.Count - 1

            Chk_no = GetStringValue(ds.Tables(0).Rows(i).ItemArray(5))
            Ref_no = GetStringValue(ds.Tables(0).Rows(i).ItemArray(7))

            ' MessageBox.Show("Multi : not last row")
            For j = i + 1 To ds.Tables(0).Rows.Count - 1
                If (Not IsNothing(Chk_no)) And Chk_no = GetStringValue(ds.Tables(0).Rows(j).ItemArray(5)) Then
                    Debug.Message("Multi Cheque Found")
                    Debug.Message("|" + Chk_no + "| |" + ds.Tables(0).Rows(j).ItemArray(5) + "|")

                    If Not Multiple_Cheque_List.Contains(Chk_no) Then
                        Multiple_Cheque_List.Add(Chk_no)
                    End If
                ElseIf (Not IsNothing(Ref_no)) And Ref_no = GetStringValue(ds.Tables(0).Rows(j).ItemArray(7)) Then
                    Debug.Message("Multi Cash Found")
                    Debug.Message("|" + Ref_no + "| |" + ds.Tables(0).Rows(j).ItemArray(7) + "|")

                    If Not Multiple_Cash_List.Contains(Ref_no) Then
                        Multiple_Cash_List.Add(Ref_no)
                    End If
                End If
            Next
        Next

    End Sub

    Public Function Is_Multi_Case(Chk As String, Ref As String)

        'Resetting the multisettlement flag
        Multi_Settlement_Case = False

        If Multiple_Cheque_List.Contains(Chk) Or Multiple_Cash_List.Contains(Ref) Then
            Debug.Message("Multi Cheque /Cash")
            Multi_Settlement_Case = True
        End If

        Return Multi_Settlement_Case
    End Function

    Public Sub Update_Skip_List(Chk As String, Ref As String)

        If Not (Chk = "" Or Chk = Nothing) Then
            'Skip Multiple cheque with Reversal
            If Not Skip_Cheque_List.Contains(Chk) Then
                Skip_Cheque_List.Add(Chk)
            End If
        End If

        If Not (Ref = "" Or Ref = Nothing) Then
            'Skip Multiple cheque with Reversal
            If Not Skip_Cash_List.Contains(Ref) Then
                Skip_Cash_List.Add(Ref)
            End If
        End If

    End Sub

    Public Function Is_Skip_Case(Chk As String, Ref As String)

        'Resetting the skip flag
        Is_Skip_Case = False

        If Skip_Cheque_List.Contains(Chk) Or Skip_Cash_List.Contains(Ref) Then
            ' MessageBox.Show("Skip Cheque /Cash")
            Is_Skip_Case = True
        End If

    End Function


    'Check if the file with the same name is already processed on that day 
    Private Function Check_File_Exists(stppath As String, old_file As String)
        Dim Input_Folder_Name As String

        Dim time = DateTime.Now
        Dim st = time.ToString("yyyyMMdd")
        Input_Folder_Name = "Input" & st
        new_file = old_file.Replace(".xlsx", "_1.xlsx")
        Dim Filepath As String = stppath + "\Backup\" + Input_Folder_Name + "\" + new_file
        If File.Exists(Filepath) Then
            Check_File_Exists = True
        Else
            Check_File_Exists = False
        End If
    End Function
    Private Function Validate_Fields(trans_count As Integer)
        Dim i As Integer
        Dim cond_message As String = "Transaction Skipped"
        Get_Record_Values(trans_count)

        If Is_Skip_Case(CHKNO, REF) Then
            cond_message = "Transaction Skipped"
            exceptionhandling(cond_message, output_filename)
            Exit Function
        End If

        If IsNothing(BILLNO) And IsNothing(POLNO) Then

            reverse_pass_case(output_filename, CHKNO, REF)

            cond_message = "There must be a value for either Bill Number or Policy Number"
            exceptionhandling(cond_message, output_filename)

            If Is_Multi_Case(CHKNO, REF) Then
                Update_Skip_List(CHKNO, REF)
            End If
        End If

        Validate_Fields = True
        For i = 0 To UBound(Field_Names)
            If Not Get_Validation_Values(i) Then
                Validate_Fields = False
                Exit Function
            End If
        Next
    End Function
    Private Function Get_Validation_Values(count As Integer)
        Dim xmlnode As XmlNodeList
        Dim xmlchildnode As XmlNodeList
        Dim ValCond As String
        Dim ValCondValue As String
        Dim ValCondMessage As String

        Get_Validation_Values = True

        Load_XML_File()
        'get the number of tags with the name as Field
        xmlnode = xmldoc.GetElementsByTagName("Field")
        'get the field name 
        For i = 0 To xmlnode.Count - 1
            If xmlnode(i).Attributes.ItemOf("Name").Value = Field_Names(count) Then
                'get the number of child nodes
                xmlchildnode = xmlnode(i).ChildNodes

                For j = 0 To xmlchildnode.Count - 1
                    ValCond = xmlchildnode(j).Attributes.ItemOf("ValCond").Value 'validation condition
                    ValCondValue = xmlchildnode(j).Attributes.ItemOf("ValCondValue").Value 'validation value
                    ValCondMessage = xmlchildnode(j).Attributes.ItemOf("ValCondMessage").Value 'validation message 
                    'Check if the field validation is satisfied 
                    Get_Validation_Values = Field_Validation(count, ValCond, ValCondValue, ValCondMessage)
                    If Get_Validation_Values = False Then
                        Exit Function
                    End If
                Next
            End If
        Next
    End Function
    Private Function Field_Validation(count As Integer, Condition As String, Cond_Value As String, Cond_message As String)
        Dim provider As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture
        Dim regDate As Date = Date.Now()
        Field_Validation = True

        Dim input = Get_Input(count)
        If count = 0 Or count = 10 Or count = 12 Then
            GoTo Validate_Fields
        ElseIf input = "" Or Nothing Then
            Field_Validation = True
            Exit Function
        End If

Validate_Fields:
        Select Case (Condition)
            Case "Blank"
                If input = "" Or Nothing Then
                    Field_Validation = False
                End If

            Case "Numeric"
                If Not (IsNumeric(input)) Then
                    Field_Validation = False
                End If

            Case "Pos_Num"
                If Not (IsNumeric(input)) Then
                    Field_Validation = False
                ElseIf (input < 0) Then
                    Field_Validation = False
                End If

            Case "Length"
                If Strings.Len(input) <> Cond_Value Then
                    Field_Validation = False
                End If

            Case "Length(Max)"
                If Strings.Len(input) > Cond_Value Then
                    Field_Validation = False
                End If

            Case "Range"
                input = Convert.ToDecimal(input)
                Dim wordArr As String() = Cond_Value.Split(":")
                Dim Min = Convert.ToInt32(wordArr(0))
                Dim Max = Convert.ToInt32(wordArr(1))
                If (input < Min) Or (input > Max) Then
                    Field_Validation = False
                End If

            Case "Decimal"
                If (Not (input = "")) And (Not (IsNumeric(input))) Then
                    Field_Validation = False
                End If
                If Field_Validation = True Then
                    'Dim result = Convert.ToSingle(input)
                    Dim wordArr As String() = input.Split(".")
                    If wordArr.Length() > 1 Then
                        If Strings.Len(wordArr(1)) > Cond_Value Then
                            Field_Validation = False
                        End If
                    End If
                End If

            Case "Date"
                Dim Date1 As DateTime
                'If Not Date.TryParseExact(input, "yyyyMMdd", New CultureInfo("en-US"), DateTimeStyles.None, Date1) Then
                If Not Date.TryParse(input, New CultureInfo("en-GB"), DateTimeStyles.None, Date1) Then
                    Field_Validation = False
                End If
              '  MessageBox.Show("Date =" + input)
            Case "Fixed Values"
                Dim strArr() As String
                Dim len As Integer
                strArr = Cond_Value.Split(" ")
                Field_Validation = False
                For len = 0 To strArr.Length - 1
                    If input = strArr(len) Then
                        Field_Validation = True
                    End If
                Next
        End Select

        If Field_Validation = False Then
            'Reverse existing pass case to exception case
            reverse_pass_case(output_filename, CHKNO, REF)
            'Move current case to exception sheet
            exceptionhandling(Cond_message, output_filename)
            'Mark this cheque no/ref no for skipping other multiple settlement entries related to this case
            If Is_Multi_Case(CHKNO, REF) Then
                Update_Skip_List(CHKNO, REF)
            End If

        End If
    End Function

    Private Function Get_Record_Values(trans_count As Integer)
        Dim i As Integer
        i = trans_count
        SLNO = GetStringValue(ds.Tables(0).Rows(i).ItemArray(0))
        BILLNO = GetStringValue(ds.Tables(0).Rows(i).ItemArray(1))
        POLNO = GetStringValue(ds.Tables(0).Rows(i).ItemArray(2))
        CHKBNKCODE = GetStringValue(ds.Tables(0).Rows(i).ItemArray(3))
        CHKBRNCODE = GetStringValue(ds.Tables(0).Rows(i).ItemArray(4))
        CHKNO = GetStringValue(ds.Tables(0).Rows(i).ItemArray(5))
        'CHKDATE = GetDateValue(ds.Tables(0).Rows(i).ItemArray(6))
        CHKDATE = GetStringValue(ds.Tables(0).Rows(i).ItemArray(6))
        REF = GetStringValue(ds.Tables(0).Rows(i).ItemArray(7))
        CURR = GetStringValue(ds.Tables(0).Rows(i).ItemArray(8))
        RECV_AMT = GetStringValue(ds.Tables(0).Rows(i).ItemArray(9))
        OPT = GetStringValue(ds.Tables(0).Rows(i).ItemArray(10))
        PAID_AMT = GetStringValue(ds.Tables(0).Rows(i).ItemArray(11))
        RECEIPT = GetStringValue(ds.Tables(0).Rows(i).ItemArray(12))
        WRITE_OFF = GetStringValue(ds.Tables(0).Rows(i).ItemArray(13))
        REFUND = GetStringValue(ds.Tables(0).Rows(i).ItemArray(14))
        FUND = GetStringValue(ds.Tables(0).Rows(i).ItemArray(15))
        BNK_ACCNT = GetStringValue(ds.Tables(0).Rows(i).ItemArray(16))
    End Function
    Private Function Get_Input(count As Integer)
        Select Case count
            Case 0
                Get_Input = SLNO
            Case 1
                Get_Input = BILLNO
            Case 2
                Get_Input = POLNO
            Case 3
                Get_Input = CHKBNKCODE
            Case 4
                Get_Input = CHKBRNCODE
            Case 5
                Get_Input = CHKNO
            Case 6
                Get_Input = CHKDATE
            Case 7
                Get_Input = REF
            Case 8
                Get_Input = CURR
            Case 9
                Get_Input = RECV_AMT
            Case 10
                Get_Input = OPT
            Case 11
                Get_Input = PAID_AMT
            Case 12
                Get_Input = RECEIPT
            Case 13
                Get_Input = WRITE_OFF
            Case 14
                Get_Input = REFUND
            Case 15
                Get_Input = FUND
            Case 16
                Get_Input = BNK_ACCNT
        End Select
    End Function
    Function GetStringValue(ByVal value As Object) As String
        If value Is DBNull.Value Then
            GetStringValue = Nothing
        Else
            GetStringValue = value
        End If
    End Function
    Function GetDateValue(ByVal value As Object) As Date
        If value Is DBNull.Value Then
            GetDateValue = Nothing
        Else
            GetDateValue = value
        End If
    End Function
    Function GetDateString(ByVal value As Object) As String
        If (value Is Nothing) Or (value = "") Then
            GetDateString = ""
        Else
            Dim datestring As Date = Convert.ToDateTime(value)
            GetDateString = datestring.ToString("yyyy/MM/dd")

        End If
    End Function

    Function reverse_pass_case(path As String, cheque As String, ref As String)
        Dim conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & path & ";extended properties=Excel 12.0 Macro;")
        conn.Open()

        Dim Sheet1 As New OleDbDataAdapter("Select * From [INPUT_CASES$]", conn)
        Dim dn As New DataSet
        Sheet1.Fill(dn)

        Dim Pass_Sl_No As String
        Dim Pass_Bill_No As String
        Dim Pass_Policy_Holder As String
        Dim Pass_Cheque_Bank_Code As String
        Dim Pass_Cheque_Branch_Code As String
        Dim Pass_Cheque_No As String
        Dim Pass_Cheque_Date As String
        Dim Pass_Reference As String
        Dim Pass_Currency As String
        Dim Pass_Cheque_Cash_Amount As String
        Dim Pass_Paid_Status As String
        Dim Pass_Paid_Amount As String
        Dim Pass_Receipt As String
        Dim Pass_Write_Off_Amount As String
        Dim Pass_Refund_Amount As String
        Dim Pass_Fund As String
        Dim Pass_Bnak_Account As String
        Dim Pass_Description As String
        Debug.Message("starting reversal loop")
        For i = 0 To dn.Tables(0).Rows.Count - 1
            'Debug.Message("|" + Cheque_No + "| |" + dn.Tables(0).Rows(i).ItemArray(5) + "|")
            If ((Not (String.IsNullOrEmpty(cheque))) And (GetStringValue(dn.Tables(0).Rows(i).ItemArray(5)) = cheque)) Or
                ((Not (String.IsNullOrEmpty(ref))) And (GetStringValue(dn.Tables(0).Rows(i).ItemArray(7)) = ref)) Then

                'Read values from excel sheet 
                Pass_Sl_No = GetStringValue(dn.Tables(0).Rows(i).ItemArray(0))
                Pass_Bill_No = GetStringValue(dn.Tables(0).Rows(i).ItemArray(1))
                Pass_Policy_Holder = GetStringValue(dn.Tables(0).Rows(i).ItemArray(2))
                Pass_Cheque_Bank_Code = GetStringValue(dn.Tables(0).Rows(i).ItemArray(3))
                Pass_Cheque_Branch_Code = GetStringValue(dn.Tables(0).Rows(i).ItemArray(4))
                Pass_Cheque_No = GetStringValue(dn.Tables(0).Rows(i).ItemArray(5))
                Pass_Cheque_Date = GetStringValue(dn.Tables(0).Rows(i).ItemArray(6))
                Pass_Reference = GetStringValue(dn.Tables(0).Rows(i).ItemArray(7))
                Pass_Currency = GetStringValue(dn.Tables(0).Rows(i).ItemArray(8))
                Pass_Cheque_Cash_Amount = GetStringValue(dn.Tables(0).Rows(i).ItemArray(9))
                Pass_Paid_Status = GetStringValue(dn.Tables(0).Rows(i).ItemArray(10))
                Pass_Paid_Amount = GetStringValue(dn.Tables(0).Rows(i).ItemArray(11))
                Pass_Receipt = GetStringValue(dn.Tables(0).Rows(i).ItemArray(12))
                Pass_Write_Off_Amount = GetStringValue(dn.Tables(0).Rows(i).ItemArray(13))
                Pass_Refund_Amount = GetStringValue(dn.Tables(0).Rows(i).ItemArray(14))
                Pass_Fund = GetStringValue(dn.Tables(0).Rows(i).ItemArray(15))
                Pass_Bnak_Account = GetStringValue(dn.Tables(0).Rows(i).ItemArray(16))

                Debug.Message("Exceptional Handling 2 : " + Pass_Sl_No)
                Pass_Description = "Transaction Reversed"

                Dim excep_update As New OleDbCommand("insert into [Exception$] ([Sl No],[Bill No],[Policy Holder],[Cheque Bank Code],[Cheque Branch Code],[Cheque No],[Currency],[Cheque Date],[Reference],[Cheque/Cash Amount],[Full/Partial],[Paid Amount],[Receipt],[Write Off Amount],[Refund Amount],[Fund],[Bank Account],[Error Description]) values('" & Pass_Sl_No & "','" & Pass_Bill_No & "','" & Pass_Policy_Holder & "','" & Pass_Cheque_Bank_Code & "', '" & Pass_Cheque_Branch_Code & "','" & Pass_Cheque_No & "','" & Pass_Currency & "','" & Pass_Cheque_Date & "','" & Pass_Reference & "','" & Pass_Cheque_Cash_Amount & "','" & Pass_Paid_Status & "','" & Pass_Paid_Amount & "','" & Pass_Receipt & "','" & Pass_Write_Off_Amount & "','" & Pass_Refund_Amount & "','" & Pass_Fund & "','" & Pass_Bnak_Account & "','" & Pass_Description & "')", conn)
                excep_update.ExecuteNonQuery()

                Rows_to_Delete.Push(i)

            End If
        Next

        conn.Close()

        Delete_incorrect_pass(Rows_to_Delete, path)

    End Function

    Sub Delete_incorrect_pass(Reverse_List As Stack, path As String)
        MessageBox.Show("Delete_incorrect_pass")
        Dim xlApp As New Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet1 As Excel.Worksheet

        '~~> Opens Workbook. Change path and filename as applicable
        xlWorkBook = xlApp.Workbooks.Open(path)

        '~~> Display Excel
        xlApp.Visible = False

        '~~> Set the source worksheet
        xlWorkSheet1 = xlWorkBook.Sheets(4)

        For Each id In Reverse_List
            '~~> Add 2 to the id to adjust for heading row and index starting at 0
            xlWorkSheet1.Rows(id + 2).Delete()
        Next

        Reverse_List.Clear()

        xlWorkBook.Save()
        xlWorkBook.Close()
        xlApp.Quit()
        xlWorkBook = Nothing
        xlApp = Nothing
        Debug.Message("Exiting Delete_incorrect_pass")
    End Sub

    Function exceptionhandling(description As String, path As String)
        Dim excep_conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & path & ";extended properties=Excel 12.0 Macro;")
        excep_conn.Open()
        Dim excep_update As New OleDbCommand("insert into [EXCEPTION$] ([Sl No],[Bill No],[Policy Holder],[Cheque Bank Code],[Cheque Branch Code],[Cheque No],[Currency],[Cheque Date],[Reference],[Cheque/Cash Amount],[Full/Partial],[Paid Amount],[Receipt],[Write Off Amount],[Refund Amount],[Fund],[Bank Account],[Error Description]) values('" & SLNO & "','" & BILLNO & "','" & POLNO & "','" & CHKBNKCODE & "', '" & CHKBRNCODE & "','" & CHKNO & "','" & CURR & "','" & CHKDATE & "','" & REF & "','" & RECV_AMT & "','" & OPT & "','" & PAID_AMT & "','" & RECEIPT & "','" & WRITE_OFF & "','" & REFUND & "','" & FUND & "','" & BNK_ACCNT & "','" & description & "')", excep_conn)
        excep_update.ExecuteNonQuery()
        excep_conn.Close()
    End Function
    Function passhandling(path As String)
        Dim pass_conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & path & ";extended properties=Excel 12.0 Macro;")
        pass_conn.Open()
        Dim pass_update As New OleDbCommand("insert into [INPUT_CASES$] ([Sl No],[Bill No],[Policy Holder],[Cheque Bank Code],[Cheque Branch Code],[Cheque No],[Currency],[Cheque Date],[Reference],[Cheque/Cash Amount],[Full/Partial],[Paid Amount],[Receipt],[Write Off Amount],[Refund Amount],[Fund],[Bank Account]) values('" & SLNO & "','" & BILLNO & "','" & POLNO & "','" & CHKBNKCODE & "','" & CHKBRNCODE & "','" & CHKNO & "','" & CURR & "','" & GetDateString(CHKDATE) & "','" & REF & "','" & RECV_AMT & "','" & OPT & "','" & PAID_AMT & "','" & RECEIPT & "','" & WRITE_OFF & "','" & REFUND & "','" & FUND & "','" & BNK_ACCNT & "')", pass_conn)
        pass_update.ExecuteNonQuery()
        pass_conn.Close()
    End Function
    Private Function Load_XML_File()
        xmldoc.Load(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\DCR_Capture\Configurable_Files\DCR_Config.xml")
    End Function
End Class