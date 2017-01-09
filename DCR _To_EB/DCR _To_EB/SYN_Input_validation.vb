Imports System
Imports System.Data
Imports System.IO
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports System.Xml
Imports System.Globalization
Imports Microsoft.VisualBasic.FileIO
Imports System.Text

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
    Private Field_Names() As String = {"SLNO", "BILLNO", "POLNO", "CHKBNKCODE", "CHKBRNCODE", "CHKNO", "CHKDATE", "REF", "CURR", "RECV_AMT", "OPT", "PAID_AMT", "RECEIPT", "WRITE_OFF", "REFUND", "FUND", "BNK_ACCNT"}
    Sub checking(stppath As String)

        'MessageBox.Show(stppath)

        Dim Dir_Folder As New IO.DirectoryInfo(stppath)
        Dim Dir_Files As IO.FileInfo() = Dir_Folder.GetFiles()
        Dim Dir_File_Info As IO.FileInfo

        'Loop for each excel file in the shared folder
        For Each Dir_File_Info In Dir_Folder.GetFiles("*.xlsx")
            Dim old_file As String
            'MessageBox.Show(Dir_Folder.GetFiles.Count)
            old_file = Dir_File_Info.ToString 'Get the file name 
            ' MessageBox.Show(old_file)

            Dim record_count As Integer
            Dim Local_Folder_Path As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\DCR_Capture"
            Dim SheetName As String = "Sheet1"
            If File.Exists(stppath + "\" + old_file) Then

                Dim last_character As String = "_1.xlsx"
                'Check if the file name already contains _1
                Dim last_character_check As Boolean = old_file.Contains(last_character)

                If last_character_check = False Then
                    Dim old_filepath As String = stppath

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
                    result_conn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Local_Folder_Path & "\" & new_file & ";extended properties=Excel 12.0 Xml;")
                    result_conn.Open()
                    Dim ws_selectcmd As New OleDbDataAdapter("select * from [Sheet1$] Where ([Sl No]<>'') ", result_conn)

                    ws_selectcmd.Fill(ds)
                    record_count = ds.Tables(0).Rows.Count

                    For i = 0 To record_count - 1
                        If Validate_Fields(i) = True Then
                            passhandling(output_filename)
                        End If

                    Next i
                    result_conn.Close()

                Else
                    MessageBox.Show("The selected file Is being processed")
                End If
            End If

            ds.Clear()
        Next

    End Sub
    Private Function Validate_Fields(trans_count As Integer)
        Dim i As Integer
        Validate_Fields = True
        For i = 0 To UBound(Field_Names)
            If Not Get_Validation_Values(i, trans_count) Then
                Validate_Fields = False
                Exit Function
            End If
        Next
    End Function
    Private Function Get_Validation_Values(count As Integer, trans_count As Integer)
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
                    Get_Validation_Values = Field_Validation(count, trans_count, ValCond, ValCondValue, ValCondMessage)
                    If Get_Validation_Values = False Then
                        Exit Function
                    End If
                Next
            End If
        Next
    End Function
    Private Function Field_Validation(count As Integer, trans_count As Integer, Condition As String, Cond_Value As String, Cond_message As String)
        Dim provider As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture
        Dim regDate As Date = Date.Now()
        Field_Validation = True

        Dim input = Get_Input(count, trans_count)

        If count > 0 And input = "" Or Nothing Then
            Field_Validation = True
            Exit Function
        End If

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
                If Not Date.TryParseExact(input, "yyyyMMdd", New CultureInfo("en-US"), DateTimeStyles.None, Date1) Then
                    Field_Validation = False
                End If

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
            exceptionhandling(Cond_message, output_filename)
        End If
    End Function
    Private Function Get_Input(count As Integer, trans_count As Integer)
        Dim i As Integer
        i = trans_count
        SLNO = GetStringValue(ds.Tables(0).Rows(i).ItemArray(0))
        BILLNO = GetStringValue(ds.Tables(0).Rows(i).ItemArray(1))
        POLNO = GetStringValue(ds.Tables(0).Rows(i).ItemArray(2))
        CHKBNKCODE = GetStringValue(ds.Tables(0).Rows(i).ItemArray(3))
        CHKBRNCODE = GetStringValue(ds.Tables(0).Rows(i).ItemArray(4))
        CHKNO = GetStringValue(ds.Tables(0).Rows(i).ItemArray(5))
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

    Function exceptionhandling(description As String, path As String)
        Dim excep_conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & path & ";extended properties=Excel 12.0 Macro;")
        excep_conn.Open()
        Dim excep_update As New OleDbCommand("insert into [Exception$] ([Sl No],[Bill No],[Policy Holder],[Cheque Bank Code],[Cheque Branch Code],[Cheque No],[Currency],[Cheque Date],[Reference],[Cheque/Cash Amount],[Full/Partial],[Paid Amount],[Receipt],[Write Off Amount],[Refund Amount],[Fund],[Bank Account],[Error Description]) values('" & SLNO & "','" & BILLNO & "','" & POLNO & "','" & CHKBNKCODE & "', '" & CHKBRNCODE & "','" & CHKNO & "','" & CURR & "','" & CHKDATE & "','" & REF & "','" & RECV_AMT & "','" & OPT & "','" & PAID_AMT & "','" & RECEIPT & "','" & WRITE_OFF & "','" & REFUND & "','" & FUND & "','" & BNK_ACCNT & "','" & description & "')", excep_conn)
        excep_update.ExecuteNonQuery()
        excep_conn.Close()
    End Function
    Function passhandling(path As String)
        Dim pass_conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & path & ";extended properties=Excel 12.0 Macro;")
        pass_conn.Open()
        Dim pass_update As New OleDbCommand("insert into [To_Be_Processed$] ([Sl No],[Bill No],[Policy Holder],[Cheque Bank Code],[Cheque Branch Code],[Cheque No],[Currency],[Cheque Date],[Reference],[Cheque/Cash Amount],[Full/Partial],[Paid Amount],[Receipt],[Write Off Amount],[Refund Amount],[Fund],[Bank Account]) values('" & SLNO & "','" & BILLNO & "','" & POLNO & "','" & CHKBNKCODE & "','" & CHKBRNCODE & "','" & CHKNO & "','" & CURR & "','" & CHKDATE & "','" & REF & "','" & RECV_AMT & "','" & OPT & "','" & PAID_AMT & "','" & RECEIPT & "','" & WRITE_OFF & "','" & REFUND & "','" & FUND & "','" & BNK_ACCNT & "')", pass_conn)
        pass_update.ExecuteNonQuery()
        pass_conn.Close()
    End Function
    Private Function Load_XML_File()
        xmldoc.Load(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\DCR_Capture\Configurable_Files\DCR_Config.xml")
    End Function
End Class