﻿Imports System.Data.OleDb
Imports System.Text.RegularExpressions
Imports System.Threading
Imports IBMISeries
Imports Microsoft.Office.Interop

Public Class Syn_Input_File
    'Autoit
    Public Window_Title As String
    Dim EB_Status As String
    Dim Proc_Scrn As String
    'Variables for Excel Input File
    Public Sl_No As String
    Public Bill_No As String
    Public Policy_Holder As String
    Public Cheque_Bank_Code As String
    Public Cheque_Branch_Code As String
    Public Cheque_No As String
    Public Cheque_Date As String
    Public Currency As String
    Public Cheque_Cash_Amount As String
    Public Paid_Status As String
    Public Paid_Amount As String
    Public Receipt As String
    Public Write_Off_Amount As String
    Public Refund_Amount As String
    Public Fund As String
    Public Bnak_Account As String
    Public Reference As String
    Dim os_amount As String
    'Array
    Dim arralist As ArrayList = New ArrayList()
    Dim Skip_Cheque As ArrayList = New ArrayList()
    Dim Multi_Cheque As ArrayList = New ArrayList()
    Dim Multi_Cash As ArrayList = New ArrayList()
    Public Rows_to_Delete As Stack = New Stack()

    'Error handling
    Public First_Error As Boolean = False
    Public Second_Error As Boolean = False
    Public Third_Error As Boolean = False
    Public Reciept_Error As Boolean = False
    Public Reversal_Error As Boolean = False
    Public Firstentry As Boolean = False
    Public Fund_check_empty_cash As String
    Public Fund_check_empty_cheque As String
    Public Second_cheque_entry As Boolean
    Public Multiple_Settlement_Case As Boolean = False
    'Business check
    Public Business_hang_seng As Boolean = False
    Public Business_Claim_Refund As Boolean = False
    'PublicV_Currency
    Public V_Currency As String
    'Excel file name and path
    Public pathfile As String
    Public Status As String
    Public excelfilepath As String
    Dim counter As Integer

    'OLEDB CONNECTION
    Public Connection As OleDbConnection

    'Input file's path from shared drive
    Public Sub Read_Input_File(Input_file_path As String)
        'Read Directoryname and information of files
        Dim Dir_Folder As New IO.DirectoryInfo(Input_file_path)
        Dim Dir_Files As IO.FileInfo() = Dir_Folder.GetFiles()
        Dim Dir_File_Info As IO.FileInfo
        Form1.MetroLabel1.Text = Nothing
        Form1.MetroLabel1.Update()
        Form1.MetroLabel1.Text = "Billing In Progress"
        Form1.MetroLabel1.Update()
        'read only .xlsx files 
        For Each Dir_File_Info In Dir_Folder.GetFiles("*.xlsx")
            pathfile = Dir_File_Info.ToString
            If Strings.Right(pathfile, 7) = "op.xlsx" Then
                excelfilepath = Input_file_path & pathfile
                'OLEDB Connection OPEN / CLOSE
                Connection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & excelfilepath & ";Extended Properties=Excel 12.0 xml;")
                Connection.Open()
                Dim Adap As New OleDbDataAdapter("Select [Sl No],[Bill No],[Policy Holder],[Cheque Bank Code],[Cheque Branch Code],[Cheque No],[Cheque Date],[Reference],[Currency],[Cheque/Cash Amount],[Full/Partial],[Paid Amount],[Receipt],[Write Off Amount],[Refund Amount],[Fund],[Bank Account],[Status] From [INPUT_CASES$]", Connection)
                Dim dt As New DataSet
                Adap.Fill(dt)

                counter = 1
                Form1.MetroProgressBar1.Minimum = 0
                Form1.MetroProgressBar1.Maximum = dt.Tables(0).Rows.Count
                Form1.MetroProgressBar1.Update()

                For i = 0 To dt.Tables(0).Rows.Count - 1

                    Form1.MetroProgressBar1.Value = counter
                    Form1.MetroProgressBar1.Update()

                    'Read values from excel sheet 
                    Sl_No = GetStringValue(dt.Tables(0).Rows(i).ItemArray(0))
                    Bill_No = GetStringValue(dt.Tables(0).Rows(i).ItemArray(1))
                    Policy_Holder = GetStringValue(dt.Tables(0).Rows(i).ItemArray(2))
                    Cheque_Bank_Code = GetStringValue(dt.Tables(0).Rows(i).ItemArray(3))
                    Cheque_Branch_Code = GetStringValue(dt.Tables(0).Rows(i).ItemArray(4))
                    Cheque_No = GetStringValue(dt.Tables(0).Rows(i).ItemArray(5))
                    Cheque_Date = GetStringValue(dt.Tables(0).Rows(i).ItemArray(6))
                    Reference = GetStringValue(dt.Tables(0).Rows(i).ItemArray(7))
                    Currency = GetStringValue(dt.Tables(0).Rows(i).ItemArray(8))
                    Cheque_Cash_Amount = GetStringValue(dt.Tables(0).Rows(i).ItemArray(9))
                    Paid_Status = GetStringValue(dt.Tables(0).Rows(i).ItemArray(10))
                    Paid_Amount = GetStringValue(dt.Tables(0).Rows(i).ItemArray(11))
                    Receipt = GetStringValue(dt.Tables(0).Rows(i).ItemArray(12))
                    Write_Off_Amount = GetStringValue(dt.Tables(0).Rows(i).ItemArray(13))
                    Refund_Amount = GetStringValue(dt.Tables(0).Rows(i).ItemArray(14))
                    Fund = GetStringValue(dt.Tables(0).Rows(i).ItemArray(15))
                    Bnak_Account = GetStringValue(dt.Tables(0).Rows(i).ItemArray(16))
                    Status = GetStringValue(dt.Tables(0).Rows(i).ItemArray(17))
                    Debug.Message("Data bill number and sil number and Amount " + Bill_No + " " + Sl_No + " " + Cheque_Cash_Amount)
                    'Resetting the multisettlement flag
                    Multiple_Settlement_Case = False
                    'Determine if multiple cheque/cash case
                    If i < dt.Tables(0).Rows.Count - 1 Then
                        Debug.Message("Multi : not last row")
                        For j = i + 1 To dt.Tables(0).Rows.Count - 1
                            If (Not IsNothing(Cheque_No)) And Cheque_No = GetStringValue(dt.Tables(0).Rows(j).ItemArray(5)) Then
                                Debug.Message("Multi Cheque Found")
                                Debug.Message("|" + Cheque_No + "| |" + dt.Tables(0).Rows(j).ItemArray(5) + "|")

                                If Not Multi_Cheque.Contains(Cheque_No) Then
                                    Multi_Cheque.Add(Cheque_No)
                                End If
                            ElseIf (Not IsNothing(Reference)) And Reference = GetStringValue(dt.Tables(0).Rows(j).ItemArray(7)) Then
                                Debug.Message("Multi Cash Found")
                                Debug.Message("|" + Reference + "| |" + dt.Tables(0).Rows(j).ItemArray(7) + "|")

                                If Not Multi_Cash.Contains(Reference) Then
                                    Multi_Cash.Add(Reference)
                                End If
                            End If

                            If Multi_Cheque.Contains(Cheque_No) Or Multi_Cash.Contains(Reference) Then
                                Debug.Message("Multi Cheque /Cash")
                                Multiple_Settlement_Case = True
                                Exit For
                            End If
                        Next
                    End If

                    If Not Status = "Y" Or Status = "SKIP" Then
                        'Skip Multiple cheque with Reversal
                        If Skip_Cheque.Contains(Cheque_No) Then
                            Debug.Message("Exception case")
                            Dim excep_Skip As New OleDbCommand("Update [INPUT_CASES$] Set [Status]='" & "SKIP" & "' Where [Sl No]='" & Sl_No & "'", Connection)
                                excep_Skip.ExecuteNonQuery()
                            Dim Description As String = "Transaction Skipped"
                            Call exceptionhandling2(Sl_No, Bill_No, Policy_Holder, Cheque_Bank_Code, Cheque_Branch_Code, Cheque_No, Cheque_Date, Reference, Currency, Cheque_Cash_Amount, Paid_Status, Paid_Amount, Receipt, Write_Off_Amount, Refund_Amount, Fund, Bnak_Account, Description)
                        Else
                            Debug.Message("get bill number and sil number " + Bill_No + Sl_No)
                            'Update values to EB system from input file
                            Debug.Message("Paid_Amount amount in main Function" + Paid_Amount)
                            Call EB_Process(Sl_No, Bill_No, Policy_Holder, Cheque_Bank_Code, Cheque_Branch_Code, Cheque_No, Cheque_Date, Reference, Currency, Cheque_Cash_Amount, Paid_Status, Paid_Amount, Receipt, Write_Off_Amount, Refund_Amount, Fund, Bnak_Account)
                        End If
                    End If
                    counter = counter + 1
                Next

                'Close OLEDB Connection
                Connection.Dispose()
                Connection.Close()
                Connection = Nothing

                Delete_incorrect_pass(Rows_to_Delete)

                'Clear Array
                arralist.Clear()
                Skip_Cheque.Clear()

                'Updating Summary report sheet
                Dim cls As New Syn_Dcr_Reportvb
                Dim S_path As New Syn_Main
                cls.reporting(excelfilepath, S_path.Shared_drive(), Strings.Left(pathfile, pathfile.Length - 7) & ".xlsx")
            End If
        Next

        Form1.MetroLabel1.Text = String.Empty
        Form1.MetroLabel1.Update()
        Form1.MetroLabel1.Text = "EB Sign Out In Progress"
        Form1.MetroLabel1.Update()
        Thread.Sleep(5000)

        Call sign_out_EB()
        Form1.MetroLabel1.Text = String.Empty
        Form1.MetroLabel1.Update()
        Form1.MetroLabel1.Text = "Billing Progress Completed"
        Form1.MetroLabel1.Update()
        'Application exit
        Application.Exit()

    End Sub


    '***********************************************CHEQUE or CASH************************************************
    '************** 1.Normar DCR Capture (1 cheque , 1 bill )***************************************************** 
    '******************** 2.1 Multiple Settlement (1 cheque , multiple bills, first bill entry )******************
    '******************** 2.2 Multiple Settlement (1 cheque , multiple bills, second bill or rest entry )*********
    '************** 3. DCR Capture with write off **************************************************************** 
    '************** 4.DCR Capture for Hang Seng Policy *********************************************************** 
    '************** 5. DCR Capture with refund ******************************************************************* 
    '************** 6. DCR Capture with non-settled chaque / cash ************************************************ 
    '************** 7. Claim refund ****************************************************************************** 
    '************** 8. Reversal ********************************************************************************** 
    Function EB_Process(EB_Sl_No As String, EB_Bill_No As String, EB_Policy_Holder As String, EB_Cheque_Bank_Code As String, EB_Cheque_Branch_Code As String, EB_Cheque_No As String, EB_Cheque_Date As String, EB_Reference As String, EB_Currency As String, EB_Cheque_Cash_Amount As String, EB_Paid_Status As String, EB_Paid_Amount As String, EB_Receipt As String, EB_Write_Off_Amount As String, EB_Refund_Amount As String, EB_Fund As String, EB_Bnak_Account As String)
        Debug.Message("In EB_Process")
        Debug.Message("Paid Amount is" + EB_Paid_Amount)
        'Error handing
        First_Error = False
        Second_Error = False
        Third_Error = False

        'Connect session() IBM Emulator
        Wrapper.HLL_ConnectPS("A")
        Wrapper.HLL_Wait()
        'Selection Option 1 for DEC MEnu
        Wrapper.HLL_CopyStringToPS("1", Wrapper.getPos(14, 43))
        'Press Enter to novigate DCR Capture Scrren
        Wrapper.SendStr("@E")
        Wrapper.HLL_Wait()

        If Not arralist.Contains(EB_Cheque_No) And (Not EB_Cheque_No = Nothing Or Not EB_Cheque_No = "") Then
            Firstentry = True
            arralist.Add(EB_Cheque_No)
        ElseIf Not arralist.Contains(EB_Reference) And (Not EB_Reference = Nothing Or Not EB_Reference = "") Then
            Firstentry = True
            arralist.Add(EB_Reference)
            Debug.Message("First Entry is true ")
        ElseIf (EB_Reference = Nothing Or EB_Reference = "") And EB_Cheque_Bank_Code = Nothing And EB_Cheque_Branch_Code = Nothing And EB_Cheque_No = Nothing And EB_Cheque_Date = Nothing And (Not EB_Currency = Nothing Or Not EB_Currency = "") And (Not EB_Cheque_Cash_Amount = Nothing Or Not EB_Cheque_Cash_Amount = "") Then
            Firstentry = True
        ElseIf EB_Cheque_No = Nothing And EB_Reference = Nothing And EB_Write_Off_Amount = Nothing And EB_Refund_Amount = Nothing Then
            Firstentry = True
            'Do nothing
        Else
            'For 2nd checque in multiple payment no need to enter the cheque details
            Debug.Message("Clearing Cheque details : Amount =" + EB_Cheque_Cash_Amount)
            Firstentry = False
            EB_Cheque_Cash_Amount = ""
            EB_Cheque_Date = Nothing
            EB_Cheque_Bank_Code = Nothing
            EB_Cheque_Branch_Code = Nothing
            EB_Cheque_No = Nothing

        End If

        'Check Condition for HANG SENG POLICIES

        Debug.Message("Cheque Details |" + EB_Cheque_Bank_Code + "||" + EB_Cheque_Branch_Code + "||" + EB_Cheque_No + "||" + EB_Cheque_Date + "||" + EB_Currency + "||" + EB_Cheque_Cash_Amount + "||" + EB_Write_Off_Amount + "||" + EB_Refund_Amount + "||")

        If EB_Cheque_Bank_Code = Nothing And EB_Cheque_Branch_Code = Nothing And EB_Cheque_No = Nothing And EB_Cheque_Date = Nothing And EB_Currency = Nothing And EB_Cheque_Cash_Amount = Nothing And EB_Write_Off_Amount = Nothing And EB_Refund_Amount = Nothing And EB_Bnak_Account IsNot Nothing And EB_Paid_Amount IsNot Nothing Then
            Debug.Message("Hang Seng") '

            Business_hang_seng = True
            '**************6. FUND ******************************************************************************************
            Wrapper.HLL_CopyStringToPS(EB_Fund, Wrapper.getPos(7, 59))

            '**************1. BANK ACCOUNT CODE *****************************************************************************
            Wrapper.HLL_CopyStringToPS(EB_Bnak_Account, Wrapper.getPos(7, 61))
            '***CASH
        ElseIf EB_Cheque_Bank_Code = Nothing And EB_Cheque_Branch_Code = Nothing And EB_Cheque_No = Nothing And EB_Cheque_Date = Nothing And EB_Cheque_Cash_Amount IsNot Nothing And EB_Bnak_Account IsNot Nothing And Currency IsNot Nothing And EB_Paid_Amount IsNot Nothing Then
            Debug.Message("Cash") '
            Debug.Message("Enter the cash amount")
            Debug.Message(EB_Cheque_Cash_Amount)
            '**************1. CHECK AMOUNT **********************************************************************************
            Wrapper.HLL_CopyStringToPS(EB_Cheque_Cash_Amount, Wrapper.getPos(8, 10))

            '**************6. FUND ******************************************************************************************
            Wrapper.HLL_CopyStringToPS(EB_Fund, Wrapper.getPos(8, 59))

            '**************1. BANK ACCOUNT CODE *****************************************************************************
            Wrapper.HLL_CopyStringToPS(EB_Bnak_Account, Wrapper.getPos(8, 61))

            '**************1. CURRENCY **************************************************************************************
            Wrapper.HLL_CopyStringToPS(Valid_Combination_Currency(Fund, Bnak_Account), Wrapper.getPos(8, 68))

        ElseIf Not String.IsNullOrEmpty(EB_Cheque_Bank_Code) And Not String.IsNullOrEmpty(EB_Cheque_Branch_Code) And Not String.IsNullOrEmpty(EB_Cheque_No) And Not String.IsNullOrEmpty(EB_Cheque_Date) And Not String.IsNullOrEmpty(EB_Cheque_Cash_Amount) And Not String.IsNullOrEmpty(EB_Bnak_Account) And Currency IsNot Nothing And EB_Paid_Amount IsNot Nothing Then

            Debug.Message("Cheque")

            '**************1. CHECK AMOUNT **********************************************************************************
            Wrapper.HLL_CopyStringToPS(EB_Cheque_Cash_Amount, Wrapper.getPos(7, 10))

            '**************2. CHECK DATE (YYYYMMDD) *************************************************************************
            Dim Trim_EB_Cheque_Date As String = Replace(EB_Cheque_Date, "/", "")
            Wrapper.HLL_CopyStringToPS(Trim_EB_Cheque_Date, Wrapper.getPos(7, 25))

            '**************3. BANK CODE *************************************************************************************
            Wrapper.HLL_CopyStringToPS(EB_Cheque_Bank_Code, Wrapper.getPos(7, 36))

            '**************4. BRANCH CODE ***********************************************************************************
            Wrapper.HLL_CopyStringToPS(EB_Cheque_Branch_Code, Wrapper.getPos(7, 40))

            '**************5. CHECK NUMBER **********************************************************************************
            Wrapper.HLL_CopyStringToPS(EB_Cheque_No, Wrapper.getPos(7, 44))

            '**************6. FUND ******************************************************************************************
            Wrapper.HLL_CopyStringToPS(EB_Fund, Wrapper.getPos(7, 59))

            '**************1. BANK ACCOUNT CODE *****************************************************************************
            Wrapper.HLL_CopyStringToPS(EB_Bnak_Account, Wrapper.getPos(7, 61))

            '**************1. CURRENCY **************************************************************************************
            Wrapper.HLL_CopyStringToPS(Valid_Combination_Currency(EB_Fund, EB_Bnak_Account), Wrapper.getPos(7, 68))
        Else
            Debug.Message("Incorrect Data")
            Dim description As String = "Incorrect Data"
            exceptionhandling(description)
            Debug.Message("F1")
            Wrapper.SendStr("@1")
            Wrapper.HLL_Wait()
            Exit Function
        End If

        'Press Enter for navigation from AC10003 Screen to AC11001
        Debug.Message("first enter to EB")

        Wrapper.SendStr("@E")
        Wrapper.HLL_Wait()

        'Check for ERROR handing / Exception Fail case
        Call Check_AC10003()
        If First_Error = True Then
            'Moving to Fail
            Debug.Message("exceptionhandling First Error bill number and sil number " + Bill_No + Sl_No)
            'Call exceptionhandling(EB_Sl_No, EB_Bill_No, EB_Policy_Holder, EB_Cheque_Bank_Code, EB_Cheque_Branch_Code, EB_Cheque_No, EB_Cheque_Date, Reference, EB_Currency, EB_Cheque_Cash_Amount, EB_Paid_Status, EB_Paid_Amount, EB_Receipt, EB_Write_Off_Amount, EB_Refund_Amount, EB_Fund, EB_Bnak_Account, EB_Status)
            Call exceptionhandling(EB_Status)
            If Not String.IsNullOrEmpty(Cheque_No) Then
                Skip_Cheque.Add(Cheque_No)
            End If
            Exit Function
        End If
        Debug.Message("Enter Bill Number")
        'Check cash/cheque with billing or policy
        If EB_Bill_No = Nothing And Not String.IsNullOrEmpty(EB_Policy_Holder) Then
            'Entering to Policy Section
            Debug.Message("Policy")  '
            Wrapper.HLL_CopyStringToPS(EB_Policy_Holder, Wrapper.getPos(11, 17))
            Business_Claim_Refund = True
        Else
            'Entering to Billing Section
            Debug.Message("Billing")  '
            Wrapper.HLL_CopyStringToPS(EB_Bill_No, Wrapper.getPos(9, 17))
        End If


        'Press Enter for navigation from AC11001 Screen DCR CAPTURE to Sundry
        Debug.Message("Second enter EB")
        Wrapper.SendStr("@E")
        Wrapper.HLL_Wait()

        'Check for ERROR handing / Exception Fail case
        Call Check_AC11001()

        If Second_Error = True Then
            If Multiple_Settlement_Case Then
                Mul_Settlement_reverse()
            End If
            'Moving to Fail
            Debug.Message("Bill exceptionhandling Second Error bill number and sil number " + Bill_No + Sl_No)
            'Call exceptionhandling(EB_Sl_No, EB_Bill_No, EB_Policy_Holder, EB_Cheque_Bank_Code, EB_Cheque_Branch_Code, EB_Cheque_No, EB_Cheque_Date, Reference, EB_Currency, EB_Cheque_Cash_Amount, EB_Paid_Status, EB_Paid_Amount, EB_Receipt, EB_Write_Off_Amount, EB_Refund_Amount, EB_Fund, EB_Bnak_Account, EB_Status)
            Call exceptionhandling(EB_Status)
            'cheque reversal
            Debug.Message("REVERSAL check condition") '
            If Not String.IsNullOrEmpty(Cheque_No) And Firstentry = False Then
                Debug.Message("REVERSAL inside") '
                Call Reversal_Checque()
            End If


            Business_hang_seng = False
            Business_Claim_Refund = False
            Exit Function
        End If

        'Sundry logic
        Debug.Message("Enter Sundry Details")

        'Get OS Amount from EB System
        Wrapper.HLL_ReadScreen(Wrapper.getPos(14, 39), 15, os_amount)
        os_amount = Trim(Strings.Left(os_amount, 15))

        If (Business_hang_seng = True) Then
            Debug.Message("Hang Seng true ")  '
            'Sundry H004 is Constant value for Hang Seng
            '**************1. PAID STATUS *****************************************************************************************
            Wrapper.HLL_CopyStringToPS(Update_OPT(EB_Paid_Status), Wrapper.getPos(14, 57))

            '**************1. PAID AMOUNT *****************************************************************************************
            Wrapper.HLL_CopyStringToPS(EB_Paid_Amount, Wrapper.getPos(14, 61))

            Call Update_Sundry("H004", Valid_Combination_Currency(EB_Fund, EB_Bnak_Account), EB_Paid_Amount & "-")

            '**************False *************************************************************************************
            Business_hang_seng = False

        ElseIf (Business_Claim_Refund = True) Then
            Debug.Message("Claim Refund") '
            'Sundry H013 is Constant value for Claim Refund
            Call Update_Sundry("H013", Valid_Combination_Currency(EB_Fund, EB_Bnak_Account), EB_Cheque_Cash_Amount)


        ElseIf Firstentry = True Then
            'Normal DCR Capture (1 cheque, 1 bill)
            Debug.Message("NDCR") '

            '**************1. PAID STATUS *****************************************************************************************
            Wrapper.HLL_CopyStringToPS(Update_OPT(EB_Paid_Status), Wrapper.getPos(14, 57))

            '**************1. PAID AMOUNT *****************************************************************************************
            Wrapper.HLL_CopyStringToPS(EB_Paid_Amount, Wrapper.getPos(14, 61))

            If Not ((EB_Cheque_Cash_Amount - EB_Paid_Amount) = "0") Then
                'Multiple_Settlement 
                If (Multiple_Settlement_Case) Then
                    '   Multiple_Settlement with Write Off 
                    Debug.Message("Multiple Settlement")
                    If Not String.IsNullOrEmpty(EB_Write_Off_Amount) Then
                        Debug.Message("Multipple settlement with write off")
                        Call Update_Sundry("A999", Valid_Combination_Currency(EB_Fund, EB_Bnak_Account), EB_Cheque_Cash_Amount - (EB_Paid_Amount + EB_Write_Off_Amount))
                        Call Update_Sundry_Two("H002", Valid_Combination_Currency(EB_Fund, EB_Bnak_Account), EB_Write_Off_Amount)
                        '   Multiple_Settlement with refund 
                    ElseIf Not String.IsNullOrEmpty(EB_Refund_Amount) Then
                        Debug.Message("Multipple settlement with refund")
                        Debug.Message("Cheque and Cash amount" + Cheque_Cash_Amount + "Paid Amount" + Paid_Amount + "Refund Amount" + Refund_Amount)
                        Debug.Message(Cheque_Cash_Amount + (Paid_Amount - Refund_Amount))
                        Call Update_Sundry("A999", Valid_Combination_Currency(EB_Fund, EB_Bnak_Account), EB_Cheque_Cash_Amount - (EB_Paid_Amount + EB_Refund_Amount))
                        Call Update_Sundry_Two("H003", Valid_Combination_Currency(EB_Fund, EB_Bnak_Account), EB_Refund_Amount)
                    Else
                        Debug.Message("Multi settlement without write off or refund")
                        Call Update_Sundry("A999", Valid_Combination_Currency(EB_Fund, EB_Bnak_Account), EB_Cheque_Cash_Amount - EB_Paid_Amount)
                    End If

                Else
                    'Write_Off /Refund 
                    If Not String.IsNullOrEmpty(Write_Off_Amount) Or Not String.IsNullOrEmpty(Refund_Amount) Then
                        Debug.Message("Write off or Refund") '
                        Call Update_Sundry(Sundry_Account(""), Valid_Combination_Currency(EB_Fund, EB_Bnak_Account), Sundry_Account_Amount(""))
                    Else
                        'Non-Settled Cheque/Cash
                        Debug.Message(" Non Settled Cheque/Cash") '
                        Call Update_Sundry("H001", Valid_Combination_Currency(EB_Fund, EB_Bnak_Account), EB_Cheque_Cash_Amount - EB_Paid_Amount)
                    End If
                End If

            End If

        ElseIf Firstentry = False Then
            Debug.Message("one cheque  multiple bill second entry") '

            '**************1. PAID STATUS *****************************************************************************************
            Wrapper.HLL_CopyStringToPS(Update_OPT(EB_Paid_Status), Wrapper.getPos(14, 57))

            '**************1. PAID AMOUNT *****************************************************************************************
            Wrapper.HLL_CopyStringToPS(EB_Paid_Amount, Wrapper.getPos(14, 61))

            Call Update_Sundry(Sundry_Account(""), Valid_Combination_Currency(EB_Fund, EB_Bnak_Account), Sundry_Account_Amount(""))

            Dim Sundry_Amount As Integer
            If (Multiple_Settlement_Case) Then
                '   Multiple_Settlement with Write Off 
                Debug.Message("Multiple Settlement")
                If Not String.IsNullOrEmpty(EB_Write_Off_Amount) Then
                    Debug.Message("Multiple settlement with write off")
                    Sundry_Amount = Convert.ToInt32(EB_Paid_Amount) + Convert.ToInt32(EB_Write_Off_Amount)
                    Call Update_Sundry("A999", Valid_Combination_Currency(EB_Fund, EB_Bnak_Account), Sundry_Amount.ToString & "-")
                    Call Update_Sundry_Two("H002", Valid_Combination_Currency(EB_Fund, EB_Bnak_Account), Correct_Negative(EB_Write_Off_Amount))
                    '   Multiple_Settlement with refund 
                ElseIf Not String.IsNullOrEmpty(EB_Refund_Amount) Then
                    Sundry_Amount = Convert.ToInt32(EB_Paid_Amount) + Convert.ToInt32(EB_Refund_Amount)
                    Call Update_Sundry("A999", Valid_Combination_Currency(EB_Fund, EB_Bnak_Account), Sundry_Amount.ToString & "-")
                    Call Update_Sundry_Two("H003", Valid_Combination_Currency(EB_Fund, EB_Bnak_Account), EB_Refund_Amount)
                Else
                    Call Update_Sundry("A999", Valid_Combination_Currency(EB_Fund, EB_Bnak_Account), EB_Paid_Amount & "-")
                End If
            End If

            Second_cheque_entry = True
        End If

        Debug.Message(" Third  enter  EB") '
        '**************Press Enter *************************************************************************************
        Wrapper.SendStr("@E")
        ' Wrapper.HLL_Wait()
        Call Check_AC11001_AC10003()
        If Third_Error = True Then
            Debug.Message("Bill exceptionhandling Third Error bill number and sil number " + Bill_No + Sl_No)
            'Moving to Fail 

            'Connection = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & excelfilepath & ";Extended Properties=Excel 12.0 xml;")
            'Connection.Open()
            If Multiple_Settlement_Case Then
                Mul_Settlement_reverse()
            End If
            Debug.Message("exceptionhandling")
            Call exceptionhandling(EB_Status)
            'cheque reversal
            Debug.Message("REVERSAL check condition") '
            If Not String.IsNullOrEmpty(Cheque_No) And Second_cheque_entry = True Then
                Debug.Message("REVERSAL inside") '
                Call Reversal_Checque()
            End If
            Exit Function
        End If
        If (Business_Claim_Refund = True) Then

            '**************False *************************************************************************************
            Business_Claim_Refund = False

            'Move to Success 
            'Call passhandling(Sl_No, Bill_No, Policy_Holder, Cheque_Bank_Code, Cheque_Branch_Code, Cheque_No, Cheque_Date, Reference, Currency, Cheque_Cash_Amount, Paid_Status, Paid_Amount, Receipt, Write_Off_Amount, Refund_Amount, Fund, Bnak_Account, "Claim_Refund")
            Call passhandling("Claim_Refund")
        Else
            'Get Reciept number 
            Call Reciept_number()
        End If
    End Function
    Public Sub Mul_Settlement_reverse()
        Debug.Message("Multiple_Settlement_Case pass connection")
        Dim Sheet1 As New OleDbDataAdapter("Select * From [Pass$]", Connection)
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
            Debug.Message("|" + Cheque_No + "| |" + dn.Tables(0).Rows(i).ItemArray(5) + "|")
            If ((Not (String.IsNullOrEmpty(Cheque_No)) And (GetStringValue(dn.Tables(0).Rows(i).ItemArray(5)) = Cheque_No))) Then
                'If (((Cheque_No IsNot Nothing) Or (Cheque_No IsNot "")) And (dn.Tables(0).Rows(i).ItemArray(5) = Cheque_No)) Then
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

                Call exceptionhandling2(Pass_Sl_No, Pass_Bill_No, Pass_Policy_Holder, Pass_Cheque_Bank_Code, Pass_Cheque_Branch_Code, Pass_Cheque_No, Pass_Cheque_Date, Pass_Reference, Pass_Currency, Pass_Cheque_Cash_Amount, Pass_Paid_Status, Pass_Paid_Amount, Pass_Receipt, Pass_Write_Off_Amount, Pass_Refund_Amount, Pass_Fund, Pass_Bnak_Account, Pass_Description)
                Rows_to_Delete.Push(i)

            End If
        Next
    End Sub
    'Check For ERROR HANDLING for CHEQUE OR CASH PAYMENT ( AC1003to AC11001 Screen)
    Public Sub Check_AC10003()
        Debug.Message("In Check_AC10003 Function")
        Do
            EB_Status = Nothing
            EB_Status = Read_StatusBar()
            If Not EB_Status = Nothing Or Not EB_Status = "" Then
                Wrapper.SendStr("#@")
                Wrapper.HLL_Wait()
                Wrapper.SendStr("@1")
                Wrapper.HLL_Wait()
                First_Error = True
            Else
                Wrapper.HLL_ReadScreen(Wrapper.getPos(1, 2), 7, Proc_Scrn)
                Proc_Scrn = Trim(Strings.Left(Proc_Scrn, 7))
                If Proc_Scrn = "AC11001" Then
                    Exit Do
                End If
            End If
        Loop Until (Not EB_Status = Nothing Or Not EB_Status = "" Or Proc_Scrn = "AC11001")

    End Sub

    Public Sub Check_AC11001()

        Do
            EB_Status = Nothing
            EB_Status = Read_StatusBar()
            If Not EB_Status = Nothing Or Not EB_Status = "" Then
                Debug.Message(" esc")
                Wrapper.SendStr("#@")
                Wrapper.HLL_Wait()
                Debug.Message("F1")
                Wrapper.SendStr("@1")
                Wrapper.HLL_Wait()
                Debug.Message("F1")
                Wrapper.SendStr("@1")
                Wrapper.HLL_Wait()

                Wrapper.HLL_ReadScreen(Wrapper.getPos(1, 2), 7, Proc_Scrn)
                Proc_Scrn = Trim(Strings.Left(Proc_Scrn, 7))
                Debug.Message(Proc_Scrn)
                If Proc_Scrn = "AC10003" Then
                    Debug.Message("F11")
                    Wrapper.SendStr("@b")
                    Wrapper.HLL_Wait()
                End If


                Second_Error = True
            Else
                Exit Do
            End If
        Loop Until (Not EB_Status = Nothing Or Not EB_Status = "")

    End Sub

    Public Sub Check_AC11001_AC10003()
        Do
            Debug.Message("In Check_AC11001_AC10003 Function")
            EB_Status = Nothing
            EB_Status = Read_StatusBar()

            If Not EB_Status = Nothing Or Not EB_Status = "" Then
                Debug.Message("EB Status value")
                Debug.Message("esc")
                Wrapper.SendStr("#@")
                Wrapper.HLL_Wait()
                Debug.Message("F1")
                Wrapper.SendStr("@1")
                Wrapper.HLL_Wait()
                Debug.Message("F1")
                Wrapper.SendStr("@1")
                Wrapper.HLL_Wait()

                Wrapper.HLL_ReadScreen(Wrapper.getPos(1, 2), 7, Proc_Scrn)
                Proc_Scrn = Trim(Strings.Left(Proc_Scrn, 7))
                Debug.Message(Proc_Scrn)
                If Not Proc_Scrn = "AC10001" Then
                    Debug.Message("F11")
                    Wrapper.SendStr("@b")
                    Wrapper.HLL_Wait()
                End If

                Third_Error = True
            Else
                Debug.Message("EB Status Is Nothing")
                Debug.Message("F1")
                Wrapper.SendStr("@1")
                Wrapper.HLL_Wait()
                Wrapper.HLL_ReadScreen(Wrapper.getPos(7, 59), 1, Fund_check_empty_cheque)
                Fund_check_empty_cheque = Trim(Strings.Left(Fund_check_empty_cheque, 1))
                Wrapper.HLL_ReadScreen(Wrapper.getPos(8, 59), 1, Fund_check_empty_cash)
                Fund_check_empty_cash = Trim(Strings.Left(Fund_check_empty_cash, 1))

                Debug.Message(Fund_check_empty_cash + Fund_check_empty_cheque)

                If (Not Fund_check_empty_cash = Nothing Or Not Fund_check_empty_cash = "") Or (Not Fund_check_empty_cheque = Nothing Or Not Fund_check_empty_cheque = "") Then
                    Debug.Message("bill no settled incorrect data")
                    Wrapper.SendStr("@b")
                    Wrapper.HLL_Wait()
                    Fund_check_empty_cash = ""
                    Fund_check_empty_cheque = ""
                    Third_Error = True
                    EB_Status = "Bill Not settled"
                Else
                    If (Business_Claim_Refund = True) Then

                    Else
                        Wrapper.SendStr("@1")
                        Wrapper.HLL_Wait()
                        Debug.Message("Third_Error_check _Flase")
                    End If

                End If
                Exit Do
            End If
        Loop Until (Not EB_Status = Nothing Or Not EB_Status = "")
    End Sub


    'Get string values for Excel value to input FB Fileds validation
    Function GetStringValue(ByVal value As Object) As String
        If value Is DBNull.Value Then
            GetStringValue = Nothing
        Else
            GetStringValue = value
        End If
    End Function


    'GET Reciept number for success billing 
    Public Sub Reciept_number()
        Debug.Message("reciept ")
        Reciept_Error = False
        'Selection Option 3 for BILL ENQUIRY
        Debug.Message("Enter Option 3")
        Wrapper.HLL_CopyStringToPS("3", Wrapper.getPos(23, 12))
        'Press Enter to novigate DCR Capture Scrren
        Wrapper.SendStr("@E")
        Wrapper.HLL_Wait()

        Thread.Sleep(3000)
        Debug.Message("Enter Option 4")
        'Selection Option 1 for BILL ENQUIRY
        Wrapper.HLL_CopyStringToPS("4", Wrapper.getPos(12, 15))
        'ENTER BILL Number
        Debug.Message("Enter Bill Number")
        Wrapper.HLL_CopyStringToPS(Bill_No, Wrapper.getPos(20, 15))
        'Press Enter to novigate DCR Capture Scrren
        Wrapper.SendStr("@E")
        Wrapper.HLL_Wait()

        'Error handling for reciept 
        Call Chech_reciept_error()

        'Error check and exit the from the function
        If Reciept_Error = True Then
            'Moving to Fail 
            'Call exceptionhandling(Sl_No, Bill_No, Policy_Holder, Cheque_Bank_Code, Cheque_Branch_Code, Cheque_No, Cheque_Date, Reference, Currency, Cheque_Cash_Amount, Paid_Status, Paid_Amount, Receipt, Write_Off_Amount, Refund_Amount, Fund, Bnak_Account, EB_Status)
            Call exceptionhandling(EB_Status)
            Exit Sub
        End If

        'Get reciept number
        Dim receipt_number As String = Nothing
        Dim receipt_number1 As String = Nothing
        Wrapper.HLL_ReadScreen(Wrapper.getPos(20, 16), 7, receipt_number1)
        receipt_number1 = Trim(Strings.Left(receipt_number1, 7))
        Dim receipt_number2 As String = Nothing
        Wrapper.HLL_ReadScreen(Wrapper.getPos(21, 16), 7, receipt_number2)
        receipt_number2 = Trim(Strings.Left(receipt_number2, 7))
        Dim receipt_number3 As String = Nothing
        Wrapper.HLL_ReadScreen(Wrapper.getPos(22, 16), 7, receipt_number3)
        receipt_number3 = Trim(Strings.Left(receipt_number3, 7))
        If Not receipt_number3 = Nothing Or Not receipt_number3 = "" Then
            receipt_number = receipt_number3
        ElseIf Not receipt_number2 = Nothing Or Not receipt_number2 = "" Then
            receipt_number = receipt_number2
        ElseIf Not receipt_number1 = Nothing Or Not receipt_number1 = "" Then
            receipt_number = receipt_number1
        End If

        Debug.Message("Receipt Number " & receipt_number)

        'Move to Success 
        'Call passhandling(Sl_No, Bill_No, Policy_Holder, Cheque_Bank_Code, Cheque_Branch_Code, Cheque_No, Cheque_Date, Reference, Currency, Cheque_Cash_Amount, Paid_Status, Paid_Amount, Receipt, Write_Off_Amount, Refund_Amount, Fund, Bnak_Account, receipt_number)
        Call passhandling(receipt_number)

        Debug.Message("esc")
        Wrapper.SendStr("@1")
        Wrapper.HLL_Wait()
        Debug.Message("esc")
        Wrapper.SendStr("@1")
        Wrapper.HLL_Wait()
        Debug.Message("Enter Option 1")
        Wrapper.HLL_CopyStringToPS("1", Wrapper.getPos(23, 12))
        'Press Enter to novigate DCR Capture Scrren
        Wrapper.SendStr("@E")
        Wrapper.HLL_Wait()
    End Sub

    Public Sub Chech_reciept_error()
        Do

            EB_Status = Nothing
            EB_Status = Read_StatusBar()
            Debug.Message(EB_Status)

            If Not EB_Status = Nothing Or Not EB_Status = "" Then
                Wrapper.SendStr("#@")
                Wrapper.HLL_Wait()
                'Selection Option 1 for BILL ENQUIRY
                Wrapper.HLL_CopyStringToPS("", Wrapper.getPos(12, 15))
                'ENTER BILL Number
                Wrapper.HLL_CopyStringToPS("", Wrapper.getPos(20, 15))
                Wrapper.SendStr("@1")
                Wrapper.HLL_Wait()
                'ENTER BILL Number
                Wrapper.HLL_CopyStringToPS("1", Wrapper.getPos(23, 12))
                Wrapper.SendStr("@E")
                Wrapper.HLL_Wait()
                Reciept_Error = True
            Else
                Exit Do
            End If
        Loop Until (Not EB_Status = Nothing Or Not EB_Status = "")
    End Sub

    'Valid Combination of FUNDS and BK A/C and Currency
    Function Valid_Combination_Currency(V_Fund As String, V_Bank_Account As String) As String
        If V_Fund = “G” And V_Bank_Account = “HSBC03” Then
            V_Currency = “HKD”
        ElseIf V_Fund = “G” And V_Bank_Account = “HHC” Then
            V_Currency = “HKD”
        ElseIf V_Fund = “L” And V_Bank_Account = “HSC2” Then
            V_Currency = “HKD”
        ElseIf V_Fund = “L” And V_Bank_Account = “CH” Then
            V_Currency = “HKD”
        ElseIf V_Fund = “L” And V_Bank_Account = “CU” Then
            V_Currency = “USD”
        Else
            Debug.Message("Not a  Valid Currency")
        End If
        Return V_Currency
    End Function

    '**********************************************************************************
    'Function Name  - Update_Sundry() - Autoit Function
    'Parameters     - Sundry , CCY , Amount
    'Description    - Update sundry depending upon the business senarios
    '**********************************************************************************
    Private Sub Update_Sundry(sundry As String, CCY As String, amount As String)
        Debug.Message("IN update Sundry")
        Debug.Message("amount is" + amount)
        '************* SUNDRY-EG: H004   *************************************************************************************
        Wrapper.HLL_CopyStringToPS(sundry, Wrapper.getPos(20, 3))

        '************** SUNDRY CCY -EG:HKD *************************************************************************************
        Wrapper.HLL_CopyStringToPS(CCY, Wrapper.getPos(20, 9))

        '**************SUNDRY AMOUNT -EG:100 *************************************************************************************
        Wrapper.HLL_CopyStringToPS(amount, Wrapper.getPos(20, 13))

    End Sub
    '**********************************************************************************
    'Function Name  - Update_Sundry_Two() - Autoit Function
    'Parameters     - Sundry , CCY , Amount
    'Description    - Update sundry in case of Multiple Settlent with write off or Refund
    '**********************************************************************************
    Private Sub Update_Sundry_Two(sundry As String, CCY As String, amount As String)

        '************* SUNDRY-EG: H004   *************************************************************************************
        Wrapper.HLL_CopyStringToPS(sundry, Wrapper.getPos(20, 29))

        '************** SUNDRY CCY -EG:HKD *************************************************************************************
        Wrapper.HLL_CopyStringToPS(CCY, Wrapper.getPos(20, 35))

        '**************SUNDRY AMOUNT -EG:100 *************************************************************************************
        Wrapper.HLL_CopyStringToPS(amount, Wrapper.getPos(20, 39))

    End Sub

    Function Correct_Negative(Amount As String) As String
        If Amount.Contains("-") = True Then
            Correct_Negative = Trim(Strings.Replace(Amount, "-", "")) & "-"
            Debug.Message(Correct_Negative)
        Else
            Correct_Negative = Amount
        End If
    End Function
    Function Sundry_Account(ByVal type As Object) As String
        If Write_Off_Amount = Nothing And Refund_Amount = Nothing Then
            type = "A999"
        ElseIf Not String.IsNullOrEmpty(Refund_Amount) Then
            type = "H003"
        ElseIf Not String.IsNullOrEmpty(Write_Off_Amount) Then
            type = "H002"
        End If
        Debug.Message("Sundry " & type)
        Return type
    End Function

    'Function G_L(ByRef s_type As Object) As String

    '    Wrapper.HLL_ReadScreen(Wrapper.getPos(14, 2), 10, s_type)
    '    s_type = Trim(Strings.Left(s_type, 10))

    '    If Not s_type Then
    'End Function
    Function Sundry_Account_Amount(ByVal type As Object) As String
        If Write_Off_Amount = Nothing And Refund_Amount = Nothing Then
            Debug.Message("NO WO And NO RF") '
            type = Paid_Amount & "-"
        ElseIf Not String.IsNullOrEmpty(Refund_Amount) Then
            Debug.Message("RF")
            type = Refund_Amount
        ElseIf Not String.IsNullOrEmpty(Write_Off_Amount) Then
            ' MessageBox.Show("WO")
            If Write_Off_Amount.Contains("-") = True Then
                Write_Off_Amount = Trim(Strings.Replace(Write_Off_Amount, "-", "")) & "-"
            Else
                Write_Off_Amount = Write_Off_Amount
            End If
            Debug.Message(Write_Off_Amount)

            type = Write_Off_Amount
        End If
        '  Debug.Message("Sundry_Amount " & type)
        Return type
    End Function

    '**********************************************************************************
    'Function Name  - Read_StatusBar() - Autoit Function
    'Parameters     - Autoit and Window Title
    'Description    - Extract the status from the Window title using the window handle
    'Return         - Status Bar Contents
    '**********************************************************************************
    Private Function Read_StatusBar()

        Thread.Sleep(300)
        Dim status1 As String
        Window_Title = "Session A - [24 x 80]"
        AutoIt.AutoItX.AutoItSetOption("WinTitleMatchMode", 2)
        'Get the status bar text
        status1 = AutoIt.AutoItX.StatusBarGetText(Window_Title)
        status1 = Trim(status1)
        Read_StatusBar = status1

    End Function

    Private Function Update_OPT(opt As String)
        If opt = "F" Then
            opt = ""
        ElseIf opt = "P" Then
            opt = "P"
        End If
        Return opt
    End Function

    Private Sub Reversal_Checque()
        'Add Cheque number - Reversal Cheque
        Debug.Message("Reversal_Checque Cheque Number" + Cheque_No)
        Skip_Cheque.Add(Cheque_No)

        'Selection Option 2 for REVERSAL
        Debug.Message("Reversal Select Option 2")  '
        Wrapper.HLL_CopyStringToPS("2", Wrapper.getPos(14, 43))

        'Press Enter to novigate DCR Capture Scrren
        Debug.Message("Press Enter")   '
        Wrapper.SendStr("@E")
        Wrapper.HLL_Wait()
        Debug.Message("Enter the Bank Code , Branch Code And Cheque Number")

        '**************1. BANK CODE *************************************************************************************
        Wrapper.HLL_CopyStringToPS(Cheque_Bank_Code, Wrapper.getPos(10, 20))

        '**************2. BRANCH CODE ***********************************************************************************
        Wrapper.HLL_CopyStringToPS(Cheque_Branch_Code, Wrapper.getPos(10, 26))

        '**************3. CHECK NUMBER **********************************************************************************
        Wrapper.HLL_CopyStringToPS(Cheque_No, Wrapper.getPos(10, 32))


        Debug.Message("Press ENTER")  '
        'Press enter
        Wrapper.SendStr("@E")
        Wrapper.HLL_Wait()


        'error handling
        Call Chech_reversal_error()
        If Reversal_Error = True Then
            Exit Sub
        End If
        'Press F5
        Debug.Message("Press F5")   '
        Wrapper.SendStr("@5")
        Wrapper.HLL_Wait()
        'Press enter
        Wrapper.SendStr("@E")
        Wrapper.HLL_Wait()
        Debug.Message("The Cheque has been reversed")
    End Sub
    Public Sub Chech_reversal_error()
        Debug.Message("Chech_reversal_error")
        Do
            EB_Status = Nothing
            EB_Status = Read_StatusBar()
            Debug.Message(EB_Status)   '

            If Not EB_Status = Nothing Or Not EB_Status = "" Then
                Wrapper.SendStr("#@")
                Wrapper.HLL_Wait()
                Wrapper.SendStr("@1")
                Wrapper.HLL_Wait()
                Reversal_Error = True
            Else
                Exit Do
            End If
        Loop Until (Not EB_Status = Nothing Or Not EB_Status = "")
    End Sub
    Sub Delete_incorrect_pass(Reverse_List As Stack)
        Debug.Message("Delete_incorrect_pass")
        Dim xlApp As New Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet1 As Excel.Worksheet

        '~~> Opens Workbook. Change path and filename as applicable
        xlWorkBook = xlApp.Workbooks.Open(excelfilepath)

        '~~> Display Excel
        xlApp.Visible = False

        '~~> Set the source worksheet
        xlWorkSheet1 = xlWorkBook.Sheets(2)

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
        ' MessageBox.Show("Exiting Delete_incorrect_pass")
    End Sub

    Public Sub ReleaseComObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        End Try
    End Sub
    Public Sub sign_out_EB()
        Debug.Message("F1") '
        Wrapper.SendStr("@1")
        Wrapper.HLL_Wait()
        'send 90
        Wrapper.HLL_CopyStringToPS("90", Wrapper.getPos(23, 12))
        'press enter
        Wrapper.SendStr("@E")
        Wrapper.HLL_Wait()
    End Sub


    'Exception handling with fail condition and update data to Exceptionsheet in input .xlsx files
    Function exceptionhandling(description As String)
        Debug.Message("Exceptional Handling" + Bill_No + " " + Cheque_Cash_Amount)
        Dim excep_update As New OleDbCommand("insert into [Exception$] ([Sl No],[Bill No],[Policy Holder],[Cheque Bank Code],[Cheque Branch Code],[Cheque No], [Currency],[Cheque Date], [Reference],[Cheque/Cash Amount],[Full/Partial],[Paid Amount], [Receipt],[Write Off Amount],[Refund Amount], [Fund],[Bank Account],[Error Description]) values('" & Sl_No & "','" & Bill_No & "','" & Policy_Holder & "','" & Cheque_Bank_Code & "', '" & Cheque_Branch_Code & "','" & Cheque_No & "','" & Currency & "','" & Cheque_Date & "','" & Reference & "','" & Cheque_Cash_Amount & "','" & Paid_Status & "','" & Paid_Amount & "','" & Receipt & "','" & Write_Off_Amount & "','" & Refund_Amount & "','" & Fund & "','" & Bnak_Account & "','" & description & "')", Connection)
        excep_update.ExecuteNonQuery()
        Dim excep_Status As New OleDbCommand("Update [INPUT_CASES$] Set [Status]='" & "Y" & "' Where [Sl No]='" & Sl_No & "'", Connection)
        excep_Status.ExecuteNonQuery()
    End Function

    'Update data with Pass condition to Passsheet in input .xlsx files
    Function passhandling(reciept_no As String)
        Debug.Message("Pass Handling" + Bill_No + " : " + Cheque_Cash_Amount)
        Dim excep_update As New OleDbCommand("insert into [Pass$] ([Sl No],[Bill No],[Policy Holder],[Cheque Bank Code],[Cheque Branch Code],[Cheque No],[Currency],[Cheque Date],[Reference],[Cheque/Cash Amount],[Full/Partial],[Paid Amount],[Receipt],[Write Off Amount],[Refund Amount],[Fund],[Bank Account],[Receipt Number]) values('" & Sl_No & "','" & Bill_No & "','" & Policy_Holder & "','" & Cheque_Bank_Code & "','" & Cheque_Branch_Code & "','" & Cheque_No & "','" & Currency & "','" & Cheque_Date & "','" & Reference & "','" & Cheque_Cash_Amount & "','" & Paid_Status & "','" & Paid_Amount & "','" & Receipt & "','" & Write_Off_Amount & "','" & Refund_Amount & "','" & Fund & "','" & Bnak_Account & "','" & reciept_no & "')", Connection)
        excep_update.ExecuteNonQuery()
        Dim excep_Status As New OleDbCommand("Update [INPUT_CASES$] Set [Status]='" & "Y" & "' Where [Sl No]='" & Sl_No & "'", Connection)
        excep_Status.ExecuteNonQuery()
    End Function

    'Exception handling with fail condition and update data to Exceptionsheet in input .xlsx files
    Function exceptionhandling2(EB_Sl_No As String, EB_Bill_No As String, EB_Policy_Holder As String, EB_Cheque_Bank_Code As String, EB_Cheque_Branch_Code As String, EB_Cheque_No As String, EB_Cheque_Date As String, Reference As String, EB_Currency As String, EB_Cheque_Cash_Amount As String, EB_Paid_Status As String, EB_Paid_Amount As String, EB_Receipt As String, EB_Write_Off_Amount As String, EB_Refund_Amount As String, EB_Fund As String, EB_Bnak_Account As String, description As String)
        Debug.Message("In Exceptional Handling 2 : " + Bill_No + " " + EB_Cheque_Cash_Amount)
        Dim excep_update As New OleDbCommand("insert into [Exception$] ([Sl No],[Bill No],[Policy Holder],[Cheque Bank Code],[Cheque Branch Code],[Cheque No],[Currency],[Cheque Date],[Reference],[Cheque/Cash Amount],[Full/Partial],[Paid Amount],[Receipt],[Write Off Amount],[Refund Amount],[Fund],[Bank Account],[Error Description]) values('" & EB_Sl_No & "','" & EB_Bill_No & "','" & EB_Policy_Holder & "','" & EB_Cheque_Bank_Code & "', '" & EB_Cheque_Branch_Code & "','" & EB_Cheque_No & "','" & EB_Currency & "','" & EB_Cheque_Date & "','" & Reference & "','" & EB_Cheque_Cash_Amount & "','" & EB_Paid_Status & "','" & EB_Paid_Amount & "','" & Receipt & "','" & EB_Write_Off_Amount & "','" & EB_Refund_Amount & "','" & EB_Fund & "','" & EB_Bnak_Account & "','" & description & "')", Connection)
        excep_update.ExecuteNonQuery()
        Dim excep_Status As New OleDbCommand("Update [INPUT_CASES$] Set [Status]='" & "Y" & "' Where [Sl No]='" & EB_Sl_No & "'", Connection)
        excep_Status.ExecuteNonQuery()
    End Function

    'Update data with Pass condition to Passsheet in input .xlsx files
    Function passhandling2(EB_Sl_No As String, EB_Bill_No As String, EB_Policy_Holder As String, EB_Cheque_Bank_Code As String, EB_Cheque_Branch_Code As String, EB_Cheque_No As String, EB_Cheque_Date As String, Reference As String, EB_Currency As String, EB_Cheque_Cash_Amount As String, EB_Paid_Status As String, EB_Paid_Amount As String, EB_Receipt As String, EB_Write_Off_Amount As String, EB_Refund_Amount As String, EB_Fund As String, EB_Bnak_Account As String, reciept_no As String)
        Debug.Message("Pass Handling" + Bill_No + " : " + EB_Cheque_Cash_Amount)
        Dim excep_update As New OleDbCommand("insert into [Pass$] ([Sl No],[Bill No],[Policy Holder],[Cheque Bank Code],[Cheque Branch Code],[Cheque No],[Currency],[Cheque Date],[Reference],[Cheque/Cash Amount],[Full/Partial],[Paid Amount],[Receipt],[Write Off Amount],[Refund Amount],[Fund],[Bank Account],[Receipt Number]) values('" & EB_Sl_No & "','" & EB_Bill_No & "','" & EB_Policy_Holder & "','" & EB_Cheque_Bank_Code & "','" & EB_Cheque_Branch_Code & "','" & EB_Cheque_No & "','" & EB_Currency & "','" & EB_Cheque_Date & "','" & Reference & "','" & EB_Cheque_Cash_Amount & "','" & EB_Paid_Status & "','" & EB_Paid_Amount & "','" & EB_Receipt & "','" & EB_Write_Off_Amount & "','" & EB_Refund_Amount & "','" & EB_Fund & "','" & EB_Bnak_Account & "','" & reciept_no & "')", Connection)
        excep_update.ExecuteNonQuery()
        Dim excep_Status As New OleDbCommand("Update [INPUT_CASES$] Set [Status]='" & "Y" & "' Where [Sl No]='" & EB_Sl_No & "'", Connection)
        excep_Status.ExecuteNonQuery()
    End Function

End Class
