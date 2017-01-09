Imports System.Data.OleDb
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
    'Error handling
    Public First_Error As Boolean = False
    Public Second_Error As Boolean = False
    Public Third_Error As Boolean = False
    Public Reciept_Error As Boolean = False
    Public Reversal_Error As Boolean = False
    Public Firstentry As Boolean = False
    Public Fund_check_empty_cash As String
    Public Fund_check_empty_cheque As String
    'Business check
    Public Business_hang_seng As Boolean = False
    Public Business_Claim_Refund As Boolean = False
    'PublicV_Currency
    Public V_Currency As String
    'Excel file name and path
    Public pathfile As String
    Public excelfilepath As String
    'OLEDB CONNECTION
    Public Connection As OleDbConnection
    Public receipt_number As String
    'Input file's path from shared drive
    Public Sub Read_Input_File(Input_file_path As String)
        'Read Directoryname and information of files
        Dim Dir_Folder As New IO.DirectoryInfo(Input_file_path)
        Dim Dir_Files As IO.FileInfo() = Dir_Folder.GetFiles()
        Dim Dir_File_Info As IO.FileInfo
        'read only .xlsx files 
        For Each Dir_File_Info In Dir_Folder.GetFiles("*.xlsx")
            pathfile = Dir_File_Info.ToString
            If Strings.Right(pathfile, 7) = "op.xlsx" Then
                excelfilepath = Input_file_path & pathfile
                'MessageBox.Show(excelfilepath)
                'OLEDB Connection OPEN / CLOSE
                Connection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & excelfilepath & ";Extended Properties=Excel 12.0 xml;")
                Connection.Open()
                Dim Adap As New OleDbDataAdapter("Select [Sl No],[Bill No],[Policy Holder],[Cheque Bank Code],[Cheque Branch Code],[Cheque No],[Cheque Date],[Reference],[Currency],[Cheque/Cash Amount],[Full/Partial],[Paid Amount],[Receipt],[Write Off Amount],[Refund Amount],[Fund],[Bank Account] From [To_Be_Processed$]", Connection)
                Dim dt As New DataSet
                Adap.Fill(dt)
                For i = 0 To dt.Tables(0).Rows.Count - 1
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

                    'Skip Multiple cheque with Reversal
                    If Not Skip_Cheque.Contains(Cheque_No) Then
                        'Update values to EB system from input file
                        Call EB_Process(Sl_No, Bill_No, Policy_Holder, Cheque_Bank_Code, Cheque_Branch_Code, Cheque_No, Cheque_Date, Reference, Currency, Cheque_Cash_Amount, Paid_Status, Paid_Amount, Receipt, Write_Off_Amount, Refund_Amount, Fund, Bnak_Account)

                    End If
                Next

                'Close OLEDB Connection
                Connection.Dispose()
                Connection.Close()
                Connection = Nothing
                'Clear Array
                arralist.Clear()
                Skip_Cheque.Clear()
                'Updating Summary report sheet
                Dim cls As New Syn_Dcr_Reportvb
                Dim S_path As New Syn_Main
                cls.reporting(excelfilepath, S_path.Shared_drive(), Strings.Left(pathfile, pathfile.Length - 7) & ".xlsx")
                MessageBox.Show("Completed.")
            End If
        Next

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
        ElseIf EB_Cheque_No = Nothing And EB_Reference = Nothing And EB_Write_Off_Amount = Nothing And EB_Refund_Amount = Nothing Then
            Firstentry = True
            'Do nothing
        Else
            Firstentry = False
            EB_Cheque_Cash_Amount = ""
            EB_Cheque_Date = Nothing
            EB_Cheque_Bank_Code = Nothing
            EB_Cheque_Branch_Code = Nothing
            EB_Cheque_No = Nothing
        End If

        'Check Condition for HANG SENG POLICIES
        If EB_Cheque_Bank_Code = Nothing And EB_Cheque_Branch_Code = Nothing And EB_Cheque_No = Nothing And EB_Cheque_Date = Nothing And EB_Currency = Nothing And EB_Cheque_Cash_Amount = Nothing And EB_Write_Off_Amount = Nothing And EB_Refund_Amount = Nothing Then
            MessageBox.Show("Hang Seng")
            Business_hang_seng = True
            '**************6. FUND ******************************************************************************************
            Wrapper.HLL_CopyStringToPS(EB_Fund, Wrapper.getPos(7, 59))

            '**************1. BANK ACCOUNT CODE *****************************************************************************
            Wrapper.HLL_CopyStringToPS(EB_Bnak_Account, Wrapper.getPos(7, 61))

            'Check for CASH PAYMENT 
        ElseIf EB_Cheque_Bank_Code = Nothing And EB_Cheque_Branch_Code = Nothing And EB_Cheque_No = Nothing And EB_Cheque_Date = Nothing Then
            MessageBox.Show("Cash")

            '**************1. CHECK AMOUNT **********************************************************************************
            Wrapper.HLL_CopyStringToPS(EB_Cheque_Cash_Amount, Wrapper.getPos(8, 10))

            '**************6. FUND ******************************************************************************************
            Wrapper.HLL_CopyStringToPS(EB_Fund, Wrapper.getPos(8, 59))

            '**************1. BANK ACCOUNT CODE *****************************************************************************
            Wrapper.HLL_CopyStringToPS(EB_Bnak_Account, Wrapper.getPos(8, 61))

            '**************1. CURRENCY **************************************************************************************
            Wrapper.HLL_CopyStringToPS(Valid_Combination_Currency(Fund, Bnak_Account), Wrapper.getPos(8, 68))
            'Check for CHECK PAYMENT
        ElseIf Not String.IsNullOrEmpty(EB_Cheque_Bank_Code) And Not String.IsNullOrEmpty(EB_Cheque_Branch_Code) And Not String.IsNullOrEmpty(EB_Cheque_No) And Not String.IsNullOrEmpty(EB_Cheque_Date) Then
            MessageBox.Show("Cheque")

            '**************1. CHECK AMOUNT **********************************************************************************
            Wrapper.HLL_CopyStringToPS(EB_Cheque_Cash_Amount, Wrapper.getPos(7, 10))

            '**************2. CHECK DATE (YYYYMMDD) *************************************************************************
            Wrapper.HLL_CopyStringToPS(EB_Cheque_Date, Wrapper.getPos(7, 25))

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

        End If

        'Press Enter for navigation from AC10003 Screen to AC11001
        Wrapper.SendStr("@E")
        Wrapper.HLL_Wait()

        'Check for ERROR handing / Exception Fail case
        Call Check_AC10003()

        If First_Error = True Then
            'Moving to Fail 
            Call exceptionhandling(EB_Sl_No, EB_Bill_No, EB_Policy_Holder, EB_Cheque_Bank_Code, EB_Cheque_Branch_Code, EB_Cheque_No, EB_Cheque_Date, EB_Currency, EB_Cheque_Cash_Amount, EB_Paid_Status, EB_Paid_Amount, EB_Receipt, EB_Write_Off_Amount, EB_Refund_Amount, EB_Fund, EB_Bnak_Account, EB_Status)
            Exit Function
        End If

        'Check cash/cheque with billing or policy
        If EB_Bill_No = Nothing And Not String.IsNullOrEmpty(EB_Policy_Holder) Then
            'Entering to Policy Section
            'MessageBox.Show("Policy")
            Wrapper.HLL_CopyStringToPS(EB_Policy_Holder, Wrapper.getPos(11, 17))
            Business_Claim_Refund = True
        Else
            'Entering to Billing Section
            'MessageBox.Show("Billing")
            Wrapper.HLL_CopyStringToPS(EB_Bill_No, Wrapper.getPos(9, 17))
        End If


        'Press Enter for navigation from AC11001 Screen DCR CAPTURE to Sundry
        Wrapper.SendStr("@E")
        Wrapper.HLL_Wait()

        'Check for ERROR handing / Exception Fail case
        Call Check_AC11001()

        If Second_Error = True Then
            'Moving to Fail 
            Call exceptionhandling(EB_Sl_No, EB_Bill_No, EB_Policy_Holder, EB_Cheque_Bank_Code, EB_Cheque_Branch_Code, EB_Cheque_No, EB_Cheque_Date, EB_Currency, EB_Cheque_Cash_Amount, EB_Paid_Status, EB_Paid_Amount, EB_Receipt, EB_Write_Off_Amount, EB_Refund_Amount, EB_Fund, EB_Bnak_Account, EB_Status)
            Business_hang_seng = False
            Business_Claim_Refund = False
            Exit Function
        End If

        'Sundry logic

        'Get OS Amount from EB System
        Wrapper.HLL_ReadScreen(Wrapper.getPos(14, 39), 15, os_amount)
        os_amount = Trim(Strings.Left(os_amount, 15))

        If (Business_hang_seng = True) Then
            MessageBox.Show("Hang Seng true ")
            'Sundry H004 is Constant value for Hang Seng
            '**************1. PAID STATUS *****************************************************************************************
            Wrapper.HLL_CopyStringToPS(Update_OPT(EB_Paid_Status), Wrapper.getPos(14, 57))

            '**************1. PAID AMOUNT *****************************************************************************************
            Wrapper.HLL_CopyStringToPS(EB_Paid_Amount, Wrapper.getPos(14, 61))

            Call Update_Sundry("H004", Valid_Combination_Currency(EB_Fund, EB_Bnak_Account), EB_Paid_Amount & "-")

            '**************False *************************************************************************************
            Business_hang_seng = False

        ElseIf (Business_Claim_Refund = True) Then
            MessageBox.Show("Claim Refund")
            'Sundry H013 is Constant value for Claim Refund
            Call Update_Sundry("H013", Valid_Combination_Currency(EB_Fund, EB_Bnak_Account), EB_Paid_Amount)

            '**************False *************************************************************************************
            Business_Claim_Refund = False

        ElseIf Firstentry = True Then
            'Normal DCR Capture (1 cheque, 1 bill)
            MessageBox.Show("one cheque  First entry multiple bill first entry")

            '**************1. PAID STATUS *****************************************************************************************
            Wrapper.HLL_CopyStringToPS(Update_OPT(EB_Paid_Status), Wrapper.getPos(14, 57))

            '**************1. PAID AMOUNT *****************************************************************************************
            Wrapper.HLL_CopyStringToPS(EB_Paid_Amount, Wrapper.getPos(14, 61))

            If Not ((EB_Cheque_Cash_Amount - EB_Paid_Amount) = "0") Then
                Call Update_Sundry(Sundry_Account(""), Valid_Combination_Currency(EB_Fund, EB_Bnak_Account), EB_Cheque_Cash_Amount - EB_Paid_Amount)
            End If

        ElseIf Firstentry = False Then
            MessageBox.Show("one cheque  multiple bill second entry")

            '**************1. PAID STATUS *****************************************************************************************
            Wrapper.HLL_CopyStringToPS(Update_OPT(EB_Paid_Status), Wrapper.getPos(14, 57))

            '**************1. PAID AMOUNT *****************************************************************************************
            Wrapper.HLL_CopyStringToPS(EB_Paid_Amount, Wrapper.getPos(14, 61))

            Call Update_Sundry(Sundry_Account(""), Valid_Combination_Currency(EB_Fund, EB_Bnak_Account), Sundry_Account_Amount(""))

        End If

        MessageBox.Show("Press enter")
        '**************Press Enter *************************************************************************************
        Wrapper.SendStr("@E")
        ' Wrapper.HLL_Wait()

        'Check for ERROR handing / Exception Fail case
        Call Check_AC11001_AC10003()

        If Third_Error = True Then
            'Moving to Fail 
            Call exceptionhandling(EB_Sl_No, EB_Bill_No, EB_Policy_Holder, EB_Cheque_Bank_Code, EB_Cheque_Branch_Code, EB_Cheque_No, EB_Cheque_Date, EB_Currency, EB_Cheque_Cash_Amount, EB_Paid_Status, EB_Paid_Amount, EB_Receipt, EB_Write_Off_Amount, EB_Refund_Amount, EB_Fund, EB_Bnak_Account, EB_Status)
            'cheque reversal
            If Not String.IsNullOrEmpty(Cheque_No) Then
                ' MessageBox.Show("REVERSAL")
                Call Reversal_Checque()
            End If
            Exit Function
            End If

            'Get Reciept number 
            Call Reciept_number()

    End Function

    'Check For ERROR HANDLING for CHEQUE OR CASH PAYMENT ( AC1003to AC11001 Screen)
    Public Sub Check_AC10003()
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
                Wrapper.SendStr("#@")
                Wrapper.HLL_Wait()
                Wrapper.SendStr("@1")
                Wrapper.HLL_Wait()
                Wrapper.SendStr("@1")
                Wrapper.HLL_Wait()
                Wrapper.SendStr("@b")
                Wrapper.HLL_Wait()
                Second_Error = True
            Else
                Exit Do
            End If
        Loop Until (Not EB_Status = Nothing Or Not EB_Status = "")

    End Sub

    Public Sub Check_AC11001_AC10003()
        Do

            EB_Status = Nothing
            EB_Status = Read_StatusBar()
            'MessageBox.Show(EB_Status)

            If Not EB_Status = Nothing Or Not EB_Status = "" Then
                '  MessageBox.Show("esc")
                Wrapper.SendStr("#@")
                Wrapper.HLL_Wait()
                ' MessageBox.Show("F1")
                Wrapper.SendStr("@1")
                Wrapper.HLL_Wait()
                ' MessageBox.Show("F1")
                Wrapper.SendStr("@1")
                Wrapper.HLL_Wait()
                ' MessageBox.Show("F11")
                Wrapper.SendStr("@b")
                Wrapper.HLL_Wait()
                'MessageBox.Show("Error_check _True")
                Third_Error = True
            Else
                MessageBox.Show("F1")
                Wrapper.SendStr("@1")
                Wrapper.HLL_Wait()
                Wrapper.HLL_ReadScreen(Wrapper.getPos(7, 59), 1, Fund_check_empty_cheque)
                Fund_check_empty_cheque = Trim(Strings.Left(Fund_check_empty_cheque, 1))
                Wrapper.HLL_ReadScreen(Wrapper.getPos(8, 59), 1, Fund_check_empty_cash)
                Fund_check_empty_cash = Trim(Strings.Left(Fund_check_empty_cash, 1))

                'MessageBox.Show(Fund_check_empty_cash + Fund_check_empty_cheque)

                If (Not Fund_check_empty_cash = Nothing Or Not Fund_check_empty_cash = "") Or (Not Fund_check_empty_cheque = Nothing Or Not Fund_check_empty_cheque = "") Then
                    Wrapper.SendStr("@b")
                    Wrapper.HLL_Wait()
                    Fund_check_empty_cash = ""
                    Fund_check_empty_cheque = ""
                    Third_Error = True
                    MessageBox.Show("Error_check _True")
                Else
                    Wrapper.SendStr("@1")
                    Wrapper.HLL_Wait()
                    MessageBox.Show("Error_check _Flase")
                End If
                ' MessageBox.Show("Exit Do")
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
        'MessageBox.Show("reciept ")
        Reciept_Error = False
        'Selection Option 3 for BILL ENQUIRY
        MessageBox.Show("Enter Option 3")
        Wrapper.HLL_CopyStringToPS("3", Wrapper.getPos(23, 12))
        'Press Enter to novigate DCR Capture Scrren
        Wrapper.SendStr("@E")
        Wrapper.HLL_Wait()

        'MessageBox.Show("Enter option 4")
        'Selection Option 1 for BILL ENQUIRY
        Wrapper.HLL_CopyStringToPS("4", Wrapper.getPos(12, 15))
        'ENTER BILL Number
        ' MessageBox.Show("Enter Bill Number")
        Wrapper.HLL_CopyStringToPS(Bill_No, Wrapper.getPos(20, 15))
        'Press Enter to novigate DCR Capture Scrren
        Wrapper.SendStr("@E")
        Wrapper.HLL_Wait()

        'Error handling for reciept 
        Call Chech_reciept_error()

        'Error check and exit the from the function
        If Reciept_Error = True Then
            'Moving to Fail 
            Call exceptionhandling(Sl_No, Bill_No, Policy_Holder, Cheque_Bank_Code, Cheque_Branch_Code, Cheque_No, Cheque_Date, Currency, Cheque_Cash_Amount, Paid_Status, Paid_Amount, Receipt, Write_Off_Amount, Refund_Amount, Fund, Bnak_Account, EB_Status)
            Exit Sub
        End If
        'Get reciept number

        Dim receipt_number1 As String
        Wrapper.HLL_ReadScreen(Wrapper.getPos(20, 16), 7, receipt_number1)
        receipt_number1 = Trim(Strings.Left(receipt_number1, 7))
        Dim receipt_number2 As String
        Wrapper.HLL_ReadScreen(Wrapper.getPos(21, 16), 7, receipt_number2)
        receipt_number2 = Trim(Strings.Left(receipt_number2, 7))
        Dim receipt_number3 As String
        Wrapper.HLL_ReadScreen(Wrapper.getPos(22, 16), 7, receipt_number3)
        receipt_number3 = Trim(Strings.Left(receipt_number3, 7))
        If Not receipt_number3 = Nothing Or Not receipt_number3 = "" Then
            receipt_number = receipt_number3
        ElseIf Not receipt_number2 = Nothing Or Not receipt_number2 = "" Then
            receipt_number = receipt_number2
        ElseIf Not receipt_number1 = Nothing Or Not receipt_number1 = "" Then
            receipt_number = receipt_number1

        End If


        'MessageBox.Show("Receipt Number " & receipt_number)

        'Move to Success 
        Call passhandling(Sl_No, Bill_No, Policy_Holder, Cheque_Bank_Code, Cheque_Branch_Code, Cheque_No, Cheque_Date, Currency, Cheque_Cash_Amount, Paid_Status, Paid_Amount, Receipt, Write_Off_Amount, Refund_Amount, Fund, Bnak_Account, receipt_number)

        'MessageBox.Show("esc")
        Wrapper.SendStr("@1")
        Wrapper.HLL_Wait()
        'MessageBox.Show("esc")
        Wrapper.SendStr("@1")
        Wrapper.HLL_Wait()
        ' MessageBox.Show("Enter Option 1")
        Wrapper.HLL_CopyStringToPS("1", Wrapper.getPos(23, 12))
        'Press Enter to novigate DCR Capture Scrren
        Wrapper.SendStr("@E")
        Wrapper.HLL_Wait()
    End Sub

    Public Sub Chech_reciept_error()
        Do

            EB_Status = Nothing
            EB_Status = Read_StatusBar()
            ' MessageBox.Show(EB_Status)

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
            'MessageBox.Show("Not a  Valid Currency")
        End If
        Return V_Currency
    End Function

    '**********************************************************************************
    'Function Name  - Update_Sundry() - Autoit Function
    'Parameters     - Sundry , CCY , Amount
    'Description    - Update sundry depending upon the business senarios
    '**********************************************************************************
    Private Sub Update_Sundry(sundry As String, CCY As String, amount As String)

        '************* SUNDRY-EG: H004   *************************************************************************************
        Wrapper.HLL_CopyStringToPS(sundry, Wrapper.getPos(20, 3))

        '************** SUNDRY CCY -EG:HKD *************************************************************************************
        Wrapper.HLL_CopyStringToPS(CCY, Wrapper.getPos(20, 9))

        '**************SUNDRY AMOUNT -EG:100 *************************************************************************************
        Wrapper.HLL_CopyStringToPS(amount, Wrapper.getPos(20, 13))

    End Sub


    Function Sundry_Account(ByVal type As Object) As String
        If Write_Off_Amount = Nothing And Refund_Amount = Nothing Then
            type = "A999"
        ElseIf Not String.IsNullOrEmpty(Refund_Amount) Then
            type = "H003"
        ElseIf Not String.IsNullOrEmpty(Write_Off_Amount) Then
            type = "H002"
        End If
        MessageBox.Show("Sundry " & type)
        Return type
    End Function

    'Function G_L(ByRef s_type As Object) As String

    '    Wrapper.HLL_ReadScreen(Wrapper.getPos(14, 2), 10, s_type)
    '    s_type = Trim(Strings.Left(s_type, 10))

    '    If Not s_type Then
    'End Function
    Function Sundry_Account_Amount(ByVal type As Object) As String
        If Write_Off_Amount = Nothing And Refund_Amount = Nothing Then
            type = Paid_Amount & "-"
        ElseIf Not String.IsNullOrEmpty(Refund_Amount) Then
            type = Refund_Amount
        ElseIf Not String.IsNullOrEmpty(Write_Off_Amount) Then
            Write_Off_Amount = Trim(Strings.Replace(Write_Off_Amount, "-", "")) & "-"
            type = Write_Off_Amount
        End If
        MessageBox.Show("Sundry_Amount " & type)
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
        Skip_Cheque.Add(Cheque_No)
        'Selection Option 2 for REVERSAL
        ' MessageBox.Show("Select option 2")
        Wrapper.HLL_CopyStringToPS("2", Wrapper.getPos(14, 43))
        'Press Enter to novigate DCR Capture Scrren
        ' MessageBox.Show("Press Enter")
        Wrapper.SendStr("@E")
        Wrapper.HLL_Wait()


        '**************1. BANK CODE *************************************************************************************
        Wrapper.HLL_CopyStringToPS(Cheque_Bank_Code, Wrapper.getPos(10, 20))

        '**************2. BRANCH CODE ***********************************************************************************
        Wrapper.HLL_CopyStringToPS(Cheque_Branch_Code, Wrapper.getPos(10, 26))

        '**************3. CHECK NUMBER **********************************************************************************
        Wrapper.HLL_CopyStringToPS(Cheque_No, Wrapper.getPos(10, 32))


        'MessageBox.Show("Press ENTER")
        'Press enter
        Wrapper.SendStr("@E")
        Wrapper.HLL_Wait()


        'error handling
        Call Chech_reversal_error()
        If Reversal_Error = True Then
            Exit Sub
        End If
        'Press F5
        'MessageBox.Show("Press F5")
        Wrapper.SendStr("F5")
        Wrapper.HLL_Wait()
        'Press enter
        Wrapper.SendStr("@E")
        Wrapper.HLL_Wait()

    End Sub
    Public Sub Chech_reversal_error()
        Do
            EB_Status = Nothing
            EB_Status = Read_StatusBar()
            'MessageBox.Show(EB_Status)

            If Not EB_Status = Nothing Or Not EB_Status = "" Then
                ' MessageBox.Show("esc")
                Wrapper.SendStr("#@")
                Wrapper.HLL_Wait()
                ' MessageBox.Show("F1")
                Wrapper.SendStr("@1")
                Wrapper.HLL_Wait()
                Reversal_Error = True
            Else
                Exit Do
            End If
        Loop Until (Not EB_Status = Nothing Or Not EB_Status = "")
    End Sub


    'Exception handling with fail condition and update data to Exceptionsheet in input .xlsx files
    Function exceptionhandling(EB_Sl_No As String, EB_Bill_No As String, EB_Policy_Holder As String, EB_Cheque_Bank_Code As String, EB_Cheque_Branch_Code As String, EB_Cheque_No As String, EB_Cheque_Date As String, EB_Currency As String, EB_Cheque_Cash_Amount As String, EB_Paid_Status As String, EB_Paid_Amount As String, EB_Receipt As String, EB_Write_Off_Amount As String, EB_Refund_Amount As String, EB_Fund As String, EB_Bnak_Account As String, description As String)
        Dim excep_update As New OleDbCommand("insert into [Exception$] ([Sl No],[Bill No],[Policy Holder],[Cheque Bank Code],[Cheque Branch Code],[Cheque No],[Currency],[Cheque Date],[Cheque/Cash Amount],[Full/Partial],[Paid Amount],[Receipt],[Write Off Amount],[Refund Amount],[Fund],[Bank Account],[Error Description]) values('" & EB_Sl_No & "','" & EB_Bill_No & "','" & EB_Policy_Holder & "','" & EB_Cheque_Branch_Code & "', '" & EB_Cheque_Branch_Code & "','" & EB_Cheque_No & "','" & EB_Currency & "','" & EB_Cheque_Date & "','" & EB_Cheque_Cash_Amount & "','" & EB_Paid_Status & "','" & EB_Paid_Amount & "','" & Receipt & "','" & EB_Write_Off_Amount & "','" & EB_Refund_Amount & "','" & EB_Fund & "','" & EB_Bnak_Account & "','" & description & "')", Connection)
        excep_update.ExecuteNonQuery()
        Dim excep_Status As New OleDbCommand("Update [To_Be_Processed$] Set [Status]='" & "Y" & "' Where [Sl No]='" & EB_Sl_No & "'", Connection)
        excep_Status.ExecuteNonQuery()
    End Function

    'Update data with Pass condition to Passsheet in input .xlsx files
    Function passhandling(EB_Sl_No As String, EB_Bill_No As String, EB_Policy_Holder As String, EB_Cheque_Bank_Code As String, EB_Cheque_Branch_Code As String, EB_Cheque_No As String, EB_Cheque_Date As String, EB_Currency As String, EB_Cheque_Cash_Amount As String, EB_Paid_Status As String, EB_Paid_Amount As String, EB_Receipt As String, EB_Write_Off_Amount As String, EB_Refund_Amount As String, EB_Fund As String, EB_Bnak_Account As String, reciept_no As String)
        Dim excep_update As New OleDbCommand("insert into [Pass$] ([Sl No],[Bill No],[Policy Holder],[Cheque Bank Code],[Cheque Branch Code],[Cheque No],[Currency],[Cheque Date],[Cheque/Cash Amount],[Full/Partial],[Paid Amount],[Receipt],[Write Off Amount],[Refund Amount],[Fund],[Bank Account],[Receipt Number]) values('" & EB_Sl_No & "','" & EB_Bill_No & "','" & EB_Policy_Holder & "','" & EB_Cheque_Bank_Code & "','" & EB_Cheque_Branch_Code & "','" & EB_Cheque_No & "','" & EB_Currency & "','" & EB_Cheque_Date & "','" & EB_Cheque_Cash_Amount & "','" & EB_Paid_Status & "','" & EB_Paid_Amount & "','" & EB_Receipt & "','" & EB_Write_Off_Amount & "','" & EB_Refund_Amount & "','" & EB_Fund & "','" & EB_Bnak_Account & "','" & reciept_no & "')", Connection)
        excep_update.ExecuteNonQuery()
        Dim excep_Status As New OleDbCommand("Update [To_Be_Processed$] Set [Status]='" & "Y" & "' Where [Sl No]='" & EB_Sl_No & "'", Connection)
        excep_Status.ExecuteNonQuery()
    End Function


End Class
