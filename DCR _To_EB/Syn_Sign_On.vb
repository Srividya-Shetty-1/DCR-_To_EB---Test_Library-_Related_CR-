Imports System.Threading
Imports IBMISeries
Imports Microsoft.Office.Interop

Public Class Syn_Sign_On
    'Autoit and Lib Seletion
    Public Session_ID As String
    Public Window_Title As String
    Public Sub Sign_On()
        Dim sign_str As String
        Dim Library_str As String
        Dim Username As String
        Dim Password As String
        Dim p As Process()
        Dim rls_counter As Integer = 0
        Dim TestEnv As Boolean = 1

        p = Process.GetProcessesByName("pcsws")

        If p.Count = 0 Then

            Dim MyProcess As Process
            MyProcess = Process.Start(My.Computer.FileSystem.SpecialDirectories.Desktop & "\HKAS02.ws")

            MyProcess.WaitForExit(5000)

            Window_Title = MyProcess.MainWindowTitle
        Else
            Window_Title = p(0).MainWindowTitle
        End If

        'CONNECT TO THE PRESENTATION SPACE
        Wrapper.HLL_QuerySession()
        Session_ID = Trim(Strings.Left(Wrapper.mstr_QueryData, 1))
        Wrapper.HLL_ConnectPS(Session_ID)
        Debug.Message(Session_ID)

        Dim xlapp As New Excel.Application
        Dim xlwb As Excel.Workbook = xlapp.Workbooks.Open(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\Username and Password - DCR and EB.xlsx", Password:="DCR") 'Excel workbook to store the PDF table
        Dim xlsheet As Excel.Worksheet

        xlsheet = xlwb.Worksheets("Sheet1")

        Username = xlsheet.Range("A2").Value
        Password = xlsheet.Range("B2").Value
        xlwb.Save()
        xlwb.Close()
        xlapp.Quit()
        xlwb = Nothing
        xlapp = Nothing

        If p.Count = 0 Then
            AutoIt.AutoItX.AutoItSetOption("WinTitleMatchMode", 2)
            AutoIt.AutoItX.ControlSetText("IBM i signon", "", "Edit2", Username)
            AutoIt.AutoItX.ControlSetText("IBM i signon", "", "Edit3", Password)
            AutoIt.AutoItX.ControlSend("IBM i signon", "", "Button1", "{ENTER}")
        End If

        'Read the Sign On Screen and Enter the login credentials
        Do
            Thread.Sleep(300)
            sign_str = Nothing
            Wrapper.HLL_ReadScreen(Wrapper.getPos(1, 36), 7, sign_str)
            sign_str = Trim(Strings.Left(sign_str, 7))
            rls_counter = rls_counter + 1

            If sign_str = "Sign On" Then
                Exit Do
            ElseIf rls_counter = 30 Then
                ' MessageBox.Show(Form1, "Tool could not sign-on to EB system" & vbCrLf & "Please try again later", "EB Sign-On Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Application.Exit()
                End
            End If
        Loop Until (sign_str = "Sign On")

        Wrapper.HLL_CopyStringToPS(Username, Wrapper.getPos(6, 53))
        Wrapper.HLL_CopyStringToPS(Password, Wrapper.getPos(7, 53))
        Wrapper.SendStr("@E")
        Wrapper.HLL_Wait()

        Library_str = Nothing

        'Traverse towards library Screen
        Do
            Library_str = Nothing
            Wrapper.HLL_ReadScreen(Wrapper.getPos(2, 2), 7, Library_str)
            Library_str = Trim(Strings.Left(Library_str, 7))
            If Library_str = "SUSEL01" Then
                Exit Do
            Else
                Wrapper.SendStr("@E")
                Wrapper.HLL_Wait()
            End If
        Loop Until (Library_str = "SUSEL01​")

        'Select Library in test enviornment 

        If (TestEnv) Then
            Debug.Message("In test Enviornment")
            Dim Select_Library_str As String
            Dim EB_Status As String
            Dim Library_Found As Boolean = False
            Dim i As Integer = 12
            Dim Lib_name As String = "*NONE"

            'Loop through all library pages
            Do
                'Loop through options in each page
                Do
                    Select_Library_str = Nothing
                    Wrapper.HLL_ReadScreen(Wrapper.getPos(i, 48), 8, Select_Library_str)
                    Select_Library_str = Trim(Strings.Left(Select_Library_str, 7))
                    If Select_Library_str = Lib_name Then
                        Library_Found = True
                        Exit Do
                    End If

                    i = i + 1
                Loop Until (i = 21 Or Select_Library_str = "")

                If (Library_Found = True Or Select_Library_str = "") Then
                    Exit Do
                Else
                    EB_Status = Nothing
                    'Pagedown
                    Wrapper.SendStr("@v")
                    Wrapper.HLL_Wait()
                    EB_Status = Read_StatusBar()
                    If Not EB_Status = Nothing Or Not EB_Status = "" Then
                        Library_Found = False
                        Exit Do
                    Else
                        i = 12
                    End If
                End If
            Loop



            If Library_Found = True Then
                Wrapper.HLL_CopyStringToPS("X", Wrapper.getPos(i, 6))
                Wrapper.SendStr("@E")
                Wrapper.HLL_Wait()
            Else
                MessageBox.Show("Library not found")
                Wrapper.SendStr("@1")
                Wrapper.HLL_Wait()
                'Quit program
                Exit Sub
            End If

            'Wrapper.HLL_CopyStringToPS("X", Wrapper.getPos(12, 6))
            'Wrapper.SendStr("@E")
            'Wrapper.HLL_Wait()

            Wrapper.HLL_CopyStringToPS("1", Wrapper.getPos(23, 12))
                Wrapper.SendStr("@E")
                Wrapper.HLL_Wait()
                Wrapper.HLL_CopyStringToPS("24", Wrapper.getPos(23, 12))
                Wrapper.SendStr("@E")
            Wrapper.HLL_Wait()

            'Select Library in Production enviornment 
        Else
            Dim Select_Library_str As String
            Dim EB_Status As String
            Dim Library_Found As Boolean = False
            Dim i As Integer = 12
            Dim Lib_name As String = "DCRMENUM"

            'Loop through all library pages
            Do
                'Loop through options in each page
                Do
                    Select_Library_str = Nothing
                    Wrapper.HLL_ReadScreen(Wrapper.getPos(i, 48), 8, Select_Library_str)
                    Select_Library_str = Trim(Strings.Left(Select_Library_str, 7))
                    If Select_Library_str = Lib_name Then
                        Library_Found = True
                        Exit Do
                    End If

                    i = i + 1
                Loop Until (i = 21 Or Select_Library_str = "")

                If (Library_Found = True Or Select_Library_str = "") Then
                    Exit Do
                Else
                    EB_Status = Nothing
                    'Pagedown
                    Wrapper.SendStr("@v")
                    Wrapper.HLL_Wait()
                    EB_Status = Read_StatusBar()
                    If Not EB_Status = Nothing Or Not EB_Status = "" Then
                        Library_Found = False
                        Exit Do
                    Else
                        i = 12
                    End If
                End If
            Loop



            If Library_Found = True Then
                Wrapper.HLL_CopyStringToPS("X", Wrapper.getPos(i, 6))
                Wrapper.SendStr("@E")
                Wrapper.HLL_Wait()
            Else
                MessageBox.Show("Library not found")
                Wrapper.SendStr("@1")
                Wrapper.HLL_Wait()
                'Quit program
                Exit Sub
            End If
        End If
        'End If
        Wrapper.HLL_CopyStringToPS("1", Wrapper.getPos(23, 12))
        Wrapper.SendStr("@E")
        Wrapper.HLL_Wait()

    End Sub
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
End Class
