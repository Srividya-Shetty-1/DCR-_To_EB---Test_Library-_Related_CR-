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

        p = Process.GetProcessesByName("pcsws")

        If p.Count = 0 Then

            Dim MyProcess As Process
            MyProcess = Process.Start(My.Computer.FileSystem.SpecialDirectories.Desktop & "\DCR.WS")

            MyProcess.WaitForExit(5000)

            Window_Title = MyProcess.MainWindowTitle
        Else
            Window_Title = p(0).MainWindowTitle
        End If

        'CONNECT TO THE PRESENTATION SPACE
        Wrapper.HLL_QuerySession()
        Session_ID = Trim(Strings.Left(Wrapper.mstr_QueryData, 1))
        Wrapper.HLL_ConnectPS(Session_ID)
        'MessageBox.Show(Session_ID)

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

        'Traverse towards library screen
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

        Wrapper.HLL_CopyStringToPS("X", Wrapper.getPos(12, 6))
        Wrapper.SendStr("@E")
        Wrapper.HLL_Wait()
        Wrapper.HLL_CopyStringToPS("1", Wrapper.getPos(23, 12))
        Wrapper.SendStr("@E")
        Wrapper.HLL_Wait()
        Wrapper.HLL_CopyStringToPS("24", Wrapper.getPos(23, 12))
        Wrapper.SendStr("@E")
        Wrapper.HLL_Wait()
        Wrapper.HLL_CopyStringToPS("1", Wrapper.getPos(23, 12))
        Wrapper.SendStr("@E")
        Wrapper.HLL_Wait()

    End Sub
End Class
