Imports System.IO
Imports System.Windows.Forms
Imports System.Threading

Module TermUsers
    'Script Name:       TermUsers
    'Author:            Steve O'Neal
    'Created:           4/17/2015
    'Last Modified:     
    'Version:           1.0
    'Dependency:        
    'Sample script that opens a command line window, 
    'executes a script, and then
    'closes the window.  

    Sub Main()


        ' Start a the command line interface
        Shell("cmd.exe", 3)

        'Pause this script for 5 seconds for the command line to pull up	
        System.Threading.Thread.Sleep(5000)

        ' Make sure the application is in the foreground.
        AppActivate("Administrator: C:\VisualBasic\TermUsers.exe")

        ' Send a text to the application.

        System.Windows.Forms.Sendkeys.SendWait("UserLookUp_cmdline_2.0.bat")
        System.Windows.Forms.Sendkeys.SendWait("{Enter}")

        Dim triggerFile As String
        triggerFile = "C:\Trigger\helloworld.txt"

        Do
            If Len(Dir(triggerFile)) <> 0 Then
                System.Windows.Forms.Sendkeys.SendWait("exit")
                System.Windows.Forms.Sendkeys.SendWait("{Enter}")
                Exit Do
            End If
        Loop

        MsgBox("Batch Processing Complete!!!")


    End Sub
End Module
