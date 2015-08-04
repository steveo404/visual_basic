Imports System.IO
Imports System.Windows.Forms

Module ReadText

    'Script strips the list of users from an Active Directory group
    'The results will be placed in a text file in the same directory that the script was run
    'Compile commands: vbc <script_name.vb> /out:<output_name>

    Sub Main()
        Dim OpenFileDialog1 As New OpenFileDialog
        Dim fileName As String

        OpenFileDialog1.Title = "Please Select a File"
        OpenFileDialog1.DefaultExt = ".txt"
        OpenFileDialog1.AddExtension = True
        OpenFileDialog1.Filter = "Text Files|*.txt"
        OpenFileDialog1.InitialDirectory = "C:\"
        OpenFileDialog1.ShowDialog()


        Dim strm As System.IO.Stream
        strm = OpenFileDialog1.OpenFile()
        fileName = OpenFileDialog1.FileName.ToString()

        strm.Close()


        Dim myReader As StreamReader = New StreamReader(fileName)
        Dim myWriter As StreamWriter = New StreamWriter("StripUserResults.txt")

        Dim line As String = ""
        Dim strLength As String
        Dim j, i As Integer
        Dim userID As String
        Dim check As String
        Dim trigger As Integer
        Dim counter As Integer

        trigger = 0
        counter = 0

        Do Until IsNothing(line)
            line = myReader.ReadLine()
            If Not IsNothing(line) Then
                check = Mid(line, 1, 5)
                If trigger > 0 Then
                    If line = "The command completed successfully." Then Exit Do
                    strLength = Len(Trim(line))
                    If strLength > 51 Then
                        j = 1
                        For i = 1 To 3
                            userID = Trim(Mid(line, j, 25))
                            myWriter.WriteLine(userID)
                            counter = counter + 1
                            j = j + 25
                        Next
                    End If
                    If (strLength > 31) And (strLength < 51) Then
                        j = 1
                        For i = 1 To 2
                            userID = Trim(Mid(line, j, 25))
                            myWriter.WriteLine(userID)
                            counter = counter + 1
                            j = j + 25
                        Next
                    End If
                    If strLength < 26 Then
                        userID = Trim(Mid(line, j, 25))
                        myWriter.WriteLine(userID)
                        counter = counter + 1
                    End If

                End If
                If check = "-----" Then
                    trigger = 1
                End If
            End If
        Loop

        myReader.Close()
        myWriter.Close()
        Console.WriteLine("Click Enter when Finished")
        Console.WriteLine(counter)
        Console.Readline()

        MsgBox("Script Complete!!!")


    End Sub
End Module
