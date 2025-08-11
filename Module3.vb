Imports System.ComponentModel
Imports System.IO
Imports System.Linq.Expressions
Imports System.Reflection
Imports SevenUnq

Module UPDINFOx
     Dim str1 As String
     Dim pt1 As String = ""
     Dim pt2 As String = ""
     Dim pt3 As String = "End If"
     Dim pt4 As String = ""
     Friend words() As String
     Dim whr1 As Integer
     Dim keep As New List(Of String)
     Friend unq As New List(Of String)
     Friend unqct1 As New List(Of Integer)
     Friend inform As New List(Of String)
     Friend review As New List(Of String)
     Friend f1 As New SevenUnq.Form1
     Dim theApplication() As String = System.IO.File.ReadAllLines("C:\Users\vbart\source\REPOS\SevenUnq\original.txt")
     Dim CurrentLine As Integer
     Dim onnow As String
     Sub UpdInfo(ByVal v1 As String)
          Dim whr2 As Integer = 0
          'Dim vet As Integer = 0
          whr2 = unq.IndexOf(v1)
          onnow = v1

          If whr2 = -1 Then
               'Debugger.Break()
               unq.Add(v1)
               unqct1.Add(1)
               inform.Add(Chr(10) & "                                       " & 1 & "      " & str1)
               pt1 = "If UserProperties.onnow = " & Chr(34) & v1 & Chr(34) & " Then " & Chr(10)
               pt2 = "       If UserProperties.CurrentLine = " & CurrentLine & " Then " & Chr(10)
               pt2 &= "          Narration &= GroupBox4.Text" & Chr(10)
               pt2 &= "       " & pt3 & Chr(10)
               pt1 &= pt2 & pt3
               keep.Add(pt1)
          End If
     End Sub
     Sub Main()

          keep.Clear()

          For nn = 1 To theApplication.Count - 1 Step 1
               CurrentLine = nn
               If (nn >= 12 And nn <= 28) _
                                   OrElse (nn >= 32 And nn <= 33) _
                                   OrElse (nn >= 36 And nn <= 41) _
                                   OrElse (nn >= 62 And nn <= 69) Then
                    f1.Text = nn & "   " & theApplication(nn - 1)
                    f1.Show()
                    f1.BringToFront()
                    System.Threading.Thread.Sleep(100)
                    str1 = theApplication(nn - 1).ToString.Trim
                    words = str1.Split(New Char() {" "c})
                    UpdInfo(words(0) & Chr(32) & CurrentLine)
               End If
          Next


          Using wr1 As StreamWriter = New StreamWriter("C:\Users\vbart\source\repos\SevenUnq\upd2.txt")
               For nn = 0 To unq.Count - 1 Step 1
                    wr1.WriteLine(nn + 1 & "     " & unq(nn) & "     " & inform(nn))
               Next
          End Using


          Using wr1 As StreamWriter = New StreamWriter("C:\Users\vbart\source\repos\SevenUnq\nestedIf.txt")
               For nn = 0 To unq.Count - 1 Step 1
                    wr1.WriteLine(keep(nn))
               Next
          End Using

          Using wr1 As StreamWriter = New StreamWriter("C:\Users\vbart\source\repos\SevenUnq\mapForNext.txt")
               For nn = 0 To unq.Count - 1 Step 1
                    If unq(nn).Contains("For") Or unq(nn).Contains("Next") Then
                         wr1.WriteLine(nn + 1 & "     " & unq(nn) & "     " & inform(nn))
                    End If
               Next
          End Using
          'Environment.Exit(0)
          Exit Sub
     End Sub
End Module

