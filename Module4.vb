Imports System.ComponentModel
Imports System.IO
Imports System.Linq.Expressions
Imports System.Reflection
Imports SevenUnq

Module SevenLists
     Friend f1 As New SevenUnq.Form1
     Friend unq As New List(Of Integer)
     Friend unqct As New List(Of Integer)
     Friend beforeExe As New List(Of String)
     Friend afterExe As New List(Of String)
     Friend other As New List(Of String)
     Friend ln1 As New List(Of String)
     Friend logs As New List(Of String)
     Friend authorComments As New List(Of String)
     Friend outcome As New List(Of String)
     Friend words() As String
     Friend str1 As String
     Friend plug As String
     Dim whr1 As Integer
     Friend pad10 As String = "           "
     Friend CurrentLine As Integer
     Dim theApplication() As String = System.IO.File.ReadAllLines("C:\Users\vbart\source\REPOS\SevenUnq\original.txt")
     Sub WriteBeforeAndAfter()
          whr1 = unq.IndexOf(CInt(CurrentLine))

          If whr1 = -1 Then
               ' Debugger.Break()
               unq.Add(CurrentLine)
               unqct.Add(unqct.Count - 1)
               nmbr = unqct.Count - 1
               beforeExe.Add("")
               afterExe.Add("")
               other.Add("")
               ln1.Add(theApplication(CurrentLine - 1).Trim)
               logs.Add("")
               authorComments.Add("")
               str1 = theApplication(CurrentLine - 1).ToString.Trim
               words = str1.Split(New Char() {" "c})
               plug = ""
               If words(0).Contains("If") Then
                    beforeExe(nmbr) &= "Comparison about to be made"
                    afterExe(nmbr) &= "if true, transfer control to " & CurrentLine + 1
                    other(nmbr) &= ""
                    logs(nmbr) &= ""
                    authorComments(nmbr) &= ""
                    If nmbr = 2 - 1 Then
                         plug &= pad10 & " sz = " & UserProperties.sz
                         plug &= pad10 & "W(" & UserProperties.sz & ") = " & UserProperties.W(UserProperties.sz)
                         plug &= pad10 & " diff = " & UserProperties.diff
                         other(nmbr) = plug
                    End If
                    If nmbr = 7 - 1 Then
                         plug &= pad10 & " nn = " & UserProperties.ct
                         plug &= pad10 & "W(" & ct & ") = " & W(ct)
                         plug &= pad10 & " diff = " & UserProperties.diff
                         other(nmbr) = plug

                    End If
                    If nmbr = 26 - 1 Then
                         plug &= pad10 & " SUM = " & UserProperties.SUM
                         plug &= pad10 & "First = " & UserProperties.First
                         plug &= pad10 & " (sz * diff) + First " & UserProperties.First + (sz * diff)
                         other(nmbr) = plug
                    End If
               End If
               If words(0).Contains("End") Then
                    beforeExe(nmbr) &= " transfer control to " & CurrentLine + 1
                    afterExe(nmbr) &= " in NXT(), created W "
                    other(nmbr) &= ""
                    logs(nmbr) &= ""
                    authorComments(nmbr) &= ""
               End If
               If words(0).Contains("For") Then
                    beforeExe(nmbr) &= " perform actions in loop certain number of times "
                    afterExe(nmbr) &= " check exit requirements "
                    other(nmbr) &= ""
                    logs(nmbr) &= ""
                    authorComments(nmbr) &= ""
               End If
               If words(0).Contains("Exit") Then
                    beforeExe(nmbr) &= " leave immediately "
                    afterExe(nmbr) &= " exited loop-prevented processing errors"
                    other(nmbr) &= ""
                    logs(nmbr) &= ""
                    authorComments(nmbr) &= ""
               End If

               If words(0).Contains("Next") Then
                    beforeExe(nmbr) &= " check  nn or i "
                    afterExe(nmbr) &= "  last value, exit "
                    other(nmbr) &= ""
                    logs(nmbr) &= ""
                    authorComments(nmbr) &= ""
               End If
               If words(0).Contains("Do") Then
                    beforeExe(nmbr) &= " created ZZ, get new W"
                    afterExe(nmbr) &= " time2quit = true, exit"
                    other(nmbr) &= ""
                    logs(nmbr) &= ""
                    authorComments(nmbr) &= ""
               End If
               If words(0).Contains("Loop") Then
                    beforeExe(nmbr) &= " check time2quit "
                    afterExe(nmbr) &= "if true, no more W to convert "
                    other(nmbr) &= ""
                    logs(nmbr) &= ""
                    authorComments(nmbr) &= ""
               End If
               If words(0).Contains("ct") Then
                    beforeExe(nmbr) &= " local variable "
                    afterExe(nmbr) &= UserProperties.ct
                    other(nmbr) &= ""
                    logs(nmbr) &= ""
                    authorComments(nmbr) &= ""
                    If nmbr = 6 - 1 Then
                         plug &= pad10 & " ct = " & UserProperties.ct
                         other(nmbr) = plug
                    End If
                    If nmbr = 15 - 1 Then
                         plug &= pad10 & " ct = " & UserProperties.ct
                         other(nmbr) = plug
                    End If

               End If
               If words(0).Contains("W(") Then
                    beforeExe(nmbr) &= " W value about to change"
                    Dim arr = CType(_w, Array)
                    afterExe(nmbr) &= "W " & String.Join(Chr(32), _w)
                    other(nmbr) &= ""
                    logs(nmbr) &= ""
                    authorComments(nmbr) &= ""
                    If nmbr = 3 - 1 Then
                         plug &= pad10 & " sz = " & UserProperties.sz
                         plug &= pad10 & "W(" & UserProperties.sz & ") = " & UserProperties.W(UserProperties.sz)
                         other(nmbr) = plug
                    End If
                    If nmbr = 4 - 1 Then
                         plug &= pad10 & " sz = " & UserProperties.sz
                         plug &= pad10 & "W(" & UserProperties.sz & ") = " & UserProperties.W(UserProperties.sz)
                         plug &= pad10 & " diff = " & UserProperties.diff
                         other(nmbr) = plug
                    End If
                    If nmbr = 12 - 1 Then
                         plug &= pad10 & " WHR = " & UserProperties.WHR
                         plug &= pad10 & "W(" & UserProperties.WHR & ") = " & UserProperties.W(UserProperties.WHR)
                         other(nmbr) = plug

                    End If
                    If nmbr = 16 - 1 Then
                         plug &= pad10 & " nn = " & UserProperties.ct
                         plug &= pad10 & "W(" & UserProperties.ct & ") = " & UserProperties.W(UserProperties.ct)
                         plug &= pad10 & " rhs = " & UserProperties.rhs
                         other(nmbr) = plug

                    End If

               End If
               If words(0).Contains("ZZ(") Then
                    beforeExe(nmbr) &= " ZZ value about to change "
                    Dim arr = CType(_zz, Array)
                    afterExe(nmbr) &= "ZZ " & String.Join(Chr(32), _zz)
                    other(nmbr) &= ""
                    logs(nmbr) &= ""
                    authorComments(nmbr) &= ""
                    If nmbr = 24 - 1 Then
                         plug &= pad10 & " nn = " & UserProperties.ct
                         plug &= pad10 & "ZZ(" & UserProperties.ct & ") = " & UserProperties.ZZ(UserProperties.ct)
                         other(nmbr) = plug

                    End If


               End If
               If words(0).Contains("SUM") Then
                    beforeExe(nmbr) &= " update SUM "
                    afterExe(nmbr) &= SUM
                    other(nmbr) &= ""
                    logs(nmbr) &= ""
                    authorComments(nmbr) &= ""
                    If nmbr = 21 - 1 Then
                         plug &= pad10 & " SUM = " & UserProperties.SUM
                         plug &= pad10 & "First = " & UserProperties.First
                         other(nmbr) = plug
                    End If
                    If nmbr = 22 - 1 Then
                         plug &= pad10 & " SUM = " & UserProperties.SUM
                         plug &= pad10 & " nn = " & UserProperties.ct
                         plug &= pad10 & "W(" & UserProperties.ct & ") = " & UserProperties.W(UserProperties.ct)
                         other(nmbr) = plug
                    End If
               End If
               If words(0).Contains("rhs") Then
                    beforeExe(nmbr) &= " update rhs "
                    afterExe(nmbr) &= rhs
                    other(nmbr) &= ""
                    logs(nmbr) &= ""
                    authorComments(nmbr) &= ""
                    If nmbr = 13 - 1 Then
                         plug &= pad10 & " WHR = " & UserProperties.WHR
                         plug &= pad10 & "W(" & UserProperties.WHR & ") = " & UserProperties.W(UserProperties.WHR)
                         other(nmbr) = plug

                    End If
               End If
               If words(0).Contains("Nxt()") Then
                    beforeExe(nmbr) &= " Create new W "
                    afterExe(nmbr) &= " transferred control to first line in Nxt() "
                    other(nmbr) &= ""
                    logs(nmbr) &= ""
                    authorComments(nmbr) &= ""
               End If
               If words(0).Contains("Time2quit") Then
                    beforeExe(nmbr) &= " processed last W "
                    afterExe(nmbr) &= " if true, set Time2quit to true "
                    other(nmbr) &= ""
                    logs(nmbr) &= ""
                    authorComments(nmbr) &= ""
               End If
               If words(0).Contains("BringToFront()") Then
                    beforeExe(nmbr) &= " positioned control "
                    afterExe(nmbr) &= " displayed control "
                    other(nmbr) &= ""
                    logs(nmbr) &= ""
                    authorComments(nmbr) &= ""
               End If
               If words(0).Contains("Me.Text") Then
                    beforeExe(nmbr) &= " assign value to Me.text "
                    afterExe(nmbr) &= " Me.text is checkOutput "
                    other(nmbr) &= ""
                    logs(nmbr) &= ""
                    authorComments(nmbr) &= ""
               End If
               If words(0).Contains("Me.Show()") Then
                    beforeExe(nmbr) &= " prepare to show Me.Text "
                    afterExe(nmbr) &= " displayed Me.Text "
                    other(nmbr) &= ""
                    logs(nmbr) &= ""
                    authorComments(nmbr) &= ""
               End If

          End If

     End Sub
     Sub Main()

          For nn = 1 To theApplication.Count - 1 Step 1
               CurrentLine = nn
               If (nn >= 13 And nn <= 29) _
                                   OrElse (nn >= 33 And nn <= 34) _
                                   OrElse (nn >= 42 And nn <= 43) _
                                   OrElse (nn >= 45 And nn <= 47) _
                                   OrElse (nn >= 75 And nn <= 82) Then
                    WriteBeforeAndAfter()
               End If
          Next

          Using wr1 As StreamWriter = New StreamWriter("C:\Users\vbart\source\repos\SevenUnq\upd2.txt")
               For nn = 0 To unq.Count - 1 Step 1
                    wr1.WriteLine(nn + 1 & "     " & unq(nn) & "     " & ln1(nn))
               Next
          End Using

          Using wr1 As StreamWriter = New StreamWriter("C:\Users\vbart\source\repos\SevenUnq\LINNO.txt")
               For nn = 0 To unq.Count - 1 Step 1
                    wr1.WriteLine("     " & ln1(nn))
               Next
          End Using

          Using wr1 As StreamWriter = New StreamWriter("C:\Users\vbart\source\repos\SevenUnq\before.txt")
               For nn = 0 To unq.Count - 1 Step 1
                    wr1.WriteLine(beforeExe(nn))
               Next
          End Using

          Using wr1 As StreamWriter = New StreamWriter("C:\Users\vbart\source\repos\SevenUnq\after.txt")
               For nn = 0 To unq.Count - 1 Step 1
                    wr1.WriteLine(afterExe(nn))
               Next
          End Using

          Using wr1 As StreamWriter = New StreamWriter("C:\Users\vbart\source\repos\SevenUnq\otherDetails.txt")
               For nn = 0 To unq.Count - 1 Step 1
                    wr1.WriteLine(other(nn))
               Next
          End Using

          Using wr1 As StreamWriter = New StreamWriter("C:\Users\vbart\source\repos\SevenUnq\ac.txt")
               For nn = 0 To unq.Count - 1 Step 1
                    wr1.WriteLine(authorComments(nn))
               Next
          End Using

          '  Environment.Exit(0)

     End Sub
End Module

