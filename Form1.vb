Imports System.ComponentModel
Imports System.Data.Common
Imports System.Linq.Expressions
Imports System.Net.Security
Imports System.Reflection
Imports System.Text
Imports System.Xml
Imports SevenUnq

Public Class Form1
     Sub Nxt()
          Me.GroupBox2.Text = Chr(10) : Me.GroupBox4.Text = Chr(10)
          SH() : If W(sz) < diff Then : GB4()
               SH() : W(sz) += 1 : GB4()
               SH() : ElseIf W(sz) = diff Then : GB4()
               SH() : For nn = sz To 1 Step -1 : GB4()
                    ct = nn
                    SH() : If W(nn) < diff Then : GB4()
                         SH() : WHR = nn : GB4()
                         SH() : Exit For : GB4()
                         SH() : End If : GB4()
                    SH() : Next : GB4()
               SH() : W(WHR) += 1 : GB4()
               SH() : rhs = W(WHR) : GB4()
               SH() : For nn = WHR + 1 To sz Step 1 : GB4()
                    ct = nn
                    SH() : W(nn) = rhs : GB4()
                    SH() : Next : GB4()
               SH() : End If : GB4()
     End Sub
     Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
          SevenLists.Main() : GB1()
          SH() : Do Until Time2quit = True : GB4()
               SH() : For i As Integer = 0 To sz : ct = i : ZZ(i) = W(i) : Next : GB4()
               Dim arr = CType(_w, Array) : Label1.Text = Nothing : Label2.Text = Nothing : Label3.Text = Nothing : Label4.Text = Nothing : Label5.Text = Nothing
               Label1.Visible = True : GroupBox3.Text = Nothing
               Label1.ForeColor = Color.Black
               Label1.Text = "NXT() creates the next W" & Chr(10)
               Label2.Visible = True
               Label2.ForeColor = Color.Red
               Label2.Text = "Input:W = " & String.Join(", ", arr.Cast(Of Object)) & Chr(10) : Application.DoEvents()
               SH() : SUM = First : GB4()
               SH() : For nn = 1 To diff Step 1 : GB4()
                    ct = nn
                    SH() : SUM += W(nn) : GB4()
                    SH() : ZZ(nn) += nn : GB4()
                    SH() : Next : GB4()
               arr = CType(_zz, Array)
               Label3.Visible = True
               Label3.ForeColor = Color.Black
               Label3.Text = "W is stored into ZZ -- each ZZ(nn) is increased by index nn." & Chr(10)
               Label4.Visible = True
               Label4.ForeColor = Color.Red
               Label4.Text = "Output:ZZ = " & String.Join(", ", arr.Cast(Of Object)) & Chr(10)
               Label4.Text &= "ZZ = "
               For nn = 0 To sz Step 1
                    str6 = "(" & W(nn) & " + " & nn & ") +"
                    If nn = sz Then
                         str6 = "(" & W(nn) & " + " & nn & ") " & Chr(10)
                    End If
                    Label4.Text &= str6 & Chr(32)
               Next
               str6 = ""
               nmbr = First
               Label5.Visible = True
               Label5.ForeColor = Color.Black
               Label5.Text &= "All sums will be greater than or equal to the first sum. " & Chr(10) & " For row = " & rw & " and sz = " & sz & ", that sum is " & First & Chr(10)
               For nn = 0 To sz Step 1
                    nmbr = W(nn) + nmbr
               Next
               Label5.Text &= " ZZ SUM = " & nmbr & Chr(32)
               Me.Show()
               Application.DoEvents()
               System.Threading.Thread.Sleep(300)
               SH() : Nxt() : GB4()
               SH() : If SUM = (sz * diff) + First Then : GB4()
                    SH() : Time2quit = True : GB4()
                    SH() : End If : GB4()
               SH() : Loop : GB4()
          SH() : BringToFront() : GB4()
          SH() : Me.Text = "checkOutput" : GB4()
          SH() : Me.Show() : GB4()

     End Sub
     Public Sub SH()
          UserProperties.CurrentLine = New StackFrame(1, True).GetFileLineNumber()
          SevenLists.CurrentLine = UserProperties.CurrentLine
          UserProperties.Str1 = theApplication(UserProperties.CurrentLine - 1).ToString.Trim
          Dim words = UserProperties.Str1.Split(New Char() {" "c})
          If UserProperties.CurrentLine = 105 OrElse UserProperties.CurrentLine = 29 Then
               GroupBox2.Text = Nothing
               Label1.Visible = False : Label1.Text = Nothing
               Label2.Visible = False : Label2.Text = Nothing
               Label3.Visible = False : Label3.Text = Nothing
               Label4.Visible = False : Label4.Text = Nothing
               Label5.Visible = False : Label5.Text = Nothing
               Application.DoEvents()
          End If
          GroupBox2.Text &= Chr(10) & UserProperties.CurrentLine & "    " & UserProperties.Str1
          GroupBox2.Show()
          GroupBox2.BackColor = System.Drawing.Color.Black
          GroupBox2.ForeColor = System.Drawing.Color.White
          onnow = words(0) & Chr(32) & UserProperties.CurrentLine
          Me.GroupBox4.Text = Nothing
          Narration = Nothing

          Dim ble() As String = System.IO.File.ReadAllLines("C:\Users\vbart\source\REPOS\SevenUnq\before.txt")
          Dim linno() As String = System.IO.File.ReadAllLines("C:\Users\vbart\source\REPOS\SevenUnq\LINNO.txt")
          Dim upd() As String = System.IO.File.ReadAllLines("C:\Users\vbart\source\REPOS\SevenUnq\upd2.txt")
          Dim od() As String = System.IO.File.ReadAllLines("C:\Users\vbart\source\REPOS\SevenUnq\otherDetails.txt")
          Dim od1 As String = WhichLine(linno, UserProperties.Str1)

          GroupBoxes.Str1 = UserProperties.Str1
          For nn = 0 To upd.Count - 1 Step 1
               If upd(nn).Trim.Contains(Chr(32) & UserProperties.CurrentLine & Chr(32)) Then
                    nmbr = nn
                    Exit For
               End If
          Next
          GroupBoxes.beforenmbr = UserProperties.nmbr
          Narration &= Chr(10) & Chr(10) & Now & "  BEFORE LINE EXECUTION " & Chr(10) & upd(nmbr) & Chr(10) & ble(nmbr) & Chr(10) & od1

     End Sub
     Function WhichLine(linno() As String, input As String)
          Dim plugVariables As String = ""
          Dim pad10 As String = "          "
          Dim str3 As String = ""
          If linno(0).Trim = input.Trim Then
               str3 &= " diff = " & diff & Chr(10)
               str3 &= " sz = " & sz & Chr(10)
               str3 &= "W(" & sz & ") = " & W(sz) & Chr(10)
               str3 &= "if true, execute line 14 " & Chr(10)
               str3 &= "if false, goto next comparison" & Chr(10)
               plugVariables &= str3
               UserProperties.TorF = W(sz) < diff
               GroupBoxes.TorF = UserProperties.TorF
          ElseIf linno(1).Trim = input.Trim Then
               str3 &= " sz = " & sz & Chr(10)
               str3 &= "W(" & sz & ") = " & W(sz) & Chr(10)
               str3 &= "last position -sz- in W must never exceed diff" & Chr(10)
               str3 &= "add one" & Chr(10)
               plugVariables &= str3
          ElseIf linno(2).Trim = input.Trim Then
               str3 &= " diff = " & diff & Chr(10)
               str3 &= " sz = " & sz & Chr(10)
               str3 &= "W(" & sz & ") = " & W(sz) & Chr(10)
               str3 &= "if true, find WHR " & Chr(10)
               str3 &= " the way 2 comparisons are worded- THIS SHOULD NEVER BE FALSE " & Chr(10)
               plugVariables &= str3
               UserProperties.TorF = W(sz) = diff
               GroupBoxes.TorF = UserProperties.TorF
          ElseIf linno(3).Trim = input.Trim Then
               str3 &= " nn = name of local variable " & Chr(10)
               str3 &= " sz = " & sz & Chr(10)
               str3 &= " during runtime, sz is a constant value " & Chr(10)
               plugVariables &= str3
          ElseIf linno(4).Trim = input.Trim Then
               str3 &= " ct = " & ct & Chr(10)
               str3 &= " if for loop lose focus-nn value is lost" & Chr(10)
               str3 &= " nn value is stored in ct " & Chr(10)
               plugVariables &= str3
          ElseIf linno(5).Trim = input.Trim Then
               str3 &= " diff = " & diff & Chr(10)
               str3 &= " nn = " & ct & Chr(10)
               str3 &= "W(" & ct & ") = " & W(ct) & Chr(10)
               str3 &= "if true, execute line 19 " & Chr(10)
               str3 &= "if false, goto line 23" & Chr(10)
               plugVariables &= str3
               UserProperties.TorF = W(ct) < diff
               GroupBoxes.TorF = UserProperties.TorF
          ElseIf linno(6).Trim = input.Trim Then
               str3 &= " WHR = " & WHR & Chr(10)
               str3 &= " this is position in W to update that makes the next W " & Chr(10)
               plugVariables &= str3
          ElseIf linno(7).Trim = input.Trim Then
               str3 &= " Leave immediately. WHR has been found" & Chr(10)
               str3 &= " goto line 23" & Chr(10)
               plugVariables &= str3
          ElseIf linno(8).Trim = input.Trim Then
               str3 &= "last comparisons complete, finish making new W" & Chr(10)
               plugVariables &= str3
          ElseIf linno(9).Trim = input.Trim Then
               str3 &= "nn is greater than  1, goto line 17" & Chr(10)
               str3 &= "nn is equal  1, goto line 23" & Chr(10)
               plugVariables &= str3
               UserProperties.TorF = ct > 1
               GroupBoxes.TorF = UserProperties.TorF
          ElseIf linno(10).Trim = input.Trim Then
               str3 &= " WHR = " & WHR & Chr(10)
               str3 &= "W(" & WHR & ") = " & W(WHR) & Chr(10)
               str3 &= " WHR in W must never exceed diff" & Chr(10)
               str3 &= "add one" & Chr(10)
               plugVariables &= str3
          ElseIf linno(11).Trim = input.Trim Then
               str3 &= " rhs =  W(" & WHR & ")" & Chr(10)
               plugVariables &= str3
          ElseIf linno(12).Trim = input.Trim Then
               str3 &= " nn = name of local variable " & Chr(10)
               str3 &= " WHR = " & WHR & Chr(10)
               str3 &= " during runtime, WHR will change " & Chr(10)
               str3 &= " sz = " & sz & Chr(10)
               plugVariables &= str3
          ElseIf linno(13).Trim = input.Trim Then
               str3 &= " ct = " & ct & Chr(10)
               str3 &= " if for loop lose focus-nn value is lost" & Chr(10)
               str3 &= " nn value is stored in ct " & Chr(10)
               plugVariables &= str3
          ElseIf linno(14).Trim = input.Trim Then
               str3 &= " nn = " & ct & Chr(10)
               str3 &= "W(" & ct & ") = " & W(ct) & Chr(10)
               str3 &= " rhs =  " & rhs & Chr(10)
               plugVariables &= str3
          ElseIf linno(15).Trim = input.Trim Then
               str3 &= "nn is lesser than  sz, goto line 26" & Chr(10)
               str3 &= "nn is equal  sz, goto line 29" & Chr(10)
               plugVariables &= str3
               UserProperties.TorF = ct < sz
               GroupBoxes.TorF = UserProperties.TorF
          ElseIf linno(16).Trim = input.Trim Then
               str3 &= "both comparisons complete, exit Nxt()" & Chr(10)
               plugVariables &= str3
          ElseIf linno(17).Trim = input.Trim Then
               str3 &= "Time2quit decides if control is passed to line 34 or line 80" & Chr(10)
               UserProperties.TorF = Time2quit = True
               GroupBoxes.TorF = UserProperties.TorF
               plugVariables &= str3
          ElseIf linno(18).Trim = input.Trim Then
               str3 &= "clever AI way of replacing Array.Copy(_w,_zz,sz+1)" & Chr(10)
               plugVariables &= str3
          ElseIf linno(19).Trim = input.Trim Then
               str3 &= " SUM = " & SUM & Chr(10)
               str3 &= " First - a constant value- is assigned to SUM" & Chr(10)
               str3 &= " there are a number of things that can be calculated " & Chr(10)
               str3 &= " first sum, last sum, the sum of 1 thru row " & Chr(10)
               str3 &= " number of combinations and permutations are also calculable - finite mathematics " & Chr(10)
               plugVariables &= str3
          ElseIf linno(20).Trim = input.Trim Then
               str3 &= " nn = name of local variable " & Chr(10)
               str3 &= " diff = " & diff & Chr(10)
               str3 &= " loop that writes to ZZ " & Chr(10)
               plugVariables &= str3
          ElseIf linno(21).Trim = input.Trim Then
               str3 &= " SUM = " & SUM & Chr(10)
               str3 &= " increase SUM by W(" & ct & ")" & Chr(10)
               plugVariables &= str3
          ElseIf linno(22).Trim = input.Trim Then
               str3 &= " ZZ( " & ct & ") = " & ZZ(ct) & Chr(10)
               str3 &= " increase ZZ(" & ct & ") by " & ct & Chr(10)
               plugVariables &= str3
          ElseIf linno(23).Trim = input.Trim Then
               str3 &= "nn is lesser than  diff, goto line 44" & Chr(10)
               str3 &= "nn is equal  diff, goto line 48" & Chr(10)
               plugVariables &= str3
               UserProperties.TorF = ct < diff
               GroupBoxes.TorF = UserProperties.TorF
          ElseIf linno(24).Trim = input.Trim Then
               str3 &= "transfer control to Nxt()" & Chr(10)
               plugVariables &= str3
          ElseIf linno(25).Trim = input.Trim Then
               str3 &= "is this the last SUM?" & Chr(10)
               str3 &= " SUM = " & SUM & Chr(10)
               str3 &= " compare SUM to (sz * diff) + First " & SUM & Chr(10)
               str3 &= " that number equals " & (sz * diff) + First & Chr(10)
               plugVariables &= str3
               UserProperties.TorF = SUM = (sz * diff) + First
               GroupBoxes.TorF = UserProperties.TorF
          ElseIf linno(26).Trim = input.Trim Then
               str3 &= " last sum has been found " & Chr(10)
               str3 &= " set Time2quit to true " & Chr(10)
               plugVariables &= str3

          ElseIf linno(27).Trim = input.Trim Then
               str3 &= "last sum comparison complete, goto line 79" & Chr(10)
               plugVariables &= str3
          ElseIf linno(28).Trim = input.Trim Then
               str3 &= "Time2quit = false, goto line 34 " & Chr(10)
               str3 &= "Time2quit = true, goto line 80" & Chr(10)
               plugVariables &= str3
               UserProperties.TorF = Time2quit = True
               GroupBoxes.TorF = UserProperties.TorF
          ElseIf linno(29).Trim = input.Trim Then
               str3 &= "transfer control to BringToFront" & Chr(10)
               str3 &= " make control e.g. Form1 visible via z-order" & Chr(10)
               plugVariables &= str3
          ElseIf linno(30).Trim = input.Trim Then
               str3 &= "assign a value to Form1.Text" & Chr(10)
               str3 &= " circular reference is prevented by using Me.Text" & Chr(10)
               plugVariables &= str3
          ElseIf linno(31).Trim = input.Trim Then
               str3 &= "transfer control to Form1.Show" & Chr(10)
               str3 &= " circular reference is prevented by using Me.show()" & Chr(10)
               plugVariables &= str3

          End If
          WhichLine = plugVariables
     End Function

     Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

     End Sub
End Class
