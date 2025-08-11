Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
Imports SevenUnq
Module GroupBoxes
     Public ReadOnly f1 As New SevenUnq.Form1
     Friend CurrentLine As Integer
     Friend od1 As String
     Friend Str1 As String
     Friend actualLine As String
     Friend beforenmbr As Integer
     Friend TorF As String = ""
     Sub GB1()
          f1.GroupBox1.Text = "Program name is cmb_prm." & Chr(10)
          f1.GroupBox1.Text = f1.GroupBox1.Text & "This tiny application lists permutations or combinations via" & Chr(10)
          f1.GroupBox1.Text = f1.GroupBox1.Text & "two arrays W and ZZ." & Chr(10)
          f1.GroupBox1.Text = f1.GroupBox1.Text & "W, the skeleton subset, becomes ZZ." & Chr(10)
          f1.GroupBox1.Text = f1.GroupBox1.Text & "ZZ is created by adding index nn to W(nn)." & Chr(10)
          f1.GroupBox1.Text = f1.GroupBox1.Text & "During runtime, rw, sz, diff, First sum, last sum, summation of 1 thru rw are constants ." & Chr(10)
          f1.GroupBox1.Text = f1.GroupBox1.Text & "RHS determines what enters the output file. These facts make the app reliable." & Chr(10)
          f1.GroupBox1.Text = f1.GroupBox1.Text & "Step thru this application without setting debug breaks or pressing F11." & Chr(10)
          f1.GroupBox1.Text = f1.GroupBox1.Text & "Most instructions will operate on W and ZZ." & Chr(10)
          f1.GroupBox1.Text = f1.GroupBox1.Text & "--PlugAndNarrate may be the best way to understand someone else's code" & Chr(10)
          f1.GroupBox1.Text = f1.GroupBox1.Text & "that is written ENTIRELY in VB.NET" & Chr(10)

          GB2()
          GB3()
          f1.Label1.Visible = False
          f1.Label2.Visible = False
          f1.Label3.Visible = False
          f1.Label4.Visible = False
          f1.Label5.Visible = False
          ' GB4()
          f1.Visible = True
          f1.Show()

          'UserProperties.kfal.Clear()
          Application.DoEvents()
          f1.GroupBox4.Text = Chr(10)
          ' UserProperties.AfterWork.Clear()

     End Sub
     Sub GB2()

          f1.GroupBox2.Visible = True
          f1.GroupBox2.Text = Chr(10)
     End Sub
     Sub GB3()
          f1.GroupBox3.Visible = True
          f1.GroupBox3.Text = Chr(10)
          '   Debugger.Break()
     End Sub
     Sub GB4()

          f1.GroupBox4.Show()
          Dim linno() As String = System.IO.File.ReadAllLines("C:\Users\vbart\source\REPOS\SevenUnq\LINNO.txt")
          Dim upd() As String = System.IO.File.ReadAllLines("C:\Users\vbart\source\REPOS\SevenUnq\upd2.txt")
          Dim od() As String = System.IO.File.ReadAllLines("C:\Users\vbart\source\REPOS\SevenUnq\otherDetails.txt")
          Dim od1 As String = WhichLine(linno, GroupBoxes.Str1)
          'Debugger.Break()

          Dim AFTERLE() As String = System.IO.File.ReadAllLines("C:\Users\vbart\source\REPOS\SevenUnq\after.txt")
          Narration &= Chr(10) & Chr(10) & Now & "  AFTER LINE EXECUTION " & Chr(10) & upd(beforenmbr) & Chr(10) & AFTERLE(beforenmbr) & Chr(10) & od1
          NarrateStep(Narration)
          Application.DoEvents()
          System.Threading.Thread.Sleep(1000)
          f1.GroupBox4.Text = Nothing
          Narration = Nothing

     End Sub
     Public Sub NarrateStep(stepDesc As String)
          f1.GroupBox4.Text &= $"{stepDesc}" & vbCrLf
          Application.DoEvents()
          System.Threading.Thread.Sleep(2000)
     End Sub
     Function WhichLine(linno() As String, input As String)
          Dim narrateOutcome As String = ""

          Dim pad10 As String = "          "
          Dim str3 As String = ""
          If linno(0).Trim = input.Trim Then
               If GroupBoxes.TorF.ToLower = "true" Then
                    str3 &= "transferred control to line 14"
               Else
                    str3 &= "transferred control to line 15"
               End If
               narrateOutcome &= str3
          ElseIf linno(1).Trim = input.Trim Then
               str3 &= "W(" & sz & ") = " & W(sz) & Chr(10)
               narrateOutcome &= str3
          ElseIf linno(2).Trim = input.Trim Then
               If GroupBoxes.TorF.ToLower = "true" Then
                    str3 &= "transferred control to line 16"
               Else
                    str3 &= "transferred control to line 23"
               End If
               narrateOutcome &= str3
          ElseIf linno(3).Trim = input.Trim Then
               narrateOutcome &= str3
          ElseIf linno(4).Trim = input.Trim Then
               str3 &= " ct = " & ct & Chr(10)
               narrateOutcome &= str3
          ElseIf linno(5).Trim = input.Trim Then
               If GroupBoxes.TorF.ToLower = "true" Then
                    str3 &= "transferred control to line 19"
               Else
                    str3 &= "transferred control to line 23"
               End If
               narrateOutcome &= str3
          ElseIf linno(6).Trim = input.Trim Then
               str3 &= " WHR = " & WHR & Chr(10)
               narrateOutcome &= str3
          ElseIf linno(7).Trim = input.Trim Then
               str3 &= "transferred control to line 23"
               narrateOutcome &= str3
          ElseIf linno(8).Trim = input.Trim Then
               str3 &= "transferred control to line 22"
               narrateOutcome &= str3
          ElseIf linno(9).Trim = input.Trim Then
               If GroupBoxes.TorF.ToLower = "true" Then
                    str3 &= "transferred control to line 17"
               Else
                    str3 &= "transferred control to line 23"
               End If
               narrateOutcome &= str3
          ElseIf linno(10).Trim = input.Trim Then
               str3 &= "W(" & WHR & ") = " & W(WHR) & Chr(10)
               narrateOutcome &= str3
          ElseIf linno(11).Trim = input.Trim Then
               str3 &= " rhs =  " & rhs & Chr(10)
               narrateOutcome &= str3
          ElseIf linno(12).Trim = input.Trim Then
               narrateOutcome &= str3
          ElseIf linno(13).Trim = input.Trim Then
               str3 &= " ct = " & ct & Chr(10)
               narrateOutcome &= str3
          ElseIf linno(14).Trim = input.Trim Then
               str3 &= "W(" & ct & ") = " & W(ct) & Chr(10)
               narrateOutcome &= str3
          ElseIf linno(15).Trim = input.Trim Then
               If GroupBoxes.TorF.ToLower = "true" Then
                    str3 &= "transferred control to line 26"
               Else
                    str3 &= "transferred control to line 29"
               End If
               narrateOutcome &= str3
          ElseIf linno(16).Trim = input.Trim Then
               str3 &= "transferred control to line 30"
               narrateOutcome &= str3
          ElseIf linno(17).Trim = input.Trim Then
               If GroupBoxes.TorF.ToLower = "true" Then
                    str3 &= "transferred control to line 34"
               Else
                    str3 &= "transferred control to line 80"
               End If
               narrateOutcome &= str3
          ElseIf linno(18).Trim = input.Trim Then
               Dim arr = CType(_w, Array)
               Dim arr1 = CType(_zz, Array)
               str3 &= "Input:W = " & String.Join(", ", arr.Cast(Of Object)) & Chr(10)
               str3 &= "Input:ZZ = " & String.Join(", ", arr1.Cast(Of Object)) & Chr(10)
               narrateOutcome &= str3
          ElseIf linno(19).Trim = input.Trim Then
               str3 &= " SUM = " & SUM & Chr(10)
               narrateOutcome &= str3
          ElseIf linno(20).Trim = input.Trim Then
               narrateOutcome &= str3
          ElseIf linno(21).Trim = input.Trim Then
               str3 &= " SUM = " & SUM & Chr(10)
               narrateOutcome &= str3
          ElseIf linno(22).Trim = input.Trim Then
               str3 &= " ZZ( " & ct & ") = " & ZZ(ct) & Chr(10)
               narrateOutcome &= str3
          ElseIf linno(23).Trim = input.Trim Then
               If GroupBoxes.TorF.ToLower = "true" Then
                    str3 &= "transferred control to line 44"
               Else
                    str3 &= "transferred control to line 48"
               End If
               narrateOutcome &= str3
          ElseIf linno(24).Trim = input.Trim Then
               str3 &= "transferred control to Nxt()" & Chr(10)
               narrateOutcome &= str3
          ElseIf linno(25).Trim = input.Trim Then
               If GroupBoxes.TorF.ToLower = "true" Then
                    str3 &= "transferred control to line 77"
               Else
                    str3 &= "transferred control to line 79"
               End If
               narrateOutcome &= str3
          ElseIf linno(26).Trim = input.Trim Then
               str3 &= "  Time2quit = True " & Chr(10)
               narrateOutcome &= str3
          ElseIf linno(27).Trim = input.Trim Then
               str3 &= " transferred control to line 79 " & Chr(10)
               narrateOutcome &= str3
          ElseIf linno(28).Trim = input.Trim Then
               If GroupBoxes.TorF.ToLower = "true" Then
                    str3 &= "transferred control to line 80"
               Else
                    str3 &= "transferred control to line 34"
               End If
               narrateOutcome &= str3
          ElseIf linno(29).Trim = input.Trim Then
               str3 &= " control brought to BringToFront" & Chr(10)
               narrateOutcome &= str3
          ElseIf linno(30).Trim = input.Trim Then
               str3 &= "Form1.Text = " & f1.Text & Chr(10)
               narrateOutcome &= str3
          ElseIf linno(31).Trim = input.Trim Then
               str3 &= " Form1 is displayed " & Chr(10)
               narrateOutcome &= str3
          End If
          WhichLine = narrateOutcome
     End Function

End Module
