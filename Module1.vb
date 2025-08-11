Imports System.Collections.ObjectModel
Imports System.Drawing.Drawing2D
Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Runtime.Remoting.Messaging
Imports System.Text.RegularExpressions
Module UserProperties
     Public Const rw As Integer = 10
     Public Const sz As Integer = 5
     Public Const diff As Integer = rw - sz
     Public Const First = (diff * (diff + 1)) \ 2
     Public f1 As New SevenUnq.Form1
     Public ReadOnly theApplication() As String = LoadLines()
     Public ss As String
     Public line As String
     Public nmbr As Integer
     Public atPos As Integer
     Public idx As Integer
     Public ct As Integer
     Public locatedAt As Integer
     Public rhs As Integer
     Public word As String
     Public DESCRIPTION As String
     Public savedWords As String()
     Public lineValues As Object()
     Public word_type As New List(Of String)
     Public str6 As String
     Friend arrayCopy As String = "For i As Integer = 0 To sz : ct = i : ZZ(i) = W(i) : Next"
     ' 
     Friend startHere As Integer = 1
     Friend gl As New List(Of Integer)
     Friend colons As New List(Of Integer)
     Friend leftParen As New List(Of Integer)
     Friend RghtParen As New List(Of Integer)
     Friend comma As New List(Of Integer)
     Friend kfal As New List(Of Integer)
     Friend MethodRunning
     Friend narratedLines As New List(Of String)
     ' Public match As Match = Regex.Match(line, "\(([^)]*)\)")
     Friend outcome As String
     Friend AfterWork As New List(Of String)
     Friend onnow As String
     Friend TorF As String
     Friend Function CalculateNumberOf(ByVal str3 As String, ByVal sf As String) As List(Of Integer)
          CalculateNumberOf = gl
          startHere = 1
          gl.Clear()
          locatedAt = InStr(startHere, str3, sf)
          Do Until locatedAt = 0
               gl.Add(locatedAt)
               startHere = locatedAt + 1
               locatedAt = InStr(startHere, str3, sf)
          Loop
          CalculateNumberOf = gl
     End Function
     Friend Function LoadLines() As String()
          Try

               Return System.IO.File.ReadAllLines("C:\Users\vbart\source\REPOS\SevenUnq\original.txt")

          Catch ex As Exception
               Return {}
               Throw New Exception("ReadAllLines Failed")
          End Try
     End Function

     Public Function TryResolveSymbol(name As String) As Object
          Dim t = GetType(UserProperties)
          Dim fi = t.GetField(name, BindingFlags.Public Or BindingFlags.Static)


          If fi IsNot Nothing AndAlso fi.IsLiteral Then

               Return fi.GetRawConstantValue()
          End If

          If fi IsNot Nothing Then

               Return fi.GetValue(Nothing)
          End If

          Dim pi = t.GetProperty(name, BindingFlags.Public Or BindingFlags.Static)

          If pi IsNot Nothing Then

               Return pi.GetValue(Nothing)
          End If


          Return ""
     End Function
     Private _file_ As String
     Public Property File1() As String
          Get
               Return _file_
          End Get
          Set(ByVal value As String)
               _file_ = value
          End Set
     End Property
     Private _FLNM_ As String
     Public Property FLNM() As String
          Get
               Return _FLNM_
          End Get
          Set(ByVal value As String)
               _FLNM_ = value
          End Set
     End Property

     Friend ReadOnly _w(sz) As Integer
     Public Property W(index As Integer) As Integer
          Get
               Return _w(index)
          End Get
          Set(value As Integer)
               _w(index) = value
          End Set
     End Property
     Friend ReadOnly _zz(sz) As Integer
     Public Property ZZ(index As Integer) As Integer
          Get
               Return _zz(index)
          End Get
          Set(value As Integer)
               _zz(index) = value
          End Set
     End Property
     Private _sum_ As Integer
     Public Property SUM() As Integer
          Get
               Return _sum_
          End Get
          Set(ByVal value As Integer)
               _sum_ = value
          End Set
     End Property
     Private _whr_ As Integer
     Public Property WHR() As Integer
          Get
               Return _whr_
          End Get
          Set(ByVal value As Integer)
               _whr_ = value
          End Set
     End Property
     Private _currentLine_ As Integer
     Public Property CurrentLine() As Integer
          Get
               Return _currentLine_
          End Get
          Set(ByVal value As Integer)
               _currentLine_ = value
          End Set
     End Property
     Private _str1_ As String
     Public Property Str1() As String
          Get
               Return _str1_
          End Get
          Set(ByVal value As String)
               _str1_ = value
          End Set
     End Property
     Private _narration_ As String
     Public Property Narration() As String
          Get
               Return _narration_
          End Get
          Set(ByVal value As String)
               _narration_ = value
          End Set
     End Property

     Private _time2quit_ As Boolean = False
     Public Property Time2quit() As Boolean
          Get
               Return _time2quit_
          End Get
          Set(ByVal value As Boolean)
               _time2quit_ = value
          End Set
     End Property
End Module
