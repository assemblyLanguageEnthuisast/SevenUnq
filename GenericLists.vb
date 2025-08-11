


Imports System.Reflection

Module GenericLists
     Friend userLocals As New List(Of String) From {"i", "nn"
 }

     Friend vbKeywords As New List(Of String) From {
     "AddHandler", "AddressOf", "Alias", "And", "AndAlso", "As", "Boolean", "ByRef", "Byte",
     "ByVal", "Call", "Case", "Catch", "CBool", "CByte", "CChar", "CDate", "CDbl", "CDec",
     "Char", "CInt", "Class", "Const", "Continue", "CSByte", "CShort", "CSng", "CStr", "CType",
     "CUInt", "CULng", "CUShort", "Date", "Decimal", "Declare", "Default", "Delegate", "Dim",
     "DirectCast", "Do", "Double", "Each", "Else", "ElseIf", "End", "EndIf", "Enum", "Erase",
     "Error", "Event", "Exit", "False", "Finally", "For", "Friend", "Function", "Get", "GetType",
     "Global", "GoSub", "GoTo", "Handles", "If", "Implements", "Imports", "In", "Inherits",
     "Integer", "Interface", "Is", "IsNot", "Let", "Lib", "Like", "Long", "Loop", "Me", "Mod",
     "Module", "MustInherit", "MustOverride", "MyBase", "MyClass", "Namespace", "Narrowing",
     "New", "Next", "Not", "Nothing", "NotInheritable", "NotOverridable", "Object", "Of", "On",
     "Operator", "Option", "Optional", "Or", "OrElse", "Overloads", "Overridable", "Overrides",
     "ParamArray", "Partial", "Private", "Property", "Protected", "Public", "RaiseEvent",
     "ReadOnly", "ReDim", "REM", "RemoveHandler", "Resume", "Return", "SByte", "Select", "Set",
     "Shadows", "Shared", "Short", "Single", "Static", "Step", "Stop", "String", "Structure",
     "Sub", "SyncLock", "Then", "Throw", "To", "True", "Try", "TryCast", "TypeOf", "UInteger",
     "ULong", "UShort", "Using", "Variant", "Wend", "When", "While", "Widening", "With",
     "WithEvents", "WriteOnly", "Xor"
 }
     Friend precedesFields As New List(Of String) From {"Dim", "Private", "Public", "Protected", "Friend",
 "Shared", "Static", "Const", "ReadOnly", "WithEvents"
 }
     Friend userArrays As New List(Of String) From {"W", "ZZ", "_w", "_zz"}
     Friend userConstants As New List(Of String) From {"rw", "sz", "diff", "First"}
     Friend InUPrprts As New List(Of String) From {"f1", "ss", "line", "nmbr", "atPos", "idx",
"ct", "locatedAt", "rhs", "word", "DESCRIPTION", "savedWords", "lineValues", "word_type", "str6",
"arrayCopy", "startHere", "gl", "colons", "leftParen", "RghtParen", "kfal", "MethodRunning", "narratedLines",
"outcome", "File1", "FLNM", "SUM", "WHR", "CurrentLine", "str1", "Narration", "Time2quit"
}
     ReadOnly t As Type = GetType(Form1)
     Friend userFields = t.GetMembers(BindingFlags.DeclaredOnly Or BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic).
               Where(Function(m) m.MemberType = MemberTypes.Field).
               Select(Function(f) f.Name).
               ToList()
     Function GetTokenType(token As String) As String
          If vbKeywords.Contains(token) Then
               Return "Keyword"
          ElseIf IsNumeric(token) Then
               Return "Literal"
          ElseIf token.StartsWith("""") AndAlso token.EndsWith("""") Then
               Return "Literal"
          ElseIf token.StartsWith("'") Then
               Return "Comment"
          ElseIf New List(Of String) From {"+", "-", "*", "/", "=", "<", ">", "And", "Or", "+=", "-=", "*=", "/=", "\=", "<>"}.Contains(token) Then
               Return "Operator"
          ElseIf New List(Of String) From {"(", ")", ",", ".", ":"}.Contains(token) Then
               Return "Punctuation"
          Else
               Return "Identifier"
          End If

     End Function
     Friend BlockStarters As New HashSet(Of String) From {
     "If", "Select", "For", "For Each", "Do", "While", "Try", "With", "SyncLock",
     "Sub", "Function", "Property", "Operator", "Event",
     "Class", "Structure", "Module", "Interface", "Enum", "Namespace"
 }

     Friend BlockEnders As New Dictionary(Of String, String) From {
     {"If", "End If"},
     {"Select", "End Select"},
     {"For", "Next"},
     {"For Each", "Next"},
     {"Do", "Loop"},
     {"While", "End While"},
     {"Try", "End Try"},
     {"With", "End With"},
     {"SyncLock", "End SyncLock"},
     {"Sub", "End Sub"},
     {"Function", "End Function"},
     {"Property", "End Property"},
     {"Operator", "End Operator"},
     {"Event", "End Event"},
     {"Class", "End Class"},
     {"Structure", "End Structure"},
     {"Module", "End Module"},
     {"Interface", "End Interface"},
     {"Enum", "End Enum"},
     {"Namespace", "End Namespace"}
 }

     Friend declaredNames = t.GetMembers().Select(Function(m) m.Name).ToList()
     Friend members = t.GetMembers(BindingFlags.DeclaredOnly Or BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
     Function Description(ByVal theToken As String)
          Description = ""
          Dim isItUserField, isItUserArray, isItUserConstant, doesItPrecedeField, isItaProperty1, isItaProperty2 As Boolean
          Dim isItUserArray2, isitUserArray3 As Boolean
          Dim whatIsIt As String
          Dim descrip As String
          Dim nm As String = theToken
          Dim whr As Integer
          If theToken = Nothing Then
               Exit Function
          End If
          If theToken.Contains("(") Then
               whr = InStr(theToken, "(")
               If whr > 0 Then
                    nm = Mid(theToken, 1, whr - 1)
               End If
          End If

          doesItPrecedeField = GenericLists.precedesFields.Contains(theToken)
          isItUserField = GenericLists.userFields.Contains(theToken)
          isItUserConstant = GenericLists.userConstants.Contains(theToken)
          isItUserArray = GenericLists.userArrays.Contains(nm)
          isItaProperty1 = GenericLists.InUPrprts.Contains(nm)
          isItaProperty2 = GenericLists.InUPrprts.Contains(theToken)
          isitUserArray3 = theToken.Contains("_w,")
          isItUserArray2 = theToken.Contains("_zz,")
          whatIsIt = GenericLists.GetTokenType(theToken)
          If isItUserField = True Then
               descrip = "FIELD"
          ElseIf isItUserConstant = True Then
               descrip = "CONSTANT"
          ElseIf isItUserArray = True Or isItUserArray2 = True Or isitUserArray3 = True Then
               descrip = "ARRAY"
          ElseIf doesItPrecedeField = True Then
               descrip = "ACCESSMODIFIERS"
          ElseIf isItaProperty2 = True Or isItaProperty1 = True Then
               descrip = "PROPERTY"
          Else
               descrip = whatIsIt
          End If
          Description = descrip

     End Function
End Module
