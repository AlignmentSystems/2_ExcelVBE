Attribute VB_Name = "ReformatMePlease"
Option Explicit
Option Compare Text
Option Base 0
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :
'   Company     :       Alignment Systems Limited
'   Date        :       10th September 2014
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================

Sub TestThisCode()
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :
'   Company     :       Alignment Systems Limited
'   Date        :       24th March 2015
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
'Constants
Const cstrMethodName As String = "ReformatMePlease.TestThisCode  "
Const string1 As String = "Hello"
Const string2 As String = "world"
Const string3 As String = "!"

If Len(string2) = Len(string3) Then
    If Len(string1) = Len(string3) Then
        Debug.Print "blah"
    Else
        Debug.Print "blah blah"
    End If
    Debug.Print "blah blah blah"
Else
    Debug.Print "blah blah blah blah"
End If

End Sub
