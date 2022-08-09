Attribute VB_Name = "variables"
Public Const MAX_PATH = 260

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Function Reverse(Data As String) 'Function that reverses the given string
Dim StrReversestring As String
Dim IntCount As Integer

For IntCount = 1 To Len(Data)
    StrReversestring = Mid$(Data, IntCount, 1) & StrReversestring
Next
Reverse = StrReversestring
End Function
