VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExampleClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------
' ExampleClass
'--------------

Public Function Method1() As Long

   Dim x As Long

'1 This line must be ignored.

1  x = 4 ' Line number with code

   x = x + 1

2   Method1 = x

'3 This line must be ignored.

End Function

Public Function Method2() As Long

         Dim x As Long

1        x = 1 + 1
2        Method2 = x

3     Exit Function

4      'This line will never be reached.

End Function
