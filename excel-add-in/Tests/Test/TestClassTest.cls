VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestClassTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'AccUnit:TestClass

Public Sub TestMe()

   Dim x As Long
   x = 4

   Assert.That x, Iz.EqualTo(5)

End Sub
