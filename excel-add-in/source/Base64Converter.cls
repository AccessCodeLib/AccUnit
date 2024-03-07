VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Base64Converter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: text.Base64Converter
'---------------------------------------------------------------------------------------
' Functions to generate BASE64 strings
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>text/Base64Converter.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Explicit
Option Compare Text

Public Function ConvertFileToBase64(ByVal FilePath As String) As String
' Source: https://gist.github.com/cainhill/ad4457dcd16e13f8859a

    Const UseBinaryStreamType = 1

    Dim FileStream As Object
    Dim XmlDoc As Object
    Dim XmlElement As Object
    
    Set XmlDoc = CreateObject("Microsoft.XMLDOM")
    Set XmlElement = XmlDoc.createElement("tmp")
    
    Set FileStream = CreateObject("ADODB.Stream")
    FileStream.Open
    FileStream.Type = UseBinaryStreamType
    FileStream.LoadFromFile FilePath
    
    XmlElement.DataType = "bin.base64"
    XmlElement.NodeTypedValue = FileStream.Read
    Set FileStream = Nothing
    
    ConvertFileToBase64 = Replace(XmlElement.Text, vbLf, vbNullString)
    Set XmlElement = Nothing
    Set XmlDoc = Nothing
    
End Function


Public Sub ConvertBase64ToFile(strFilePath As String, strBase64 As String)
' Source: https://gist.github.com/cainhill/ad4457dcd16e13f8859a

    Const UseBinaryStreamType = 1
    Const SaveWillCreateOrOverwrite = 2

    Dim streamOutput: Set streamOutput = CreateObject("ADODB.Stream")
    Dim XmlDoc: Set XmlDoc = CreateObject("Microsoft.XMLDOM")
    Dim XmlElem: Set XmlElem = XmlDoc.createElement("tmp")
    
    XmlElem.DataType = "bin.base64"
    XmlElem.Text = strBase64
    streamOutput.Open
    streamOutput.Type = UseBinaryStreamType
    streamOutput.Write = XmlElem.NodeTypedValue
    streamOutput.SaveToFile strFilePath, SaveWillCreateOrOverwrite
    
    Set streamOutput = Nothing
    Set XmlDoc = Nothing
    Set XmlElem = Nothing
    
End Sub





