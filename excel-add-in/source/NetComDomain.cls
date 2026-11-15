VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NetComDomain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: NetComDomain
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' \brief        Start .NET COM dll without registration (hosts the .NET Framework 4 CLR)
' \ingroup COM
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>COM/NetComDomain.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

' The CLR is started with the hosting interfaces of mscoree.dll (ICLRMetaHost, ICLRRuntimeInfo,
' ICorRuntimeHost). They are called through their vtables (DispCallFunc), so no type library
' reference (mscoree, mscorlib) is needed. The default AppDomain is used late bound.
'
' Why not mscoree.CorRuntimeHost: it starts the legacy CLR 2.0 (needs .NET 3.5, which is an
' optional Windows feature and cannot run assemblies built for .NET Framework 4).

Private Type GuidType
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7) As Byte
End Type

#If VBA7 Then
Private Declare PtrSafe Function CLRCreateInstance Lib "mscoree.dll" ( _
      ByRef ClassId As GuidType, ByRef InterfaceId As GuidType, ByRef Instance As LongPtr) As Long
Private Declare PtrSafe Function CLSIDFromString Lib "ole32.dll" ( _
      ByVal GuidText As LongPtr, ByRef Guid As GuidType) As Long
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" ( _
      ByVal Instance As LongPtr, ByVal VtableOffset As LongPtr, ByVal CallingConvention As Long, _
      ByVal ReturnType As Integer, ByVal ArgumentCount As Long, ByRef ArgumentTypes As Integer, _
      ByRef ArgumentPointers As LongPtr, ByRef ReturnValue As Variant) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
      ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)
#End If

#If Win64 Then
Private Const PointerSize As Long = 8
#Else
Private Const PointerSize As Long = 4
#End If

Private Const ClrVersion As String = "v4.0.30319"
Private Const CallingConventionStdCall As Long = 4
Private Const HostingErrorNumber As Long = vbObjectError + 1

Private Const ClsidClrMetaHost As String = "{9280188D-0E8E-4867-B30C-7FA83884E8DE}"
Private Const IidClrMetaHost As String = "{D332DB9E-B9B3-4125-8207-A14884F53216}"
Private Const IidClrRuntimeInfo As String = "{BD39D1D2-BA2F-486A-89B0-B4B0CB466891}"
Private Const ClsidCorRuntimeHost As String = "{CB2F6723-AB3A-11D2-9C40-00C04FA30A3E}"
Private Const IidCorRuntimeHost As String = "{CB2F6722-AB3A-11D2-9C40-00C04FA30A3E}"

' vtable positions (IUnknown: 0..2)
Private Const IUnknownRelease As Long = 2
Private Const MetaHostGetRuntime As Long = 3
Private Const RuntimeInfoGetInterface As Long = 9
Private Const RuntimeHostStart As Long = 10
Private Const RuntimeHostGetDefaultDomain As Long = 13

Private m_LibraryFile As String
Private m_DefaultNameSpace As String

'---------------------------------------------------------------------------------------
' Function: CreateObject
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Creates a new instance of a class from a .NET COM dll
' </summary>
' <param name="ClassIdentifier">Name of the class without namespace</param>
' <param name="DefaultNameSpace">Namespace of the class</param>
' <param name="LibraryFile">Absolute path of the .NET COM dll</param>
' <returns>Object</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function CreateObject(ByVal ClassIdentifier As String, _
                    Optional ByVal DefaultNameSpace As String = vbNullString, _
                    Optional ByVal LibraryFile As String = vbNullString) As Object

   If VBA.Len(DefaultNameSpace) > 0 Then Me.DefaultNameSpace = DefaultNameSpace
   If VBA.Len(LibraryFile) > 0 Then Me.LibraryFile = LibraryFile

   If VBA.InStr(1, ClassIdentifier, ".") = 0 Then
      If VBA.Len(Me.DefaultNameSpace) = 0 Then
         Err.Raise vbObjectError, "NetComDomain.CreateObject", "Namespace muss angegeben werden!"
      End If
      ClassIdentifier = Me.DefaultNameSpace & "." & ClassIdentifier
   End If

   Dim DefaultDomain As Object
   Set DefaultDomain = GetClrDefaultDomain

   Set CreateObject = DefaultDomain.CreateInstanceFrom(Me.LibraryFile, ClassIdentifier).Unwrap

End Function

Private Function GetClrDefaultDomain() As Object

   Dim MetaHost As LongPtr
   Dim RuntimeInfo As LongPtr
   Dim RuntimeHost As LongPtr
   Dim DomainPointer As LongPtr
   Dim ClrVersionText As String
   Dim ClassId As GuidType
   Dim InterfaceId As GuidType
   Dim RuntimeInfoInterfaceId As GuidType
   Dim RuntimeHostClassId As GuidType
   Dim RuntimeHostInterfaceId As GuidType
   Dim ErrNumber As Long
   Dim ErrSource As String
   Dim ErrDescription As String

On Error GoTo HandleError

   ClassId = GuidFromString(ClsidClrMetaHost)
   InterfaceId = GuidFromString(IidClrMetaHost)
   RuntimeInfoInterfaceId = GuidFromString(IidClrRuntimeInfo)
   RuntimeHostClassId = GuidFromString(ClsidCorRuntimeHost)
   RuntimeHostInterfaceId = GuidFromString(IidCorRuntimeHost)
   ClrVersionText = ClrVersion

   CheckHResult CLRCreateInstance(ClassId, InterfaceId, MetaHost), "CLRCreateInstance"
   CheckHResult CallComMethod(MetaHost, MetaHostGetRuntime, _
                              StrPtr(ClrVersionText), VarPtr(RuntimeInfoInterfaceId), VarPtr(RuntimeInfo)), _
                "ICLRMetaHost.GetRuntime(" & ClrVersion & ")"
   CheckHResult CallComMethod(RuntimeInfo, RuntimeInfoGetInterface, _
                              VarPtr(RuntimeHostClassId), VarPtr(RuntimeHostInterfaceId), VarPtr(RuntimeHost)), _
                "ICLRRuntimeInfo.GetInterface"
   CheckHResult CallComMethod(RuntimeHost, RuntimeHostStart), "ICorRuntimeHost.Start"
   CheckHResult CallComMethod(RuntimeHost, RuntimeHostGetDefaultDomain, VarPtr(DomainPointer)), _
                "ICorRuntimeHost.GetDefaultDomain"

   Set GetClrDefaultDomain = ObjectFromPointer(DomainPointer)

Cleanup:
On Error GoTo 0
   ReleaseComPointer DomainPointer
   ReleaseComPointer RuntimeHost
   ReleaseComPointer RuntimeInfo
   ReleaseComPointer MetaHost
   If ErrNumber <> 0 Then
      Err.Raise ErrNumber, ErrSource, ErrDescription
   End If
   Exit Function

HandleError:
   ErrNumber = Err.Number
   ErrSource = Err.Source
   ErrDescription = Err.Description
   Resume Cleanup

End Function

Private Function GuidFromString(ByVal GuidText As String) As GuidType
   CheckHResult CLSIDFromString(StrPtr(GuidText), GuidFromString), "CLSIDFromString(" & GuidText & ")"
End Function

Private Sub CheckHResult(ByVal HResult As Long, ByVal Operation As String)
   If HResult < 0 Then
      Err.Raise HostingErrorNumber, "NetComDomain", _
                Operation & " failed (HRESULT 0x" & VBA.Hex$(HResult) & ")"
   End If
End Sub

' Calls a method of a COM interface by its vtable position.
' Returns the value returned by the method (HRESULT or reference count).
Private Function CallComMethod(ByVal Instance As LongPtr, ByVal MethodIndex As Long, _
                               ParamArray Arguments() As Variant) As Long

   Dim ArgumentCount As Long
   Dim ArgumentTypes() As Integer
   Dim ArgumentPointers() As LongPtr
   Dim ArgumentValues() As Variant
   Dim ReturnValue As Variant
   Dim Index As Long
   Dim CallResult As Long

   ArgumentCount = UBound(Arguments) + 1

   ' one spare element, so that the first element always exists
   ReDim ArgumentTypes(0 To ArgumentCount)
   ReDim ArgumentPointers(0 To ArgumentCount)
   ReDim ArgumentValues(0 To ArgumentCount)

   For Index = 0 To ArgumentCount - 1
      ArgumentValues(Index) = Arguments(Index)
      ArgumentTypes(Index) = VarType(ArgumentValues(Index))
      ArgumentPointers(Index) = VarPtr(ArgumentValues(Index))
   Next

   CallResult = DispCallFunc(Instance, MethodIndex * PointerSize, CallingConventionStdCall, vbLong, _
                             ArgumentCount, ArgumentTypes(0), ArgumentPointers(0), ReturnValue)
   CheckHResult CallResult, "DispCallFunc(vtable index " & MethodIndex & ")"

   CallComMethod = ReturnValue

End Function

Private Sub ReleaseComPointer(ByRef Pointer As LongPtr)
   If Pointer <> 0 Then
      CallComMethod Pointer, IUnknownRelease
      Pointer = 0
   End If
End Sub

' Wraps an IUnknown pointer into an Object variable. The caller still owns its own reference.
Private Function ObjectFromPointer(ByVal Pointer As LongPtr) As Object

   Dim Unknown As Object
   Dim NullPointer As LongPtr

   CopyMemory Unknown, Pointer, PointerSize      ' no AddRef
   Set ObjectFromPointer = Unknown               ' AddRef
   CopyMemory Unknown, NullPointer, PointerSize  ' no Release when Unknown goes out of scope

End Function

Public Property Get LibraryFile() As String
    LibraryFile = m_LibraryFile
End Property

Public Property Let LibraryFile(ByVal FilePath As String)
    m_LibraryFile = FilePath
End Property

Public Property Get DefaultNameSpace() As String
    DefaultNameSpace = m_DefaultNameSpace
End Property

Public Property Let DefaultNameSpace(ByVal NameSpace As String)
    m_DefaultNameSpace = NameSpace
End Property
