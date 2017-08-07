Attribute VB_Name = "IniFile"
Option Explicit

'API Function to read information from INI File
Public Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any _
    , ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long _
    , ByVal lpFileName As String) As Long

'API Function to write information to the INI File
Private Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any _
    , ByVal lpString As Any, ByVal lpFileName As String) As Long

'Get the INI Setting from the File
Public Function GetINISetting(ByVal sHeading As String, ByVal sKey As String, sINIFileName) As String
    Const cparmLen = 50
    Dim sReturn As String * cparmLen
    Dim sDefault As String * cparmLen
    Dim lLength As Long
    lLength = GetPrivateProfileString(sHeading, sKey _
            , sDefault, sReturn, cparmLen, sINIFileName)
    GetINISetting = Mid(sReturn, 1, lLength)
End Function

'Save INI Setting in the File
Public Function PutINISetting(ByVal sHeading As String, ByVal sKey As String, ByVal sSetting As String, sINIFileName) As Boolean
    On Error GoTo HandleError
    Const cparmLen = 50
    Dim sReturn As String * cparmLen
    Dim sDefault As String * cparmLen
    Dim aLength As Long
    aLength = WritePrivateProfileString(sHeading, sKey _
            , sSetting, sINIFileName)
    PutINISetting = True
    Exit Function
    
HandleError:
    Debug.Print Err.Number & " " & Err.Description
End Function
Public Function Inifile() As String
    Inifile = App.Path & "\LastRunning.ini"
End Function


Public Function GetLastRunning() As String
Dim result As String
result = GetINISetting("LASTRUNNING", "LASTRUNNING", Inifile)
GetLastRunning = result
End Function

Public Function GetValue(ByVal field As String, ByVal index As String) As String
Dim result As String
result = GetINISetting(field, field & index, Inifile)
GetValue = result
End Function

Public Sub SaveLastRunning(ByVal value As String)
    PutINISetting "LASTRUNNING", "LASTRUNNING", value, Inifile
End Sub
Public Sub SaveValue(ByVal field As String, ByVal index As String, ByVal value As String)
    PutINISetting field, field & index, value, Inifile
End Sub

