Attribute VB_Name = "INI_Util"
Option Explicit

Private Declare Function WritePrivateProfileString Lib "kernel32" _
Alias "WritePrivateProfileStringA" _
                        (ByVal lpApplicationName As String, _
                        ByVal lpKeyName As Any, _
                        ByVal lpString As Any, _
                        ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" _
Alias "GetPrivateProfileStringA" _
                        (ByVal lpApplicationName As String, _
                        ByVal lpKeyName As Any, _
                        ByVal lpDefault As String, _
                        ByVal lpReturnedString As String, _
                        ByVal nSize As Long, _
                        ByVal lpFileName As String) As Long
'usage INIWrite("My_Section","my_Key","My_Value","C:\someinifile.ini")
Public Function INIWrite(sSection As String, sKeyName As String, sNewString As String, sINIFileName As String) As Boolean
  
  Call WritePrivateProfileString(sSection, sKeyName, sNewString, sINIFileName)
  INIWrite = (Err.Number = 0)
End Function
'usage INIRead("My_Section","my_Key","C:\someinifile.ini")
Public Function INIRead(sSection As String, sKeyName As String, sINIFileName As String) As String
Dim sRet As String

  sRet = String(255, Chr(0))
  INIRead = Left(sRet, GetPrivateProfileString(sSection, ByVal sKeyName, "", sRet, Len(sRet), sINIFileName))
End Function

Public Sub save_One_Value_To_INI_File(key As String, value As String)
    Call INIWrite("SUBDB_SUBTITLE_DOWNLOAD", key, value, App.Path & "\SubDB_Subtitle_Download.ini")
End Sub
Public Function load_One_Value_From_INI_File(key As String) As String
    Dim val As String
    val = INIRead("SUBDB_SUBTITLE_DOWNLOAD", key, App.Path & "\SubDB_Subtitle_Download.ini")
    load_One_Value_From_INI_File = val
End Function
