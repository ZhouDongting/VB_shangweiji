Attribute VB_Name = "ModuleIni"
'ini文件在有回车换行符会出错，经过测试，汉字要小于86字节，英言文要小于143字节才能返回列表框。（这是我以前的code，是记录列表框内容的）
Option Explicit
Public iniFileName As String
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'****************************************获取Ini字符串值(Function)******************************************
Function GetIniS(ByVal SectionName As String, ByVal KeyWord As String, ByVal DefString As String) As String
Dim ResultString As String * 144, Temp As Integer
Dim s As String, I As Integer
Temp% = GetPrivateProfileString(SectionName, KeyWord, "", ResultString, 144, AppProFileName(iniFileName))
'检索关键词的值
If Temp% > 0 Then '关键词的值不为空
s = ""
For I = 1 To 144
If Asc(Mid$(ResultString, I, 1)) = 0 Then
Exit For
Else
s = s & Mid$(ResultString, I, 1)
End If
Next
Else
Temp% = WritePrivateProfileString(SectionName, KeyWord, DefString, AppProFileName(iniFileName))
'将缺省值写入INI文件
s = DefString
End If
GetIniS = s
End Function

'**************************************获取Ini数值(Function)***************************************************
Function GetIniN(ByVal SectionName As String, ByVal KeyWord As String, ByVal DefValue As Long) As Integer
Dim d As Long, s As String
d = DefValue
GetIniN = GetPrivateProfileInt(SectionName, KeyWord, DefValue, AppProFileName(iniFileName))
If d <> DefValue Then
s = "" & d
d = WritePrivateProfileString(SectionName, KeyWord, s, AppProFileName(iniFileName))
End If
End Function

'***************************************写入字符串值(Sub)**************************************************
Sub SetIniS(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValStr As String)
Dim res%
res% = WritePrivateProfileString(SectionName, KeyWord, ValStr, AppProFileName(iniFileName))
End Sub
'****************************************写入数值(Sub)******************************************************
Sub SetIniN(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValInt As Long)
Dim res%, s$
s$ = Str$(ValInt)
res% = WritePrivateProfileString(SectionName, KeyWord, s$, AppProFileName(iniFileName))
End Sub


''这是我自已不知道怎样清除一个键(keyword) 时
'写的一个清除字符串值的过程，是有write函数写入一个空的值实现的，'Sub DelIniS(ByVal SectionName As String, ByVal KeyWord As String)
'Dim retval As Integer
'retval = WritePrivateProfileString(SectionName, KeyWord, "", AppProFileName(iniFileName))
'End Sub
'其实0&表示前面的一个被清除，我多写了一个“”，如果是清除section就少写一个Key多一个“”。

'***************************************清除KeyWord"键"(Sub)*************************************************
Sub DelIniKey(ByVal SectionName As String, ByVal KeyWord As String)
Dim RetVal As Integer
RetVal = WritePrivateProfileString(SectionName, KeyWord, 0&, AppProFileName(iniFileName))
End Sub

'如果是清除section就少写一个Key多一个“”。
'**************************************清除 Section"段"(Sub)***********************************************
Sub DelIniSec(ByVal SectionName As String) '清除section
Dim RetVal As Integer
RetVal = WritePrivateProfileString(SectionName, 0&, "", AppProFileName(iniFileName))
End Sub

'*************************************定义Ini文件名(Function)***************************************************
'定义ini文件名
Function AppProFileName(iniFileName)
AppProFileName = App.Path & "\" & iniFileName & ".ini"
End Function







