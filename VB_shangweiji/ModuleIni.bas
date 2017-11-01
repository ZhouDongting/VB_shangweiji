Attribute VB_Name = "ModuleIni"
'ini�ļ����лس����з�������������ԣ�����ҪС��86�ֽڣ�Ӣ����ҪС��143�ֽڲ��ܷ����б�򡣣���������ǰ��code���Ǽ�¼�б�����ݵģ�
Option Explicit
Public iniFileName As String
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'****************************************��ȡIni�ַ���ֵ(Function)******************************************
Function GetIniS(ByVal SectionName As String, ByVal KeyWord As String, ByVal DefString As String) As String
Dim ResultString As String * 144, Temp As Integer
Dim s As String, I As Integer
Temp% = GetPrivateProfileString(SectionName, KeyWord, "", ResultString, 144, AppProFileName(iniFileName))
'�����ؼ��ʵ�ֵ
If Temp% > 0 Then '�ؼ��ʵ�ֵ��Ϊ��
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
'��ȱʡֵд��INI�ļ�
s = DefString
End If
GetIniS = s
End Function

'**************************************��ȡIni��ֵ(Function)***************************************************
Function GetIniN(ByVal SectionName As String, ByVal KeyWord As String, ByVal DefValue As Long) As Integer
Dim d As Long, s As String
d = DefValue
GetIniN = GetPrivateProfileInt(SectionName, KeyWord, DefValue, AppProFileName(iniFileName))
If d <> DefValue Then
s = "" & d
d = WritePrivateProfileString(SectionName, KeyWord, s, AppProFileName(iniFileName))
End If
End Function

'***************************************д���ַ���ֵ(Sub)**************************************************
Sub SetIniS(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValStr As String)
Dim res%
res% = WritePrivateProfileString(SectionName, KeyWord, ValStr, AppProFileName(iniFileName))
End Sub
'****************************************д����ֵ(Sub)******************************************************
Sub SetIniN(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValInt As Long)
Dim res%, s$
s$ = Str$(ValInt)
res% = WritePrivateProfileString(SectionName, KeyWord, s$, AppProFileName(iniFileName))
End Sub


''���������Ѳ�֪���������һ����(keyword) ʱ
'д��һ������ַ���ֵ�Ĺ��̣�����write����д��һ���յ�ֵʵ�ֵģ�'Sub DelIniS(ByVal SectionName As String, ByVal KeyWord As String)
'Dim retval As Integer
'retval = WritePrivateProfileString(SectionName, KeyWord, "", AppProFileName(iniFileName))
'End Sub
'��ʵ0&��ʾǰ���һ����������Ҷ�д��һ����������������section����дһ��Key��һ��������

'***************************************���KeyWord"��"(Sub)*************************************************
Sub DelIniKey(ByVal SectionName As String, ByVal KeyWord As String)
Dim RetVal As Integer
RetVal = WritePrivateProfileString(SectionName, KeyWord, 0&, AppProFileName(iniFileName))
End Sub

'��������section����дһ��Key��һ��������
'**************************************��� Section"��"(Sub)***********************************************
Sub DelIniSec(ByVal SectionName As String) '���section
Dim RetVal As Integer
RetVal = WritePrivateProfileString(SectionName, 0&, "", AppProFileName(iniFileName))
End Sub

'*************************************����Ini�ļ���(Function)***************************************************
'����ini�ļ���
Function AppProFileName(iniFileName)
AppProFileName = App.Path & "\" & iniFileName & ".ini"
End Function







