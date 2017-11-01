Attribute VB_Name = "Module1"


'定义的全局变量
Public hDevice As Long  '句柄
Public data_1(0 To 30000000) As Single
Public data(0 To 300000) As Double
'----------------------------------------my glob var -------------------------------
Public tad_gain As Long
Public tad_stch As Long
Public tad_endch As Long
Public tad_tdata As Long
Public tad_data(0 To 10000000) As Long
Public tad_maxlen As Long
Public tad_total As Long
Public tad_sidi As Long




' 设备函数

Declare Function MP422E_OpenDevice Lib "MP422E.dll " (ByVal DeviceNum As Long) As Long
Declare Function MP422E_CloseDevice Lib "MP422E.dll " (ByVal HANDLE As Long) As Long


'***************************************************************************************

'AD函数

Rem ---------------------------ad at  poll mode ---------------------------------------------------------
Declare Function MP422E_CAL Lib "MP422E.dll " (ByVal HANDLE As Long) As Long

Rem ---------------------- AD at timer Mode -----------------------------------
Declare Function MP422E_AD Lib "MP422E.dll" (ByVal HANDLE As Long, ByVal stch As Long, ByVal endch As Long, ByVal Gain As Long, ByVal sidi As Long, ByVal samode As Long, ByVal trsl As Long, ByVal trpol As Long, ByVal clksl As Long, ByVal clkpol As Long, ByVal tdata As Long) As Long
Declare Function MP422E_Read Lib "MP422E.dll" (ByVal HANDLE As Long, ByVal length As Long, addata As Long) As Long
Declare Function MP422E_Poll Lib "MP422E.dll" (ByVal HANDLE As Long) As Long
Declare Function MP422E_StopAD Lib "MP422E.dll" (ByVal HANDLE As Long) As Long
Declare Function MP422E_ADV Lib "MP422E.dll" (ByVal Gain As Long, ByVal data As Long) As Double


'******************************************************************
'DIO 函数
Declare Function MP422E_DO Lib "MP422E.dll" (ByVal hd As Long, ByVal DO_data As Long) As Long

' out bit, bit state=iobit, sel number=nbit
Declare Function MP422E_DO_Bit Lib "MP422E.dll" (ByVal hd As Long, ByVal iobit As Long, ByVal nbit As Long) As Long

Declare Function MP422E_DI Lib "MP422E.dll" (ByVal hd As Long) As Long

' read bit, bit state=return val, sel number=nbit
Declare Function MP422E_DI_Bit Lib "MP422E.dll" (ByVal hd As Long, ByVal nbit As Long) As Long

'********************************************************************************

'DA函数

Declare Function MP422E_DA_Mode Lib "MP422E.dll" (ByVal hd As Long, ByVal dag0 As Long, ByVal dag1 As Long) As Long
Declare Function MP422E_DA Lib "MP422E.dll" (ByVal hd As Long, ByVal ch As Long, ByVal dadata As Long) As Long
 

Rem --------------- da wave function------------------------------
Declare Function MP422E_DA_WRun Lib "MP422E.dll" (ByVal hd As Long, ByVal dag0 As Long, ByVal dwlen As Long, ByVal datdata As Long, dadata As Long) As Long
Declare Function MP422E_DA_WStop Lib "MP422E.dll" (ByVal hd As Long) As Long
 


'********************************************************************************
'EEPROM函数

Declare Function MP422E_EEPROM_Read Lib "MP422E.dll" (ByVal hd As Long, rbuf As Byte) As Long

Declare Function MP422E_EEPROM_Write Lib "MP422E.dll" (ByVal hd As Long, wbuf As Byte) As Long

'********************************************************************************

'脉冲函数

Declare Function MP422E_PRun Lib "MP422E.dll" (ByVal hd As Long, ByVal pch As Long, ByVal pmode As Long, ByVal pdata0 As Long, ByVal pdata1 As Long) As Long

Declare Function MP422E_PState Lib "MP422E.dll" (ByVal hd As Long, ByVal pch As Long) As Long
 
Declare Function MP422E_PEnd Lib "MP422E.dll" (ByVal hd As Long, ByVal pch As Long) As Long

Declare Function MP422E_PSetData Lib "MP422E.dll" (ByVal hd As Long, ByVal pch As Long, ByVal pdata0 As Long, ByVal pdata1 As Long) As Long


'********************************************************************************
'计数器函数

Declare Function MP422E_CNT_Run Lib "MP422E.dll" (ByVal hd As Long, ByVal cntch As Long, ByVal cntdata As Long) As Long

Declare Function MP422E_CNT_Read Lib "MP422E.dll" (ByVal hd As Long, ByVal cntch As Long, cdata As Long, tdata As Long) As Long
 
'********************************************************************************


