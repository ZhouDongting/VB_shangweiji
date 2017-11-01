VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "读入数据"
   ClientHeight    =   4695
   ClientLeft      =   13530
   ClientTop       =   3150
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox In_Text 
      Height          =   4035
      Left            =   300
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   180
      Width           =   3555
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Form_Load()
   Dim I As Single
'   In_Text.FontSize = 13
   In_Text.Text = "最大值" & Format(max, "#0.###") & "     " & "最小值" & Format(min, "#0.###") & vbCrLf
   
  
   In_Text.Text = In_Text.Text & "通道" & "   采样幅值" & "   采样频率" & vbCrLf
 
   For I = 0 To 200
    In_Text.Text = In_Text.Text & "  " & tad_stch & "    " & Format(data_value(I), "0.0000") & "   " & Format(frequent, "0.0000") & vbCrLf
   Next I
End Sub

