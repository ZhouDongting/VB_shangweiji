VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "AD��������"
   ClientHeight    =   9405
   ClientLeft      =   -20295
   ClientTop       =   24975
   ClientWidth     =   12330
   Icon            =   "Showhello.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9405
   ScaleWidth      =   12330
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   9780
      Top             =   6420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Showhello.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Showhello.frx":0B8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Showhello.frx":0DFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Showhello.frx":10BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Showhello.frx":13AE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      Left            =   9480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   28
      Top             =   6120
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8640
      Top             =   540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Caption         =   "��������"
      Height          =   6855
      Left            =   9480
      TabIndex        =   7
      Top             =   720
      Width           =   2655
      Begin VB.CommandButton Button_input 
         Caption         =   "��������"
         Height          =   435
         Left            =   1380
         TabIndex        =   31
         Top             =   3780
         Width           =   915
      End
      Begin VB.CommandButton Button_save 
         Caption         =   "��������"
         Height          =   435
         Left            =   1380
         TabIndex        =   30
         Top             =   3180
         Width           =   915
      End
      Begin VB.CommandButton BUtton_more 
         Caption         =   "����"
         Height          =   375
         Left            =   1800
         TabIndex        =   29
         Top             =   4680
         Width           =   615
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000007&
         Height          =   1035
         Left            =   60
         ScaleHeight     =   975
         ScaleWidth      =   2475
         TabIndex        =   20
         Top             =   180
         Width           =   2535
         Begin VB.Label Text4 
            BackColor       =   &H80000012&
            Caption         =   "0000.00"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   1140
            TabIndex        =   24
            Top             =   660
            Width           =   915
         End
         Begin VB.Label Text6 
            BackColor       =   &H80000012&
            Caption         =   "0000.00"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   315
            Left            =   1140
            TabIndex        =   23
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label7 
            BackColor       =   &H80000012&
            Caption         =   "���η�ֵ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   60
            TabIndex        =   22
            Top             =   660
            Width           =   1155
         End
         Begin VB.Label Label6 
            BackColor       =   &H80000012&
            Caption         =   "����Ƶ�ʣ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   60
            TabIndex        =   21
            Top             =   180
            Width           =   1155
         End
      End
      Begin VB.CommandButton Button_clear 
         Caption         =   "�������"
         Height          =   435
         Left            =   240
         TabIndex        =   18
         Top             =   3780
         Width           =   915
      End
      Begin VB.CommandButton Button_save1 
         Caption         =   "���沨��"
         Height          =   435
         Left            =   240
         TabIndex        =   17
         Top             =   3180
         Width           =   915
      End
      Begin VB.CommandButton Button_ad 
         Appearance      =   0  'Flat
         BackColor       =   &H80000008&
         Caption         =   "����AD"
         DownPicture     =   "Showhello.frx":1668
         Height          =   375
         Left            =   220
         MaskColor       =   &H00C0FFFF&
         TabIndex        =   16
         Top             =   4680
         Width           =   1155
      End
      Begin VB.ComboBox sample_fre1 
         Height          =   300
         ItemData        =   "Showhello.frx":1F32
         Left            =   1320
         List            =   "Showhello.frx":1F4E
         TabIndex        =   15
         Text            =   "400 KHZ"
         Top             =   1800
         Width           =   1155
      End
      Begin VB.ComboBox Gain 
         Height          =   300
         ItemData        =   "Showhello.frx":1F99
         Left            =   1320
         List            =   "Showhello.frx":1FA9
         TabIndex        =   13
         Text            =   "-10-10"
         Top             =   1380
         Width           =   1155
      End
      Begin VB.ComboBox End_Ad 
         Height          =   300
         ItemData        =   "Showhello.frx":1FDA
         Left            =   1320
         List            =   "Showhello.frx":200E
         TabIndex        =   9
         Text            =   "0"
         Top             =   2520
         Width           =   1155
      End
      Begin VB.ComboBox Start_Ad 
         Height          =   300
         ItemData        =   "Showhello.frx":2049
         Left            =   1320
         List            =   "Showhello.frx":207D
         TabIndex        =   8
         Text            =   "0"
         Top             =   2160
         Width           =   1155
      End
      Begin VB.Label Light 
         BackColor       =   &H80000008&
         Height          =   375
         Left            =   720
         TabIndex        =   25
         Top             =   4680
         Width           =   795
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000A&
         Caption         =   "����Ƶ��"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1860
         Width           =   795
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000A&
         Caption         =   "���뷶Χ"
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000A&
         Caption         =   "����ͨ��"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2580
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000A&
         Caption         =   "��ʼͨ��"
         Height          =   315
         Left            =   240
         TabIndex        =   10
         Top             =   2220
         Width           =   795
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      DragIcon        =   "Showhello.frx":20B8
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   9105
      Width           =   12330
      _ExtentX        =   21749
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   5265
            Picture         =   "Showhello.frx":2982
            Text            =   "AD��������"
            TextSave        =   "AD��������"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   5265
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   5265
            TextSave        =   "2017/5/18"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   5265
            TextSave        =   "22:43"
            Object.ToolTipText     =   "����ʱ��"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7440
      Top             =   540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Showhello.frx":325C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Showhello.frx":351E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Showhello.frx":4DB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Showhello.frx":5026
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Showhello.frx":52E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Showhello.frx":55DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Showhello.frx":5894
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12330
      _ExtentX        =   21749
      _ExtentY        =   767
      ButtonWidth     =   1323
      ButtonHeight    =   609
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��ӡ"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�Ŵ�"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��С"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��С"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��λ"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�˳�"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6480
      Top             =   780
   End
   Begin VB.Frame Frame 
      BackColor       =   &H8000000A&
      Height          =   8235
      Left            =   240
      TabIndex        =   2
      Top             =   540
      Width           =   9015
      Begin VB.Frame FFT_Frame 
         BackColor       =   &H8000000A&
         Caption         =   "FFT"
         ForeColor       =   &H00FF0000&
         Height          =   3615
         Left            =   180
         TabIndex        =   5
         Top             =   4440
         Width           =   8715
         Begin VB.TextBox Text3 
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            ForeColor       =   &H0000FFFF&
            Height          =   270
            Left            =   60
            TabIndex        =   33
            Text            =   "V/S"
            Top             =   360
            Width           =   315
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00000000&
            FillColor       =   &H0000FFFF&
            ForeColor       =   &H0000FFFF&
            Height          =   3135
            Left            =   360
            ScaleHeight     =   317.829
            ScaleMode       =   0  'User
            ScaleWidth      =   239.151
            TabIndex        =   6
            Top             =   360
            Width           =   8115
            Begin VB.Label Label8 
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               Caption         =   "KHZ"
               ForeColor       =   &H0000FFFF&
               Height          =   315
               Left            =   7680
               TabIndex        =   34
               Top             =   2760
               Width           =   495
            End
            Begin VB.Line x_line2 
               BorderColor     =   &H0000FF00&
               Visible         =   0   'False
               X1              =   0
               X2              =   240.487
               Y1              =   117.829
               Y2              =   117.829
            End
            Begin VB.Line y_line2 
               BorderColor     =   &H0000FF00&
               Visible         =   0   'False
               X1              =   106.883
               X2              =   106.883
               Y1              =   0
               Y2              =   316.279
            End
            Begin VB.Label Cordinate3 
               AutoSize        =   -1  'True
               Caption         =   "����3"
               ForeColor       =   &H0000FFFF&
               Height          =   180
               Left            =   660
               TabIndex        =   26
               Top             =   240
               Visible         =   0   'False
               Width           =   450
            End
         End
      End
      Begin VB.Frame AD_Frame 
         BackColor       =   &H8000000A&
         Caption         =   "AD_Frame"
         DragIcon        =   "Showhello.frx":36C54
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   4035
         Left            =   180
         TabIndex        =   3
         Top             =   240
         Width           =   8595
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            ForeColor       =   &H0000FFFF&
            Height          =   270
            Left            =   160
            TabIndex        =   32
            Text            =   "V"
            Top             =   360
            Width           =   120
         End
         Begin VB.PictureBox Picture1 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            FillColor       =   &H0080FF80&
            ForeColor       =   &H0000FFFF&
            Height          =   3555
            Left            =   300
            ScaleHeight     =   -11000
            ScaleMode       =   0  'User
            ScaleTop        =   5500
            ScaleWidth      =   2000
            TabIndex        =   4
            Top             =   300
            Width           =   8115
            Begin VB.Label Cordinate1 
               AutoSize        =   -1  'True
               Caption         =   "����1"
               ForeColor       =   &H0000FFFF&
               Height          =   480
               Left            =   780
               TabIndex        =   19
               Top             =   300
               Visible         =   0   'False
               Width           =   690
            End
            Begin VB.Line y_line 
               BorderColor     =   &H0000FF00&
               Visible         =   0   'False
               X1              =   908.752
               X2              =   908.752
               Y1              =   5500
               Y2              =   -5452.79
            End
            Begin VB.Line x_line 
               BorderColor     =   &H0000FF00&
               Visible         =   0   'False
               X1              =   0
               X2              =   2011.173
               Y1              =   1912.017
               Y2              =   1912.017
            End
         End
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label7"
      ForeColor       =   &H0000FF00&
      Height          =   180
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Imprint 
      Height          =   210
      Left            =   7980
      Picture         =   "Showhello.frx":3751E
      Top             =   4920
      Width           =   240
   End
   Begin VB.Image Imsave 
      Height          =   210
      Left            =   7920
      Picture         =   "Showhello.frx":37800
      Top             =   5700
      Width           =   210
   End
   Begin VB.Image Imopen 
      Height          =   195
      Left            =   8040
      Picture         =   "Showhello.frx":37AAA
      Top             =   5340
      Width           =   225
   End
   Begin VB.Image Imexit 
      Height          =   225
      Left            =   7980
      Picture         =   "Showhello.frx":37D5C
      Top             =   4380
      Width           =   180
   End
   Begin VB.Image Imcopy 
      Height          =   195
      Left            =   7860
      Picture         =   "Showhello.frx":37FBA
      Top             =   3960
      Width           =   225
   End
   Begin VB.Menu Menu_Open 
      Caption         =   "�ļ�&O"
      WindowList      =   -1  'True
      Begin VB.Menu MenuFileOpen 
         Caption         =   "��"
         Shortcut        =   ^O
      End
      Begin VB.Menu MenuFileNew 
         Caption         =   "�½�"
         Shortcut        =   ^N
      End
      Begin VB.Menu MenuFileSave 
         Caption         =   "����"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuFilePrint 
         Caption         =   "��ӡ"
      End
      Begin VB.Menu MenuFileExit 
         Caption         =   "�ر�"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu Menu_Edit 
      Caption         =   "�༭&E"
      Begin VB.Menu Menu_jian 
         Caption         =   "����"
      End
      Begin VB.Menu Menu_Copy 
         Caption         =   "����"
         Shortcut        =   ^C
      End
      Begin VB.Menu Menu_Post 
         Caption         =   "ճ��"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu Menu_Show 
      Caption         =   "��ͼ&V"
      Begin VB.Menu Menu_Show2 
         Caption         =   "��׼���"
      End
      Begin VB.Menu Menu_Show3 
         Caption         =   "�������"
      End
   End
   Begin VB.Menu Menu_Black 
      Caption         =   "����&B"
      Begin VB.Menu Menu_BlackBlack 
         Caption         =   "��ɫ"
      End
      Begin VB.Menu Menu_BlackWhite 
         Caption         =   "��ɫ"
      End
   End
   Begin VB.Menu Menu_Star 
      Caption         =   "����&S"
   End
   Begin VB.Menu Menu_Help 
      Caption         =   "����&H"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adstate As Long
'********** ȫ�ֱ���
Public s As Single
Public s_1 As Single
Public zoom As Double   '�Ŵ���С���� ��ʼ��Ϊ5
Public TempFile$
Public Skin As String 'Ƥ��·��
Dim I As Single
Dim Waveform_Color As Single
Dim Line_Color As Single
Dim Coordinate_Color As Single
Dim Back_Color As Single
Dim a As Single

'*******ȡ�ض����˵����Ĳ������
Private Declare Function GetMenu Lib "user32" _
   (ByVal hWnd As Long) As Long

Private Declare Function GetSubMenu Lib "user32" _
   (ByVal hMenu As Long, ByVal nPos As Long) As Long

Private Declare Function SetMenuItemBitmaps Lib "user32" _
   (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, _
    ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
'-------------------------------------------------

Const MF_BYPOSITION = &H400&
'********************AD��ʾ�����ʼ��************************
Private Sub Interface_Init()
 Dim a$, b$
 
 'picture1�����������ʼ�����߶ȵ�λ22�����1000�����Ͻ����꣨0,11��
 With Picture1
    .BackColor = vbBlack
    .ScaleTop = 11
    .ScaleHeight = -22
    .ScaleLeft = 0
    .ScaleWidth = 1024
 End With
    Line_Color = RGB(0, 100, 0)
    Coordinate_Color = vbYellow
    Waveform_Color = vbGreen
    
    tad_maxlen = 2000000
    hDevice = MP422E_OpenDevice(0) '�õ�AD�忨�������
'*******hDevice�ж��Ƿ�������AD�忨
    If hDevice = &HFFFFFFFF Then
      Form1.Caption = "Error Load AC6623"
      Beep
    End If
    If hDevice <> &HFFFFFFFF Then
 
    End If
 'AD�ɼ�������ʼ��
    tad_sidi = 0

    sample_fre = 400
    
    tad_stch = 0
    tad_endch = 0 'total 8ch
    tad_gain = 0 ' -10-+10V
    tad_maxlen = 2000000  ' max sam length 2M ad
    tad_total = 0 ' total length=0

  'show_location1����2 ����2����ʾ������������ӳ���
     zoom = 1
     show_location (zoom)
     
     Picture2.Scale (0, 10)-(512, -10)
     show_location2
     
'***********************�˵���ͼ�����API**********************
    Dim mHandle As Long, lRet As Long, sHandle As Long, sHandle2 As Long
    'ȡ�ò˵��ľ������ֵ��mHandle
    mHandle = GetMenu(hWnd)
    'ȡ��mHandle�����ָ�˵��ĵ�һ������ʽ�˵����ļ�&F���ľ������ֵ��sHandle
    sHandle = GetSubMenu(mHandle, 0)
 
    lRet = SetMenuItemBitmaps(sHandle, 0, MF_BYPOSITION, Imopen.Picture, Imsave.Picture)
    lRet = SetMenuItemBitmaps(sHandle, 1, MF_BYPOSITION, Imsave.Picture, Imsave.Picture)
    lRet = SetMenuItemBitmaps(sHandle, 2, MF_BYPOSITION, Imprint.Picture, Imprint.Picture)
    lRet = SetMenuItemBitmaps(sHandle, 4, MF_BYPOSITION, Imprint.Picture, Imprint.Picture)
    lRet = SetMenuItemBitmaps(sHandle, 5, MF_BYPOSITION, Imexit.Picture, Imexit.Picture)
       
    'ȡ��mHandle�����ָ�˵��ĵڶ�������ʽ�˵����༭&E���ľ������ֵ��sHandle
    sHandle = GetSubMenu(mHandle, 1)
    lRet = SetMenuItemBitmaps(sHandle, 0, MF_BYPOSITION, Imopen.Picture, Imsave.Picture)
    lRet = SetMenuItemBitmaps(sHandle, 1, MF_BYPOSITION, Imsave.Picture, Imsave.Picture)
    lRet = SetMenuItemBitmaps(sHandle, 2, MF_BYPOSITION, Imprint.Picture, Imprint.Picture)
    
End Sub

' ���꽨���ӳ���
Private Sub show_location(ByVal step_x As Integer)
 Dim I As Integer
    Cordinate1.BackStyle = 0
    Dim b As Single
    '����picture���ػ棬�߿�1�����2
    Picture1.AutoRedraw = True
    Picture1.DrawStyle = 2
    Picture1.DrawWidth = 1
    
    Picture1.BackColor = Back_Color
    Picture1.ForeColor = Line_Color
    'ѭ����������
    For I = 0 To 6000 Step step_x * 50
    b = b + 1
    Picture1.DrawStyle = 2
    If b = 5 Then
    Picture1.DrawStyle = 0
    b = 0
    End If
    Picture1.Line (I, -9000)-(I, 9000)
    Next I
    
    For I = -20 To 20 Step step_x * 2
    Picture1.Line (-8000, I)-(12000, I)
    Next I
    
    Picture1.ForeColor = RGB(250, 0, 0)
    Picture1.DrawStyle = 0
    Picture1.Line (-13000, 0)-(13000, 0)
    
    a = 0
    Picture1.ForeColor = Coordinate_Color
    'ѭ��д������
    For I = 0 To 10 * (step_x) Step step_x * 2
        Picture1.CurrentX = 2
        Picture1.CurrentY = -I + 0.1
        Picture1.Print Format(-a)
        Picture1.CurrentY = I + 0.1
        Picture1.Print Format(a)
        a = a + 2
    Next I
    a = 0
    'ѭ��д������
    For I = 0 To 5000 Step step_x * 50
        Picture1.CurrentX = I
        Picture1.CurrentY = 0
        a = I / sample_fre / step_x
        If (I <> 0) Then Picture1.Print Format(a, "0.0")
    Next I
        
End Sub


'�Ѵ��ļ��е�����װ����������
Private Sub Read_data()
  CommonDialog1.ShowOpen       '���ô��ļ��Ի���
      If CommonDialog1.FileName = "" Then Exit Sub '' �ж��ļ����Ƿ�Ϊ�գ��վ������˺���
  '       MsgBox "û��ѡ���ļ�", , " "
       
         TempFile = CommonDialog1.FileName      '���ļ���
         Dim a$
        Dim I As Integer
        
  Open TempFile For Input As #1    ' ���ļ���EOF�жϲ�Ϊ��ʱ�����ļ�����һֱ����data_value������
  
    Do While Not EOF(1)
       I = I + 1
       Line Input #1, a
       If (I > 3) Then
       data_value(I - 4) = a
       End If
    Loop
       Close #1        '�ر��ļ�
       s = 1              ' ��־λ����־���������ļ�������ҪAD�ɼ�
     
      Text1_Show
      Timer1.Enabled = True
    
End Sub

' BUTTON ����AD �¼�
Private Sub Button_ad_Click()

If Button_ad.Caption = "����AD" Then
    Button_ad.Caption = "��ͣAD"
    Form1.Cls
    MP422E_CAL hDevice           'ADУ׼����
     MP422E_AD hDevice, tad_stch, tad_endch, tad_gain, 0, 0, 0, 0, 0, 0, 10000 / sample_fre    '����AD�ɼ�
    s = 0        '��־λ  ��ҪAD�ɼ�����
    s_1 = 1
    Timer1.Enabled = True        ' ��ʱ��������100ms�ɼ�һ������
    Light.BackColor = vbRed
  
ElseIf Button_ad.Caption = "��ͣAD" Then   '�ٴε���˰�ť�����ر�AD��ʧ�ܶ�ʱ��
    Button_ad.Caption = "����AD"
    Light.BackColor = vbBlack
     MP422E_StopAD hDevice
    Timer1.Enabled = False
End If

End Sub

' ������ΰ�ť�¼�
Private Sub Button_clear_Click()
  zoom = 1
  show_location (zoom) '�����Ժ��ػ���������
  show_location2
  Timer1.Enabled = False
   MP422E_StopAD hDevice
  Erase data_value      '��ʽ�� data_value
End Sub
'****�������ݰ�ť�¼�
Private Sub Button_input_Click()
   Read_data
End Sub

'******'��ťButton_Save���������¼�****************'
Private Sub Button_save_Click()
error:
    With CommonDialog1           '����ConnonDiaolog1���ڣ������ļ���Դ������
         .FileName = ""
         .Filter = "�ĵ��ļ�(*.txt)|*.txt"
         .ShowSave
    End With
    If (CommonDialog1.FileName = "") Then Exit Sub '' �ж��ļ����Ƿ�Ϊ�գ��վ������˳���
'       MsgBox "�ļ���Ϊ��", vbExclamation
'       GoTo error
'    End If
    Open CommonDialog1.FileName For Output As #1   '' ���ļ�
    Print #1, "����Ƶ��" & "  " & frequent & " KHz   " & "��ֵ" & " " & Format((-min), "00.0000") & " V"
    Print #1, " "
    Print #1, "��������Ϊ"
    For I = 0 To 5000           '��ȡ5000�����ݣ�������datavalue
      Print #1, Format(data_value(I), "00.00000")
    Next I

    Close #1
End Sub
'******���沨��ͼ��
Private Sub Button_save1_Click()
   Dim b As Single
     With CommonDialog1           '����ConnonDiaolog1���ڣ������ļ���Դ������
         .FileName = ""
         .Filter = "ͼ���ļ�(*.bmp,*.jpg)|*.bmp;*.jpg"
         .ShowSave
    End With
    If (CommonDialog1.FileName = "") Then Exit Sub '' �ж��ļ����Ƿ�Ϊ�գ��վ������˳���
    SavePicture Picture1.Image, CommonDialog1.FileName
      With CommonDialog1           '����ConnonDiaolog1���ڣ������ļ���Դ������
         .FileName = ""
         .Filter = "ͼ���ļ�(*.bmp,*.jpg)|*.bmp;*.jpg"
         .ShowSave
    End With
    If (CommonDialog1.FileName = "") Then Exit Sub '' �ж��ļ����Ƿ�Ϊ�գ��վ������˳���
    SavePicture Picture2.Image, CommonDialog1.FileName
End Sub

'******'��������¼�****************'
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

      'cordinate1Ϊ��ǩ��captionΪ������ʾ��Ϣ���˴���������ʾ��Ϣ����Ϊ��ǰ����λ��
       Cordinate1.BackStyle = 0
       Cordinate1.Visible = True
       Cordinate1.ForeColor = Coordinate_Color
       Cordinate1.Caption = "X��" & Format((x / zoom) / sample_fre, "0.00") & "ms" & vbCrLf & "Y��" & Format(y / zoom, "0.0") & "V"
       
 '****���ݴ�ʱ������ڵ�X��Yֵ����̬���������ǩ��λ��*********'
 '***Picture1.ScaleTopΪ��ʾ�ؼ��ĸ߶ȣ�widthΪ���
    If (y < Picture1.ScaleTop * 0.23 And x > Picture1.ScaleWidth * 0.8) Then
       Cordinate1.left = x - Picture1.ScaleWidth * 0.1
       Cordinate1.top = y + 3
    ElseIf (y < Picture1.ScaleTop * 0.23) Then
       Cordinate1.left = x + Picture1.ScaleWidth * 0.03
       Cordinate1.top = y + 3
    ElseIf (x > Picture1.ScaleWidth * 0.8) Then
       Cordinate1.left = x - Picture1.ScaleWidth * 0.1
       Cordinate1.top = y - 1
    Else
    Cordinate1.left = x + Picture1.ScaleWidth * 0.03
    Cordinate1.top = y - 1
    End If
       'ʮ�ֶ�λ�߿ɼ�������������λ��Ϊ������ڵ�
    x_line.BorderColor = Waveform_Color
    y_line.BorderColor = Waveform_Color
        x_line.Visible = True
        y_line.Visible = True

        x_line.Y1 = y: x_line.Y2 = y
        y_line.X1 = x: y_line.X2 = x

End Sub

'******'��������¼�****************'
Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      Cordinate3.BackStyle = 0
        Cordinate3.Visible = True
        Cordinate3.ForeColor = Coordinate_Color
        Cordinate3.Caption = "X��" & Format(200 / (12.775 * 400 / sample_fre * 400 / sample_fre) * (x + 6.3875 * 400 / sample_fre), "#0.0") & "Khz"

 '****����X��Yֵ�����������ǩ��λ��*********'
    If (x > 4.5) Then
       Cordinate3.top = y
       Cordinate3.left = x - 2
    Else
    Cordinate3.top = y
    Cordinate3.left = x + 0.5
    End If
    x_line2.BorderColor = Waveform_Color
    y_line2.BorderColor = Waveform_Color
    x_line2.Visible = True
    y_line2.Visible = True

    x_line2.Y1 = y: x_line2.Y2 = y
    y_line2.X1 = x: y_line2.X2 = x
End Sub


'******'�������Ƴ�������ʾ���ڣ���������ʾ��ǩ����****************'
Private Sub AD_Frame_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Cordinate1.Visible = False
x_line.Visible = False
y_line.Visible = False
End Sub

'******'�������Ƴ�������ʾ���ڣ���������ʾ��ǩ����****************'
Private Sub FFT_Frame_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Cordinate3.Visible = False
x_line2.Visible = False
y_line2.Visible = False
End Sub

'********��ʱ���¼�****************'
Private Sub Timer1_Timer()
        Dim ReturnData As Long          'return AD data length
        Dim tmp As String
        Dim L As Long
        Dim sw As Single
        
  's=0��־λ��˵����ͼǰ��Ҫ�ɼ�����
If s = 0 Then
        ReturnData = MP422E_Poll(hDevice)  '����AD�ɼ����ݳ���
        If ReturnData < 0 Then ' С��0  ˵������ֹͣ�ɼ�
           MP422E_StopAD hDevice
        '   Timer1.Enabled = False
           MsgBox$ ("over!")
           GoTo tend1
        End If

        If ReturnData < 1000 Then 'С��1000�����ݹ���
           GoTo tend1
        End If
           'read data lengt=N * AD Channel number
           '��������������������
           'tad_endch - tad_stch + 1 ����ͨ������
        L = (ReturnData \ (tad_endch - tad_stch + 1)) * (tad_endch - tad_stch + 1)
           
        MP422E_Read hDevice, L, tad_data(tad_total) '���زɼ������ݣ�������tad_data������
'         Dim j As Single
'         For j = 0 To 100
'           For I = 0 To ((tad_endch - tad_stch + 1) - 1)
'              ' Form1.Print "CH"; i + tad_stch, "Data"; tad_data(tad_total + i), "Vol   "; Format(MP422E_ADV(tad_gain, tad_data(tad_total + i)), "0.000")
'                data_1(I) = MP422E_ADV(tad_gain, tad_data(I)) / 1000
'           Next I
'           I = 0
'           'tad_total = tad_total + L   ' added length
'           Dim j As Single
'           For I = 0 To 1000
'             For j = 0 To ((tad_endch - tad_stch + 1) - 1)
'                  Picture1.Line (I * zoom, data_1(I + j) * zoom)-(((I + 1) * zoom), data_1(I + j + 1)), vbGreen
'            Next j
'           Next I
'         Next j

          For I = 0 To 8000       'ȡ8000�����ݸ�datavalue�����滭ͼ�ã�mv����1000���õ�v
             data_value(I) = MP422E_ADV(tad_gain, tad_data(I)) / 1000
          Next I
        If (s_1 = 1) Then
           Text1_Show
           s_1 = 0
         End If
    
tend1:

  End If
          huaboxing   '���û�����
          data_analysis   '�������Сֵ����ֵ��Ƶ��
          huaFFT       '��FFT
      
End Sub
Public Sub huaboxing()
      Dim p As Long
       Picture1.Cls   '�������ػ���������
       show_location (zoom)
       
        Picture1.ForeColor = Waveform_Color             'ѭ����3000���㣬LINE��������x1,y1��-(x2,y2)
        For I = 0 To 3000
           Picture1.Line (I * zoom, data_value(I) * zoom)-(((I + 1) * zoom), data_value(I + 1) * zoom)
         Next I
End Sub

'���ƿ���FFT�任���Ƶ����
Private Sub huaFFT()

    show_location2
    Dim p As Double, q As Integer
    Dim Height As Double, space As Double
    Dim t As Double

    t = 1 / sample_fre / 2 * (sample_length - 1)
    sample_length = 512
    For p = 0 To sample_length - 1 '���ڱ���FFT���ݵ�����
        data_fft(p) = data_value(p)
    Next p
    
    Dim wr As Double, wi As Double, max1 As Double
    wr = Cos(pi / sample_length)
    wi = Sin(pi / sample_length)
    Call Rdft(sample_length, wr, wi, data_fft)       '����FFTת������
    
    'ȡ��fft��������ֵ
    max1 = data_fft(0)
    For q = 0 To sample_length - 1
        If (data_fft(q) > max1) Then max1 = data_fft(q)
    Next q
'    Text3.Text = Format(max1, "##.####")
    Picture2.ScaleHeight = 1.5 * max1 '��֤������ͼ���ᳬ��ͼ����
    space = 1 / sample_fre
    q = 0
    
    Picture2.Scale (-t * 10, 1.5 * max1)-(t * 10, -(1.5 * max1 / 4))
    Picture2.ForeColor = Waveform_Color
    For p = -t To t - space Step space
        Picture2.Line (p * 10, Abs(data_fft(q)))-((p + space) * 10, Abs(data_fft(q + 1))) '��Ƶ����
        q = q + 1
    Next p
End Sub

' ���꽨���ӳ���
Private Sub show_location2()

Picture2.Cls
Picture2.Scale (0, 10)-(1000, -10)
Picture2.AutoRedraw = True

Picture2.BackColor = Back_Color
Picture2.ForeColor = Line_Color
Picture2.DrawWidth = 1
Picture2.DrawStyle = 2

  For I = -8 To 9
    Picture2.ForeColor = Line_Color
    Picture2.Line (0, 10 - 2 * I)-(1023, 10 - 2 * I)
    Picture2.CurrentX = -11
    Picture2.CurrentY = -2 * I
    Picture2.ForeColor = Coordinate_Color
    Picture2.Print 6 - 2 * I
  Next I

 'ѭ��д������
  a = 0
  For I = 0 To 1000 Step 50 * 400 / sample_fre
        Picture2.ForeColor = Line_Color
        Picture2.Line (I + 1, -10)-(I + 1, 10)
        Picture2.CurrentX = I
        Picture2.CurrentY = -6
        Picture2.ForeColor = Coordinate_Color
        Picture2.Print Format(a)
        a = a + 10
    Next I

  Picture2.DrawStyle = 0
  Picture2.Line (0, -6)-(1023, -6), vbRed
End Sub

'���ݷ���
Private Sub data_analysis()

Dim I, sum As Single

min = max = sum = data_value(0)
                           '����Сֵ
For I = 0 To 2000
    If data_value(I) < min Then
    min = data_value(I)
    End If
Next I
    min = min + 0.1                  '�����ֵ

For I = 0 To 1000
  sum = sum + Abs(data_value(I))
    If data_value(I) > max Then
    max = data_value(I)
    End If
Next I
                                '���ֵ
mean = sum / 1000
                                '������
Dim t1, t2, t3 As Integer
Dim period As Double
Dim interv As Double
Dim frequency As Double
Dim j As Single

period = 1 / sample_fre   'period ��������֮���ʱ����
'�ж���������֮��Ĳ����������������

    For I = 0 To 4095
        If data_value(I + j * 512) * data_value(I + j * 512 + 1) <= 0 And data_value(I + j * 512) < 0 And t1 = 0 Then '����ֵ����������       ?
        t2 = I
        t1 = t1 + 1
        End If
        If data_value(I + j * 512) * data_value(I + j * 512 + 1) <= 0 And data_value(I + j * 512) > 0 And t1 = 1 Then '����ֵ����������       ?
        t3 = I
        t1 = t1 - 1
        End If
    Next I
     interv = 2 * Abs(t3 - t2) * period '����*���=����
     
'Text5.Text = Abs(t3 - t2)
                                                                                '?
' interv = 2 * Abs(t3 - t2) * period  '����

If interv <> 0 Then frequency = 1 / interv 'Ƶ��
    
    frequent = frequency

    If Abs(mean) < 1 Then
    Text4.Caption = Round(Abs(min), 2) & "V"
    Else
    Text4.Caption = Round(Abs(min), 2) & "V"
    End If

    If Abs(frequency) < 1 Then
    Text6.Caption = Format(frequency, "#0.000") & "kHz"
    Else
    Text6.Caption = Format(frequency, "#0.000") & "kHz"
    End If
    frequent = frequency

End Sub


'********��������¼��������Ժ����еĵ�һ������****************'
Private Sub Form_Load()
  Interface_Init
  show_location2
 ' SkinH_Attach             '��ȡƤ���ļ���λ��
 ' SkinH_AttachEx GetAppPath + "Skin\" & "itunes.she", ""   'Ӧ��Ƥ��
End Sub

Public Function GetAppPath() '��ȡ����·������׼��ʽ
On Error Resume Next
GetAppPath = Replace(App.Path + "\", "\\", "\")
End Function

'******��ʾ���������Ϣ
Private Sub Text1_Show()
    Dim I As Single
    ' Text1.FontSize = 13
     Text1.Text = "ͨ��" & "   ������ֵ" & vbCrLf

     For I = 0 To 200
     Text1.Text = Text1.Text & "  " & tad_stch & "    " & Format(data_value(I), "0.0000") & vbCrLf
     Next I
End Sub

Private Sub End_Ad_Click()


 If Button_ad.Caption = "�ر�AD" Then
    MsgBox "����������ʱ���Ĳ���", vbExclamation
    Exit Sub
  ElseIf Button_ad.Caption = "����AD" Then
    tad_endch = Val(End_Ad.Text)
  End If
End Sub

Private Sub Gain_Click()
Dim c$
 If Button_ad.Caption = "�ر�AD" Then
    MsgBox "����������ʱ���Ĳ���", vbExclamation
    Exit Sub
  ElseIf Button_ad.Caption = "����AD" Then
  If (Gain.Text = "-10 - 10") Then tad_gain = 0
  If (Gain.Text = "-5  - 5") Then tad_gain = 1
  If (Gain.Text = "-2.5 - 2.5") Then tad_gain = 2
  If (Gain.Text = "-1.25 - 1.25") Then tad_gain = 3

   End If

  
End Sub


Private Sub Start_Ad_Click()
 If Button_ad.Caption = "�ر�AD" Then
    MsgBox "����������ʱ���Ĳ���", vbExclamation
    Exit Sub
  ElseIf Button_ad.Caption = "����AD" Then
    tad_stch = Val(Start_Ad.Text)
 '   Text1_Show
  End If
End Sub

Private Sub sample_fre1_Click()
 If Button_ad.Caption = "�ر�AD" Then
    MsgBox "����������ʱ���Ĳ���", vbExclamation
    Exit Sub
  ElseIf Button_ad.Caption = "����AD" Then
  If (sample_fre1.Text = "50 KHZ") Then sample_fre = 50
  If (sample_fre1.Text = "100 KHZ") Then sample_fre = 100
  If (sample_fre1.Text = "150 KHZ") Then sample_fre = 150
  If (sample_fre1.Text = "200 KHZ") Then sample_fre = 200
  If (sample_fre1.Text = "250 KHZ") Then sample_fre = 250
  If (sample_fre1.Text = "300 KHZ") Then sample_fre = 300
  If (sample_fre1.Text = "350 KHZ") Then sample_fre = 350
  If (sample_fre1.Text = "400 KHZ") Then sample_fre = 400
  show_location2
'    sample_fre = Val(sample_fre.Text)
'    show_location (zoom)
  '  Text1_Show
  End If
End Sub


Private Sub BUtton_more_Click()
   Dialog.Show           '��������ʾ����
End Sub


'******** TOOLBAR �������¼���case�жϵ���������ڼ���λ��****************'
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Index
    Case 1    '�򿪣����ô��ӳ�
        Read_data
    Case 2   '���棬���ñ����ӳ�
        Button_save_Click
    Case 3    '��ӡ
            
    Case 4    '�ָ���
    
    Case 5   '����
    Case 6   '�Ŵ󣬷Ŵ���zoom+1
     If (zoom < 10) Then zoom = zoom + 1
        Call show_location(zoom)
        huaboxing
    Case 7    '��С��zoom-1���ж��Ƿ�Ϊ1��Ϊ1ʱΪ��С
       If (zoom = 1) Then MsgBox "���������С����", vbExclamation
        If (zoom > 1) Then zoom = zoom - 1
        show_location (zoom)
        huaboxing
    Case 8     '���ƣ���ͼ�οؼ��ϲ������1��������������Ϊ
        Picture1.ScaleTop = Picture1.ScaleTop - 1
        show_location (zoom)
        huaboxing
    Case 9     '����
        Picture1.ScaleTop = Picture1.ScaleTop + 1
        show_location (zoom)
        huaboxing
    Case 10    '������ͼ�οؼ�x��ȼ�200��������ͼ�α��
    If (Picture1.ScaleWidth = 200) Then MsgBox "������������", vbExclamation
        If (Picture1.ScaleWidth > 200) Then Picture1.ScaleWidth = Picture1.ScaleWidth - 200
        show_location (zoom)
        huaboxing
    Case 11     '��С
         Picture1.ScaleWidth = Picture1.ScaleWidth + 200
         show_location (zoom)
         huaboxing
    Case 12    '��λ��ť��ͼ�οؼ������ʼ�����Ŵ�����λ
        With Picture1
           .BackColor = vbBlack
           .ScaleTop = 11
           .ScaleHeight = -22
           .ScaleLeft = 0
           .ScaleWidth = 1000
        End With
        zoom = 1
        Picture1.Cls
        show_location (zoom)
        huaboxing
    Case 14    '���ڣ����ö�������
       frmAbout.Show
    Case 15     '�˳�
       If MsgBox("��������˳��� ", vbYesNo + vbDefaultButton1, "�˳���ʾ ") = vbYes Then
        End
       Else
        
    End If
  End Select
End Sub

'****************�˵����¼�**************************

'********����Ҽ��¼�****************'
Private Sub Form_MouseUP(Button As Integer, Shift As Integer, x As Single, y As Single)
'BUTTON 1 Ϊ���
If Button = 2 Then
   PopupMenu Menu_Open
End If
End Sub

'********�򿪰�ť�˵����¼�
Private Sub MenuFileOpen_Click()
        Read_data
End Sub

'********���水ť�˵����¼�
Private Sub MenuFileSave_Click()
   Button_save_Click
End Sub
Private Sub Menu_BlackBlack_Click()
    Back_Color = vbBlack
    Line_Color = RGB(0, 100, 0)
    Coordinate_Color = vbYellow
    Waveform_Color = vbGreen
    show_location (zoom)
    show_location2
End Sub

Private Sub Menu_BlackWhite_Click()
    Back_Color = RGB(232, 232, 232)
    Line_Color = RGB(10, 10, 10)
    Coordinate_Color = vbRed
    Waveform_Color = RGB(10, 4, 200)
    show_location (zoom)
    show_location2
End Sub


Private Sub Menu_Show2_Click()
 SkinH_AttachEx GetAppPath + "Skin\" & "asus.she ", "" 'Ӧ��Ƥ����Ӧ�ó���"
End Sub
Private Sub Menu_Show3_Click()
 SkinH_AttachEx GetAppPath + "Skin\" & "itunes.she", "" 'Ӧ��Ƥ����Ӧ�ó���
End Sub
'*****�ر��¼���ť
Private Sub MenuFileExit_Click()
    If MsgBox("��������˳��� ", vbYesNo + vbDefaultButton1, "�˳���ʾ ") = vbYes Then
     End
    Else
    End If
End Sub
