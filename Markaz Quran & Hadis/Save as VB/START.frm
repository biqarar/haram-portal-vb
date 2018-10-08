VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form START 
   BorderStyle     =   0  'None
   Caption         =   "„—ò“ ﬁ—¬‰ Ê ÕœÌÀ ò—Ì„Â «Â· »Ì  ⁄·ÌÂ« «·”·«„"
   ClientHeight    =   6585
   ClientLeft      =   8895
   ClientTop       =   1575
   ClientWidth     =   5445
   Icon            =   "START.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6585
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   6360
      TabIndex        =   11
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Motor"
      Height          =   3855
      Left            =   6600
      TabIndex        =   10
      Top             =   7920
      Visible         =   0   'False
      Width           =   2895
      Begin MSAdodcLib.Adodc Qeybat 
         Height          =   330
         Left            =   360
         Top             =   2520
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   $"START.frx":08CA
         OLEDBString     =   $"START.frx":0953
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from qeybat"
         Caption         =   "Qeybat"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc Emtahan 
         Height          =   330
         Left            =   360
         Top             =   2160
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   $"START.frx":09DC
         OLEDBString     =   $"START.frx":0A65
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from emtahan"
         Caption         =   "Emtahan"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc STU2CLASS 
         Height          =   375
         Left            =   360
         Top             =   1800
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   $"START.frx":0AEE
         OLEDBString     =   $"START.frx":0B77
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *  from stu2class"
         Caption         =   "STU2CLASS"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc mclass 
         Height          =   375
         Left            =   360
         Top             =   1440
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   $"START.frx":0C00
         OLEDBString     =   $"START.frx":0C89
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from mclass"
         Caption         =   "mclass"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc teacher 
         Height          =   375
         Left            =   360
         Top             =   1080
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   $"START.frx":0D12
         OLEDBString     =   $"START.frx":0D9B
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from teacher"
         Caption         =   "teacher"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc Student 
         Height          =   330
         Left            =   360
         Top             =   360
         Visible         =   0   'False
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   $"START.frx":0E24
         OLEDBString     =   $"START.frx":0EAD
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from student"
         Caption         =   "Student"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc Tarhha 
         Height          =   375
         Left            =   360
         Top             =   720
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   $"START.frx":0F36
         OLEDBString     =   $"START.frx":0FBF
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from tarhha"
         Caption         =   "Tarhha"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc Govahi 
         Height          =   330
         Left            =   360
         Top             =   2880
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   $"START.frx":1048
         OLEDBString     =   $"START.frx":10D1
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from govahi"
         Caption         =   "Govahi"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin ComctlLib.ProgressBar pb1 
      Height          =   135
      Left            =   0
      TabIndex        =   9
      Top             =   5520
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
      Max             =   6
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   7200
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1485
      Left            =   7440
      Picture         =   "START.frx":115A
      ScaleHeight     =   1455
      ScaleWidth      =   5280
      TabIndex        =   0
      Top             =   5400
      Visible         =   0   'False
      Width           =   5310
   End
   Begin VB.Image Image2 
      Height          =   6615
      Left            =   0
      Picture         =   "START.frx":3165
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   0
      Picture         =   "START.frx":189151
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   4095
   End
   Begin VB.Label LD 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00/00/00"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   315
      Left            =   3840
      TabIndex        =   8
      Top             =   4080
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RM.Biqarar@Gmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   4200
      Width           =   1995
   End
   Begin VB.Label L4 
      AutoSize        =   -1  'True
      Caption         =   "·ÿ›« ò„Ì ’»— ò‰Ìœ"
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   6
      Top             =   6480
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label L3 
      AutoSize        =   -1  'True
      Caption         =   "...œ— Õ«· »—ﬁ—«—Ì «— »«ÿ »« ”—Ê—"
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      TabIndex        =   5
      Top             =   6600
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Label L5 
      AutoSize        =   -1  'True
      Caption         =   "»—«Ì ‘—Ê⁄ ò·Ìò ò‰Ìœ"
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5880
      TabIndex        =   4
      Top             =   7440
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label L1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‰—„ «›“«— „œÌ—Ì  ¬„Ê“‘Ì"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   600
      TabIndex        =   3
      Top             =   7080
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Label L2 
      AutoSize        =   -1  'True
      Caption         =   "„—ò“ ﬁ—¬‰ Ê ÕœÌÀ ò—Ì„Â «Â· »Ì  ⁄·ÌÂ« «·”·«„"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      TabIndex        =   2
      Top             =   6000
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      BackStyle       =   0  'Transparent
      Caption         =   "»”„ «··Â «·—Õ„‰ «·—ÕÌ„"
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Top             =   0
      Width           =   1545
   End
End
Attribute VB_Name = "START"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SS As Integer


Private Sub Form_Load()
'LD.Caption = Taqvim.Label1.Caption

End Sub

Private Sub Form_Unload(Cancel As Integer)
user.Show

End Sub

Private Sub Timer1_Timer()
'SS = SS + 1

PB1.Value = 0
l1.Visible = True
l1.Caption = "‰—„ «›“«— „œÌ—Ì  ¬„Ê“‘Ì"





PB1.Value = 1
l2.Visible = True


l2.Caption = "„—ò“ ﬁ—¬‰ Ê ÕœÌÀ ò—Ì„Â «Â· »Ì  ⁄·ÌÂ« «·”·«„"


PB1.Value = 2
Me.Image1.Visible = True
PB1.Value = 3
L3.Visible = True
L3.Caption = "œ— Õ«· »— ﬁ—«—Ì «— »«ÿ »« ”—Ê—"

On Error GoTo 1
GoTo 2
1:
MsgBox "«„ò«‰ »—ﬁ—«—Ì «— »«ÿ »« ”—Ê— ÊÃÊœ ‰œ«—œ ·ÿ›« ’Õ  ‘»òÂ —« »——”Ì ò‰Ìœ", vbCritical + vbOKOnly, "Œÿ«"
End
2

user.Text3.Text = "RezaMohiti"
SettingF.Show

SettingF.Hide

 PB1.Value = 4
 L4.Visible = True
 
 L4.Caption = "·ÿ›« ò„Ì ’»— ò‰Ìœ"


 
 PB1.Value = 5
 
  L4.Caption = "·ÿ›« ò„Ì ’»— ò‰Ìœ"


 
 
 
 Student.Refresh
Qeybat.Refresh
mclass.Refresh
STU2CLASS.Refresh
Emtahan.Refresh
Govahi.Refresh





PB1.Value = 6
Unload Me




Exit Sub


If SS = 4 Then l1.Visible = True
If SS = 8 Then l2.Visible = True
If SS = 20 Then
L3.Visible = True

End If

If SS = 20 Then
L4.Visible = True




End If

If SS = 30 Then L5.Visible = True
If SS = 15 Then Picture1.Visible = True

If SS <= 30 Then PB1.Value = SS
If SS = 32 Then
Unload Me

End If

End Sub
