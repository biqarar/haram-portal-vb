VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form QeybatFilter 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "„œÌ—Ì  Õ÷Ê— Ê €Ì«»"
   ClientHeight    =   1965
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4740
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "QeybatFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   1200
      Width           =   4335
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   315
         Left            =   2760
         TabIndex        =   13
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   135
      End
   End
   Begin VB.CommandButton Command14 
      Caption         =   " «ÌÌœ"
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3360
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command15 
      Caption         =   "«‰’—«›"
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3360
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
   Begin VB.Frame Frame9 
      Caption         =   "Motor"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   5160
      Visible         =   0   'False
      Width           =   1455
      Begin MSAdodcLib.Adodc vadie 
         Height          =   330
         Left            =   360
         Top             =   3360
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
         Connect         =   $"QeybatFilter.frx":08CA
         OLEDBString     =   $"QeybatFilter.frx":0953
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from vadie"
         Caption         =   "vadie"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Homa"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc Qeybat 
         Height          =   330
         Left            =   360
         Top             =   3000
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
         Connect         =   $"QeybatFilter.frx":09DC
         OLEDBString     =   $"QeybatFilter.frx":0A65
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from qeybat"
         Caption         =   "Qeybat"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
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
         Top             =   2160
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
         Connect         =   $"QeybatFilter.frx":0AEE
         OLEDBString     =   $"QeybatFilter.frx":0B77
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *  from stu2class"
         Caption         =   "STU2CLASS"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
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
         Top             =   1680
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
         Connect         =   $"QeybatFilter.frx":0C00
         OLEDBString     =   $"QeybatFilter.frx":0C89
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from mclass"
         Caption         =   "mclass"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
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
         Top             =   1200
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
         Connect         =   $"QeybatFilter.frx":0D12
         OLEDBString     =   $"QeybatFilter.frx":0D9B
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from teacher"
         Caption         =   "teacher"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
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
         Connect         =   $"QeybatFilter.frx":0E24
         OLEDBString     =   $"QeybatFilter.frx":0EAD
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from student"
         Caption         =   "Student"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
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
         Connect         =   $"QeybatFilter.frx":0F36
         OLEDBString     =   $"QeybatFilter.frx":0FBF
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from tarhha"
         Caption         =   "Tarhha"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc ekhtar 
         Height          =   330
         Left            =   360
         Top             =   2520
         Visible         =   0   'False
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
         Connect         =   $"QeybatFilter.frx":1048
         OLEDBString     =   $"QeybatFilter.frx":10D1
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from ekhtar"
         Caption         =   "ekhtar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin VB.ComboBox Combo8 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      ItemData        =   "QeybatFilter.frx":115A
      Left            =   240
      List            =   "QeybatFilter.frx":115C
      TabIndex        =   5
      Text            =   "1390"
      Top             =   720
      Width           =   855
   End
   Begin VB.ComboBox Combo7 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1080
      TabIndex        =   4
      Text            =   "01"
      Top             =   720
      Width           =   735
   End
   Begin VB.ComboBox Combo4 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      ItemData        =   "QeybatFilter.frx":115E
      Left            =   1800
      List            =   "QeybatFilter.frx":1160
      TabIndex        =   3
      Text            =   "01"
      Top             =   720
      Width           =   735
   End
   Begin VB.ComboBox Combo6 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      ItemData        =   "QeybatFilter.frx":1162
      Left            =   240
      List            =   "QeybatFilter.frx":1164
      TabIndex        =   2
      Text            =   "1390"
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1080
      TabIndex        =   1
      Text            =   "01"
      Top             =   120
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      ItemData        =   "QeybatFilter.frx":1166
      Left            =   1800
      List            =   "QeybatFilter.frx":1168
      TabIndex        =   0
      Text            =   "01"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " «  «—ÌŒ"
      Height          =   315
      Left            =   2640
      TabIndex        =   10
      Top             =   720
      Width           =   480
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«“  «—ÌŒ"
      Height          =   315
      Left            =   2640
      TabIndex        =   9
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "QeybatFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command14_Click()
On Error GoTo 1

GoTo 2

1:
MsgBox "⁄œœ Ê«—œ ‘œÂ »Ì‘ «“ Õœ „Ã«“ „Ì »«‘œ", vbCritical + vbOKOnly, "Œÿ«"



Exit Sub
2:
QeybatF.Qeybat.Refresh
Label1.Caption = "œ— Õ«· »——”Ì"
Dim I As Double
Dim SearchData, START, Ennd As String
'Start = Str(Combo6.Text) & Str(Combo3.Text) & Str(Combo1.Text)
'Ennd = Str(Combo8.Text) & Str(Combo7.Text) & Str(Combo4.Text)


START = Combo6.Text & "" & Combo3.Text & "" & Combo1.Text
Ennd = Combo8.Text & "" & Combo7.Text & "" & Combo4.Text

SearchData = ""

If Val(START) > Val(Ennd) Then
MsgBox " «—ÌŒ Å«Ì«‰ “Êœ  — «“  «—ÌŒ ‘—Ê⁄ „Ì »«‘œ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
End If


For I = Val(START) To Val(Ennd)
If I = Val(START) Then
SearchData = " kodqeybat like ('" & I & "')"
Else
SearchData = SearchData & " or kodqeybat like ('" & I & "')"
End If
Next I



QeybatF.Qeybat.Refresh
QeybatF.Qeybat.RecordSource = "select * from qeybat where" & SearchData
QeybatF.Qeybat.Refresh
Label1.Caption = " ⁄œ«œ „Ê«—œ Ì«›  ‘œÂ"

Me.Label51.Caption = QeybatF.Qeybat.Recordset.RecordCount







End Sub

Private Sub Command15_Click()
QeybatF.Show
Unload Me

End Sub

Private Sub Form_Load()
Dim I As Integer

For I = 1390 To 1408
Combo6.AddItem (I)
Combo8.AddItem (I)

Next I


For I = 1 To 31 Step 1  ' »—«Ì Ê«—œ ﬂ—œ‰ ‘„«—Â —Ê“ œ— ÃœÊ· «ÿ·«⁄«  €Ì 
If I < 10 Then
Combo1.AddItem ("0" & I)
Combo4.AddItem ("0" & I)

Else
Combo1.AddItem (I)
Combo4.AddItem (I)
End If
Next I


For I = 1 To 12 Step 1  ' »—«Ì Ê«—œ ﬂ—œ‰ ‘„«—Â —Ê“ œ— ÃœÊ· «ÿ·«⁄«  €Ì 
If I < 10 Then
Combo3.AddItem ("0" & I)
Combo7.AddItem ("0" & I)

Else
Combo3.AddItem (I)
Combo7.AddItem (I)
End If
Next I










End Sub

