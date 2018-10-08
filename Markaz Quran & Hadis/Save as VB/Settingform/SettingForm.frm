VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form SettingTaahod 
   Caption         =   " ‰ŸÌ„« "
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9930
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SettingForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "À»  ‰ ŸÌ„« "
      Height          =   3375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   9735
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "ÃœÌœ"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H0080FF80&
         Caption         =   "À» "
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Ã«Ìê“Ì‰Ì"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Õ–›"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox TE5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   350
         Left            =   1560
         TabIndex        =   3
         Top             =   2760
         Width           =   7320
      End
      Begin VB.TextBox TE4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   1845
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   840
         Width           =   7320
      End
      Begin VB.TextBox TE3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   350
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   2520
      End
      Begin VB.TextBox TE2 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   350
         Left            =   4800
         TabIndex        =   9
         Text            =   "101"
         Top             =   360
         Width           =   720
      End
      Begin VB.TextBox TE1 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   350
         Left            =   6120
         TabIndex        =   8
         Text            =   "QeybatF-TaahodKatbi-Text"
         Top             =   360
         Width           =   2760
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   " Ê÷ÌÕ« "
         Height          =   315
         Left            =   8880
         TabIndex        =   14
         Top             =   2760
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "òœ"
         Height          =   315
         Left            =   9120
         TabIndex        =   13
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "‘„«—Â"
         Height          =   315
         Left            =   5640
         TabIndex        =   12
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "⁄‰Ê«‰"
         Height          =   315
         Left            =   4320
         TabIndex        =   11
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "„ ‰"
         Height          =   315
         Left            =   9000
         TabIndex        =   10
         Top             =   1440
         Width           =   240
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "SettingForm.frx":08CA
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   7011
      _Version        =   393216
      BackColor       =   12632319
      HeadLines       =   1
      RowHeight       =   27
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "XkodSetting"
         Caption         =   "òœ ”Ì” „"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "XSort"
         Caption         =   "‘„«—Â"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "XName"
         Caption         =   "⁄‰Ê«‰"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "XText"
         Caption         =   "„ ‰"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Tozih"
         Caption         =   " Ê÷ÌÕ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Op"
         Caption         =   "Op"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "D"
         Caption         =   "D"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Other1"
         Caption         =   "Other1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Other2"
         Caption         =   "Other2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Other3"
         Caption         =   "Other3"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "Other4"
         Caption         =   "Other4"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "Other5"
         Caption         =   "Other5"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429.858
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Setting 
      Height          =   330
      Left            =   1200
      Top             =   8880
      Width           =   1935
      _ExtentX        =   3413
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
      Connect         =   $"SettingForm.frx":08E0
      OLEDBString     =   $"SettingForm.frx":0969
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from settingtable"
      Caption         =   "Setting"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "SettingTaahod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Setting.Refresh
Setting.RecordSource = "select * from settingtable where xkodsetting like ('%" + "QeybatF-TaahodKatbi-Text" + "%')"
Setting.Refresh

Setting.Recordset.Sort = "xsort"

Setting.Recordset.MovePrevious

Setting.Recordset.MoveLast

TE2.Text = Val(Setting.Recordset.Fields("xsort")) + 1


Beep


End Sub

Private Sub Command3_Click()
On Error Resume Next
If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰ „Ê—œ —« Õ–› ò‰Ìœ", vbQuestion + vbYesNo, "Õ–›") = vbYes Then

Setting.Recordset.Delete
End If

End Sub

Private Sub Command5_Click()
If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ «ÿ·«⁄«  À»  ‘Êœ", vbQuestion + vbYesNo, "À»  «ÿ·«⁄« ") = vbNo Then Exit Sub



Setting.Refresh
Setting.Recordset.AddNew
Setting.Recordset.Fields("xkodsetting") = TE1.Text
Setting.Recordset.Fields("xsort") = TE2.Text
Setting.Recordset.Fields("xname") = TE3.Text
Setting.Recordset.Fields("xtext") = TE4.Text
Setting.Recordset.Fields("tozih") = TE5.Text
Setting.Recordset.Fields("op") = Entekhab.SB.Panels(1).Text
'Setting.Recordset.Fields("d") = "@"
Setting.Recordset.Fields("other1") = "@"
Setting.Recordset.Fields("other2") = "@"
Setting.Recordset.Fields("other3") = "@"
Setting.Recordset.Fields("other4") = "@"
Setting.Recordset.Fields("other5") = "@"
Setting.Recordset.Update
Setting.Refresh
MsgBox "«ÿ·«⁄«  À»  ‘œ", vbInformation + vbOKOnly, "À»  «ÿ·«⁄« "

End Sub

Private Sub Command6_Click()

Setting.Recordset.Fields("xkodsetting") = TE1.Text
Setting.Recordset.Fields("xsort") = TE2.Text
Setting.Recordset.Fields("xname") = TE3.Text
Setting.Recordset.Fields("xtext") = TE4.Text
Setting.Recordset.Fields("tozih") = TE5.Text
Setting.Recordset.Fields("op") = Entekhab.SB.Panels(1).Text
'Setting.Recordset.Fields("d") = "@"
Setting.Recordset.Fields("other1") = "@"
Setting.Recordset.Fields("other2") = "@"
Setting.Recordset.Fields("other3") = "@"
Setting.Recordset.Fields("other4") = "@"
Setting.Recordset.Fields("other5") = "@"
Setting.Recordset.Update
Setting.Refresh
Beep
MsgBox "«ÿ·«⁄«  Ã«Ìê“Ì‰ ‘œ", vbInformation + vbOKOnly, "Ã«Ìê“Ì‰Ì «ÿ·«⁄« "
End Sub

Private Sub DataGrid1_Click()
On Error Resume Next

Setting.Recordset.Update
End Sub

Private Sub DataGrid1_DblClick()
'Setting.Refresh
Setting.RecordSource = "select * from settingtable where xkodsetting like ('%" + Setting.Recordset.Fields("xkodsetting") + "%') and xsort like ('%" + Setting.Recordset.Fields("xsort") + "%')"
Setting.Refresh
 TE1.Text = Setting.Recordset.Fields("xkodsetting")
 TE2.Text = Setting.Recordset.Fields("xsort")
 TE3.Text = Setting.Recordset.Fields("xname")
TE4.Text = Setting.Recordset.Fields("xtext")
TE5.Text = Setting.Recordset.Fields("tozih")

End Sub

Private Sub Form_Load()
Setting.Refresh
Setting.RecordSource = "select * from settingtable where xkodsetting like ('%" + "QeybatF-TaahodKatbi-Text" + "%')"
Setting.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
QeybatF.Show

End Sub

