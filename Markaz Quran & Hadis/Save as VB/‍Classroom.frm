VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FClassroom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„œÌ—Ì  ·Ì”  ﬂ·«”Ì"
   ClientHeight    =   10320
   ClientLeft      =   1770
   ClientTop       =   1380
   ClientWidth     =   13095
   BeginProperty Font 
      Name            =   "B Homa"
      Size            =   9
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ûClassroom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10320
   ScaleWidth      =   13095
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      ItemData        =   "ûClassroom.frx":08CA
      Left            =   120
      List            =   "ûClassroom.frx":08CC
      TabIndex        =   80
      Text            =   "«‰ Œ«» ò‰Ìœ"
      Top             =   6960
      Width           =   6375
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H0080FF80&
      Caption         =   "„‘«ÂœÂ ·Ì”  ò·«”Ì"
      Height          =   975
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   7440
      Width           =   855
   End
   Begin VB.CommandButton Command11 
      DisabledPicture =   "ûClassroom.frx":08CE
      DownPicture     =   "ûClassroom.frx":25548
      DragIcon        =   "ûClassroom.frx":4A1C2
      Height          =   330
      Left            =   120
      Picture         =   "ûClassroom.frx":6EE3C
      Style           =   1  'Graphical
      TabIndex        =   76
      ToolTipText     =   "«‰ ﬁ«· ÃœÊ· »Â »—‰«„Â «ò”·"
      Top             =   9600
      Width           =   375
   End
   Begin VB.CommandButton Command10 
      DisabledPicture =   "ûClassroom.frx":93AB6
      DownPicture     =   "ûClassroom.frx":B8730
      DragIcon        =   "ûClassroom.frx":DD3AA
      Height          =   330
      Left            =   600
      Picture         =   "ûClassroom.frx":102024
      Style           =   1  'Graphical
      TabIndex        =   75
      ToolTipText     =   "€Ì» "
      Top             =   9600
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      DisabledPicture =   "ûClassroom.frx":126C9E
      DownPicture     =   "ûClassroom.frx":14B918
      DragIcon        =   "ûClassroom.frx":170592
      Height          =   330
      Left            =   1080
      Picture         =   "ûClassroom.frx":19520C
      Style           =   1  'Graphical
      TabIndex        =   74
      ToolTipText     =   "« „«„ ò·«”"
      Top             =   9600
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      DisabledPicture =   "ûClassroom.frx":1B9E86
      DownPicture     =   "ûClassroom.frx":1DEB00
      DragIcon        =   "ûClassroom.frx":20377A
      Height          =   330
      Left            =   9120
      Picture         =   "ûClassroom.frx":2283F4
      Style           =   1  'Graphical
      TabIndex        =   73
      ToolTipText     =   "«‰ ﬁ«· ÃœÊ· »Â »—‰«„Â «ò”·"
      Top             =   9600
      Width           =   375
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   150
      Left            =   12120
      TabIndex        =   72
      Top             =   7080
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   265
      _Version        =   327682
      Appearance      =   0
      Max             =   5
   End
   Begin VB.ListBox List4 
      BackColor       =   &H80000002&
      Height          =   2040
      ItemData        =   "ûClassroom.frx":24D06E
      Left            =   120
      List            =   "ûClassroom.frx":24D070
      Sorted          =   -1  'True
      TabIndex        =   64
      Top             =   7440
      Width           =   6375
   End
   Begin VB.ListBox List3 
      BackColor       =   &H80000002&
      Height          =   2700
      ItemData        =   "ûClassroom.frx":24D072
      Left            =   5280
      List            =   "ûClassroom.frx":24D074
      Sorted          =   -1  'True
      TabIndex        =   62
      Top             =   10320
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.ListBox List2 
      BackColor       =   &H80000002&
      Height          =   2700
      ItemData        =   "ûClassroom.frx":24D076
      Left            =   8280
      List            =   "ûClassroom.frx":24D078
      Sorted          =   -1  'True
      TabIndex        =   60
      Top             =   10320
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ç«Å ·Ì” "
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   3015
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0080C0FF&
      Caption         =   "Å«ò ò—œ‰ ÃœÊ·"
      Height          =   930
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8520
      Width           =   855
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000002&
      Height          =   2040
      ItemData        =   "ûClassroom.frx":24D07A
      Left            =   6600
      List            =   "ûClassroom.frx":24D07C
      Sorted          =   -1  'True
      TabIndex        =   56
      Top             =   7440
      Width           =   5415
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H8000000E&
      Caption         =   "À»  « „«„ ò·«”"
      Height          =   375
      Left            =   13320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000002&
      Caption         =   "„Ãœœ"
      Height          =   330
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   120
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "ûClassroom.frx":24D07E
      Height          =   135
      Left            =   16560
      TabIndex        =   54
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   238
      _Version        =   393216
      BackColor       =   12640511
      DefColWidth     =   87
      HeadLines       =   1
      RowHeight       =   29
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "·Ì”  ò·«”Ì ﬁ—¬‰ ¬„Ê“«‰"
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "Parvande"
         Caption         =   "Parvande"
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
         DataField       =   "KOdclass"
         Caption         =   "KOdclass"
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
         DataField       =   "Tshoro"
         Caption         =   "Tshoro"
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
         DataField       =   "Tpayan"
         Caption         =   "Tpayan"
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
         DataField       =   "Elat"
         Caption         =   "Elat"
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
         DataField       =   "Tozih"
         Caption         =   "Tozih"
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
      BeginProperty Column07 
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1319.811
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   9960
      TabIndex        =   6
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Frame Frame4 
      Caption         =   "„‘Œ’«  ò·«”"
      Height          =   3975
      Left            =   120
      TabIndex        =   34
      Top             =   2880
      Width           =   3615
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "«Ì«„ Â› Â"
         Height          =   330
         Left            =   2760
         TabIndex        =   87
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Ayamehafte"
         DataSource      =   "MClass"
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
         Left            =   2160
         TabIndex        =   86
         Top             =   3480
         Width           =   135
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   " «"
         Height          =   330
         Left            =   1680
         TabIndex        =   52
         Top             =   1920
         Width           =   120
      End
      Begin VB.Label lzsho 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "ZamaneShoro"
         DataSource      =   "MClass"
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
         Left            =   2160
         TabIndex        =   51
         Top             =   1920
         Width           =   135
      End
      Begin VB.Label lmadras 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "-  "
         DataField       =   "Madras"
         DataSource      =   "MClass"
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
         Left            =   2160
         TabIndex        =   50
         Top             =   2280
         Width           =   225
      End
      Begin VB.Label lostad 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Ostad"
         DataSource      =   "MClass"
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
         Left            =   2160
         TabIndex        =   49
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label lzpa 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "ZamanePayan"
         DataSource      =   "MClass"
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
         Left            =   1080
         TabIndex        =   48
         Top             =   1920
         Width           =   135
      End
      Begin VB.Label lmaqta 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Maqta"
         DataSource      =   "MClass"
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
         Left            =   2160
         TabIndex        =   47
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label ltarh 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Tarh"
         DataSource      =   "MClass"
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
         Left            =   2160
         TabIndex        =   46
         Top             =   720
         Width           =   135
      End
      Begin VB.Label lkodclass 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "KodClass"
         DataSource      =   "MClass"
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
         Left            =   2160
         TabIndex        =   45
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "„œ—”"
         Height          =   330
         Left            =   2760
         TabIndex        =   44
         Top             =   2280
         Width           =   405
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "”«⁄ "
         Height          =   330
         Left            =   2760
         TabIndex        =   43
         Top             =   1920
         Width           =   405
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "‰«„ «” «œ"
         Height          =   330
         Left            =   2760
         TabIndex        =   42
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "„ﬁÿ⁄"
         Height          =   330
         Left            =   2760
         TabIndex        =   41
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "ÿ—Õ"
         Height          =   330
         Left            =   2760
         TabIndex        =   40
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "òœ ò·«”"
         Height          =   330
         Index           =   0
         Left            =   2760
         TabIndex        =   39
         Top             =   360
         Width           =   555
      End
      Begin VB.Label ltsho 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "TShoro"
         DataSource      =   "MClass"
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
         Left            =   2160
         TabIndex        =   38
         Top             =   2640
         Width           =   135
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ ‘—Ê⁄"
         Height          =   345
         Left            =   2760
         TabIndex        =   37
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label ltpa 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "TPayan"
         DataSource      =   "MClass"
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
         Left            =   2160
         TabIndex        =   36
         Top             =   3000
         Width           =   135
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ Å«Ì«‰"
         Height          =   345
         Left            =   2760
         TabIndex        =   35
         Top             =   3120
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "„‘Œ’«  Ê—Êœ Ê Œ—ÊÃ ﬁ—¬‰ ¬„Ê“ «“ ò·«”"
      Height          =   2895
      Left            =   9960
      TabIndex        =   29
      Top             =   3960
      Width           =   3015
      Begin VB.ComboBox telat 
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   83
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox ttozih 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   240
         TabIndex        =   10
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox ttpayan 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   870
         Width           =   1455
      End
      Begin VB.TextBox ttshoro 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   " Ê÷ÌÕ"
         Height          =   330
         Left            =   2520
         TabIndex        =   33
         Top             =   2040
         Width           =   405
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ Å«Ì«‰ ò·«”"
         Height          =   330
         Left            =   1800
         TabIndex        =   32
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "⁄·  "
         Height          =   330
         Left            =   2520
         TabIndex        =   31
         Top             =   1560
         Width           =   330
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " «—Œ Ê—Êœ »Â ò·«”"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1800
         TabIndex        =   30
         Top             =   360
         Width           =   1125
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "ûClassroom.frx":24D096
      Height          =   3255
      Left            =   3840
      TabIndex        =   7
      Top             =   3600
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5741
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777152
      DefColWidth     =   107
      HeadLines       =   1
      RowHeight       =   29
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "·Ì”  ò·«” Â«"
      ColumnCount     =   17
      BeginProperty Column00 
         DataField       =   "Kodclass"
         Caption         =   "òœ ò·«”"
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
         DataField       =   "Tarh"
         Caption         =   "ÿ—Õ"
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
         DataField       =   "Maqta"
         Caption         =   "„ﬁÿ⁄"
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
         DataField       =   "Tozih"
         Caption         =   " Ê÷ÌÕ« "
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
         DataField       =   "Ostad"
         Caption         =   "«” «œ"
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
         DataField       =   "Zamaneshoro"
         Caption         =   "”«⁄  ‘—Ê⁄"
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
         DataField       =   "Zamanepayan"
         Caption         =   "”«⁄  Å«Ì«‰"
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
         DataField       =   "Madras"
         Caption         =   "„œ—”"
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
         DataField       =   "Ayamehafte"
         Caption         =   "«Ì«„ Â› Â"
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
         DataField       =   "Tshoro"
         Caption         =   " «—ÌŒ ‘—Ê⁄"
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
         DataField       =   "Tpayan"
         Caption         =   " «—ÌŒ ÅÌ«Ì«‰"
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
         DataField       =   "Tedadjalasat"
         Caption         =   " ⁄œ«œ Ã·”« "
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
      BeginProperty Column12 
         DataField       =   "Sobh"
         Caption         =   "’»Õ"
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
      BeginProperty Column13 
         DataField       =   "Asr"
         Caption         =   "⁄’—"
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
      BeginProperty Column14 
         DataField       =   "Op"
         Caption         =   "«Å—« Ê—"
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
      BeginProperty Column15 
         DataField       =   "D"
         Caption         =   " «—ÌŒ À» "
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
      BeginProperty Column16 
         DataField       =   "QMojaz"
         Caption         =   " ⁄œ«œ €Ì»  „Ã«“"
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
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   1620.284
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "„‘Œ’«  ﬁ—¬‰ ¬„Ê“"
      Height          =   2895
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   3615
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   82
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Tell"
         DataSource      =   "Student"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2640
         TabIndex        =   85
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Mob"
         DataSource      =   "Student"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1200
         TabIndex        =   84
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ò·«” Â«Ì ﬁ—¬‰ ¬„Ê“"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   81
         Top             =   2400
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "‘„«—Â Å—Ê‰œÂ"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2040
         TabIndex        =   28
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "‰«„"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   27
         Top             =   720
         Width           =   195
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "‰«„ Œ«‰Ê«œêÌ"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   26
         Top             =   1080
         Width           =   870
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Parvande"
         DataSource      =   "Student"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   1680
         TabIndex        =   25
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "name"
         DataSource      =   "Student"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   1680
         TabIndex        =   24
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "famil"
         DataSource      =   "Student"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   1680
         TabIndex        =   23
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "‰«ﬁ’Ì Å—Ê‰œÂ"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   22
         Top             =   1440
         Width           =   885
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Tozih"
         DataSource      =   "Student"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1680
         TabIndex        =   21
         Top             =   1440
         Width           =   135
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Motor"
      Height          =   255
      Left            =   9960
      TabIndex        =   19
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
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
         Connect         =   $"ûClassroom.frx":24D0AB
         OLEDBString     =   $"ûClassroom.frx":24D134
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *  from stu2class"
         Caption         =   "STU2CLASS"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Homa"
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
         Connect         =   $"ûClassroom.frx":24D1BD
         OLEDBString     =   $"ûClassroom.frx":24D246
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from mclass"
         Caption         =   "mclass"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Connect         =   $"ûClassroom.frx":24D2CF
         OLEDBString     =   $"ûClassroom.frx":24D358
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
         Connect         =   $"ûClassroom.frx":24D3E1
         OLEDBString     =   $"ûClassroom.frx":24D46A
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
            Charset         =   0
            Weight          =   700
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
         Connect         =   $"ûClassroom.frx":24D4F3
         OLEDBString     =   $"ûClassroom.frx":24D57C
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from tarhha"
         Caption         =   "Tarhha"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc userprofiletable 
         Height          =   330
         Left            =   0
         Top             =   0
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
         Connect         =   $"ûClassroom.frx":24D605
         OLEDBString     =   $"ûClassroom.frx":24D68E
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from userprofiletable"
         Caption         =   "userprofiletable"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Õ–› «“ ·Ì”  ﬂ·«”"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6600
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6960
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ã” ÃÊ »— «”«”"
      Height          =   1455
      Left            =   11520
      TabIndex        =   12
      Top             =   960
      Width           =   1455
      Begin VB.OptionButton Option5 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂœ ﬂ·«”"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Alignment       =   1  'Right Justify
         Caption         =   "‰«„ «” «œ"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         Caption         =   "ﬂœ „·Ì"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2400
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         Caption         =   "‰«„ Œ«‰Ê«œêÌ ‘„«—Â Å—Ê‰œÂ"
         Height          =   615
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "‘„«—Â Å—Ê‰œÂ"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Â„—«Â"
         Height          =   330
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "À»  œ— ·Ì”  ﬂ·«”Ì"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin MSDataGridLib.DataGrid DataGridSTUDENT 
      Bindings        =   "ûClassroom.frx":24D717
      Height          =   3375
      Left            =   3840
      TabIndex        =   78
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12648384
      DefColWidth     =   120
      HeadLines       =   1
      RowHeight       =   26
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Homa"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "«ÿ·«⁄«  ﬁ—¬‰ ¬„Ê“«‰"
      ColumnCount     =   27
      BeginProperty Column00 
         DataField       =   "Parvande"
         Caption         =   "‘„«—Â Å—Ê‰œÂ"
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
         DataField       =   "Name"
         Caption         =   "‰«„"
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
         DataField       =   "Famil"
         Caption         =   "‰«„ Œ«‰Ê«œêÌ"
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
         DataField       =   "Namepedar"
         Caption         =   "‰«„ Åœ—"
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
         DataField       =   "Tavalod"
         Caption         =   " «—ÌŒ  Ê·œ"
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
         DataField       =   "Shsh"
         Caption         =   "‘„«—Â ‘‰«”‰«„Â"
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
         DataField       =   "Sadere"
         Caption         =   "’«œ—Â"
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
         DataField       =   "Meliyat"
         Caption         =   "„·Ì "
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
         DataField       =   "Mazhab"
         Caption         =   "„–Â»"
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
         DataField       =   "Kodmeli"
         Caption         =   "òœ „·Ì"
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
         DataField       =   "Gozarname"
         Caption         =   "‘„«—Â ê–“ ‰«„Â"
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
         DataField       =   "Taahol"
         Caption         =   "Ê÷⁄Ì   «Â·"
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
      BeginProperty Column12 
         DataField       =   "Farzand"
         Caption         =   " ⁄œ«œ ›—“‰œ«‰"
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
      BeginProperty Column13 
         DataField       =   "Tahsilat"
         Caption         =   " Õ’Ì·« "
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
      BeginProperty Column14 
         DataField       =   "Ostad"
         Caption         =   "‰«„ «” «œ"
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
      BeginProperty Column15 
         DataField       =   "Tozih"
         Caption         =   " Ê÷ÌÕ« "
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
      BeginProperty Column16 
         DataField       =   "Tell"
         Caption         =   " ·›‰ À«» "
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
      BeginProperty Column17 
         DataField       =   "Mob"
         Caption         =   "Â„—«Â"
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
      BeginProperty Column18 
         DataField       =   "Scan"
         Caption         =   "«”ò‰ ›«Ì·"
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
      BeginProperty Column19 
         DataField       =   "Clas1"
         Caption         =   "ò·«” 1"
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
      BeginProperty Column20 
         DataField       =   "Op"
         Caption         =   "«Å—« Ê—"
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
      BeginProperty Column21 
         DataField       =   "D"
         Caption         =   " «—ÌŒ À»  «ÿ·«⁄« "
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
      BeginProperty Column22 
         DataField       =   "Clas2"
         Caption         =   "ò·«”2"
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
      BeginProperty Column23 
         DataField       =   "Clas3"
         Caption         =   "ò·«” 3"
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
      BeginProperty Column24 
         DataField       =   "Clas4"
         Caption         =   "ò·«” 4"
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
      BeginProperty Column25 
         DataField       =   "Clas5"
         Caption         =   "ò·«” 5"
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
      BeginProperty Column26 
         DataField       =   "NF"
         Caption         =   "‰«„ Ê ‰«„ Œ«‰Ê«œêÌ"
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
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column24 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column25 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column26 
            Object.Visible         =   0   'False
            ColumnWidth     =   1814.74
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar stb1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   79
      Top             =   9945
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "ò«—»— Ã«—Ì"
            TextSave        =   "ò«—»— Ã«—Ì"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4304
            MinWidth        =   4304
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   " «—ÌŒ «„—Ê“"
            TextSave        =   " «—ÌŒ «„—Ê“"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label38 
      AutoSize        =   -1  'True
      Caption         =   " ⁄œ«œ "
      Height          =   330
      Left            =   2400
      TabIndex        =   71
      Top             =   9600
      Width           =   405
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   1800
      TabIndex        =   70
      Top             =   9600
      Width           =   75
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      Caption         =   " ⁄œ«œ "
      Height          =   330
      Left            =   8400
      TabIndex        =   69
      Top             =   9600
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   7920
      TabIndex        =   68
      Top             =   9600
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   " ⁄œ«œ "
      Height          =   330
      Left            =   5400
      TabIndex        =   67
      Top             =   9600
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   4920
      TabIndex        =   66
      Top             =   9600
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "·Ì”  ò·Ì ò·«”"
      Height          =   330
      Left            =   480
      TabIndex        =   65
      Top             =   10560
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "·Ì”  ﬁ—¬„Ê“«‰ Õ–› ‘œÂ »Â œ·Ì· €Ì» "
      Height          =   330
      Left            =   1080
      TabIndex        =   63
      Top             =   11160
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "·Ì”  « „«„ ò·«”"
      Height          =   330
      Left            =   11760
      TabIndex        =   61
      Top             =   10920
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF8080&
      X1              =   9960
      X2              =   12960
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF8080&
      X1              =   9960
      X2              =   12960
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   9960
      X2              =   11400
      Y1              =   1995
      Y2              =   1995
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   11040
      TabIndex        =   59
      Top             =   9600
      Width           =   75
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      Caption         =   " ⁄œ«œ "
      Height          =   330
      Left            =   11520
      TabIndex        =   58
      Top             =   9600
      Width           =   405
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "·Ì”  ﬁ—¬‰ ¬„Ê“«‰ Õ«÷— œ— ò·«”"
      Height          =   330
      Left            =   9840
      TabIndex        =   57
      Top             =   6960
      Width           =   2100
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "Ã” ÃÊ œ— ÿ—Õ Ê „ﬁÿ⁄ Ê  ‰«„ «” «œ Ê òœ ò·«”"
      Height          =   330
      Left            =   9960
      TabIndex        =   53
      Top             =   3120
      Width           =   2925
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   330
      Left            =   11400
      TabIndex        =   0
      Top             =   120
      Width           =   45
   End
   Begin VB.Menu mnuhiome 
      Caption         =   "#"
   End
   Begin VB.Menu PR 
      Caption         =   "Å—Ê‰œÂ"
      Begin VB.Menu mnusabtdar 
         Caption         =   "À»  œ— ·Ì”  ò·«”Ì"
         Shortcut        =   ^S
      End
      Begin VB.Menu fEXI 
         Caption         =   "›«Ì· Œ—ÊÃÌ"
         Begin VB.Menu lclasslist1 
            Caption         =   "·Ì”  ò·«”Ì 1"
            Checked         =   -1  'True
         End
         Begin VB.Menu lclasslist 
            Caption         =   "·Ì”  ò·«”Ì 2"
         End
         Begin VB.Menu mnu_sabt_nomre 
            Caption         =   "·Ì”  À»  ‰„—« "
         End
         Begin VB.Menu elana4 
            Caption         =   "·Ì”  «⁄·«‰ (A4)"
         End
         Begin VB.Menu elana5 
            Caption         =   "·Ì”  «⁄·«‰ (A5)"
         End
         Begin VB.Menu foroshlist 
            Caption         =   "·Ì”  ›—Ê‘ ò «» (A5)"
         End
      End
      Begin VB.Menu wee 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuetmam 
         Caption         =   "À»  « „«„ ò·«”"
      End
      Begin VB.Menu wds 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnudellall 
         Caption         =   "Õ–› ﬁ—¬‰ ¬„Ê“ «“ ò·«” »Â ’Ê—  ò«„·"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "FClassroom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
mclass.Refresh
mclass.RecordSource = "select * from mclass where kodclass like ('%" + Combo1.Text + "%')"
mclass.Refresh
End Sub

Private Sub Combo2_Click()
List4.Clear

If Combo2.Text = "ﬁ—¬‰ ¬„Ê“«‰ Õ–› ‘œÂ »Â œ·Ì· €Ì» " Then
STU2CLASS.Refresh
STU2CLASS.RecordSource = "select * from STU2CLASS where kodclass like ('%" + lkodclass.Caption + "%') and elat like ('%" + "€Ì» " + "%') "
STU2CLASS.Refresh
Label37.Caption = STU2CLASS.Recordset.RecordCount

For I = 1 To STU2CLASS.Recordset.RecordCount

Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + STU2CLASS.Recordset.Fields("parvande") + "%') "
Student.Refresh

List4.AddItem (Student.Recordset.Fields("parvande") & "  -  " & Student.Recordset.Fields("name") & "  -  " & Student.Recordset.Fields("famil"))
STU2CLASS.Recordset.MoveNext
Next I


End If

If Combo2.Text = "·Ì”  ò«„· ò·«”" Then
STU2CLASS.Refresh
STU2CLASS.RecordSource = "select * from STU2CLASS where kodclass like ('%" + lkodclass.Caption + "%') "
STU2CLASS.Refresh

Label37.Caption = STU2CLASS.Recordset.RecordCount

For I = 1 To STU2CLASS.Recordset.RecordCount

Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + STU2CLASS.Recordset.Fields("parvande") + "%') "
Student.Refresh

List4.AddItem (Student.Recordset.Fields("parvande") & "  -  " & Student.Recordset.Fields("name") & "  -  " & Student.Recordset.Fields("famil"))
STU2CLASS.Recordset.MoveNext
Next I

End If

If Combo2.Text = "·Ì”  « „«„ ò·«”" Then

STU2CLASS.Refresh
STU2CLASS.RecordSource = "select * from STU2CLASS where kodclass like ('%" + lkodclass.Caption + "%') and elat like ('%" + "« „«„" + "%') "
STU2CLASS.Refresh
Label37.Caption = STU2CLASS.Recordset.RecordCount

For I = 1 To STU2CLASS.Recordset.RecordCount

Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + STU2CLASS.Recordset.Fields("parvande") + "%') "
Student.Refresh

List4.AddItem (Student.Recordset.Fields("parvande") & "  -  " & Student.Recordset.Fields("name") & "  -  " & Student.Recordset.Fields("famil"))
STU2CLASS.Recordset.MoveNext
Next I




End If






End Sub

Private Sub Command1_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "fclass-newsabt" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513

Dim X, TC As Integer
Dim C As String

If mclass.Recordset.Fields("tozih") = "« „«„ ò·«”" Then
MsgBox "«Ì‰ ò·«” »Â « „«„ —”ÌœÂ «” ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
End If



Student.RecordSource = "select * from student where parvande like ('%" + Label8.Caption + "%')" ' ‰«„ ﬁ—¬‰ ¬„Ê“ —« »Â  ‰Â«ÌÌ „Ì ¬Ê—œ
Student.Refresh


'«Ì‰ òÂ ÂÌÃÌ ‰Ì”  ›ﬁÿ ⁄œœ «” 

X = 1
TC = 0
'«Ì‰ òÂ ÂÌÃÌ ‰Ì” 


For I = 1 To 5   ' Å‰Ã ò·«” —« çò „Ì ò‰œ

C = "clas" & X
If Student.Recordset.Fields(C) <> "‰œ«—œ" Then
TC = TC + 1 '‘„«—‘ ò·«” Â«

End If
X = X + 1
Next I
If TC >= 1 Then  'ÿ—› »Ì” — «“ 1 ‘—ò  „Ì ò‰œ

If MsgBox("«Ì‰ ﬁ—¬‰ ¬„Ê“ Â„ «ò‰Ê‰ œ— ò·«” œÌê—Ì ‘—ò  „Ì ò‰œ  ¬Ì« „Ì ŒÊ«ÂÌœ ‰«„ «Ì‘«‰ œ— «Ì‰ ò·«” Â„ À»  ‘Êœ", vbQuestion + vbYesNo, " „œÌ—Ì  ·Ì”  ò·«”Ì") = vbYes Then

GoTo 12 '«ÿ„Ì‰«‰ œ«—œ «“ òœ ò·«”


Else  ' «“ ’Õ  òœ ò·«” «ÿ„Ì‰«‰ ‰œ«—œ

Exit Sub
End If
End If

12:

'¬Ì« ﬁ»·« œ— «Ì‰ ò·«”Ì »Êœ Â Ì« ‰Â⁄ç


STU2CLASS.Refresh
STU2CLASS.RecordSource = " select * from stu2class where  parvande like ('%" + Label8.Caption + "%') and kodclass like ('%" + lkodclass.Caption + "%')"
STU2CLASS.Refresh


If STU2CLASS.Recordset.BOF = True Or STU2CLASS.Recordset.EOF = True Then   'çò „Ì ò‰œ òÂ ¬Ì« œ— «Ì‰ ò·«” »ÊœÂ Ì« ‰Â
   GoTo 1  ' « Â«·«  Ê «Ì‰ ò·«” ‘—ò  ‰ò—œÂ
Else   ' ﬁ»·«  Ê «Ì‰ ò·«” »ÊœÂ »Â œ·Ì·Ì Õ–› ‘œÂ Õ«·« „Ì ŒÊœ œÊ»«—Â »—Â
    MsgBox "«Ì‰ ﬁ—¬‰ ¬„Ê“ ﬁ»·« œ— «Ì‰ ò·«” À»  ‘œÂ «” ", vbExclamation + vbOKOnly, "„œÌ—Ì  ·Ì”  ò·«”Ì"
    
        If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ «ÿ·«⁄«  »Â —Ê“Â ‘Êœ", vbQuestion + vbYesNo, "„œÌ—Ì  ·Ì”  ò·«”Ì") = vbYes Then
        
        
        
        
        
        
        
        

        
        
        
        
        
        
        
            
            STU2CLASS.Refresh
            STU2CLASS.RecordSource = " select * from stu2class where  parvande like ('%" + Label8.Caption + "%') and kodclass like ('%" + lkodclass.Caption + "%')"
    STU2CLASS.Refresh
            If Student.Recordset.Fields("clas1") = lkodclass.Caption Or Student.Recordset.Fields("clas2") = lkodclass.Caption Or Student.Recordset.Fields("clas3") = lkodclass.Caption Or Student.Recordset.Fields("clas4") = lkodclass.Caption Or Student.Recordset.Fields("clas5") = lkodclass.Caption Then
            '«Ì‰ ‘—ÿ Ì⁄‰Ì «·«‰ «Ì‰ ﬁ—¬‰ ¬„Ê“ œ— «Ì‰ ò·«” Â”  Ê Â‰Ê“ Õœ› ‰ ‘œÂ
            MsgBox "«Ì‰ ﬁ—¬‰ ¬„Ê“ Â„ «ò‰Ê‰ œ— «Ì‰ ò·«” ‘—ò  „Ì ò‰œ ", vbCritical + vbOKOnly, "Œÿ«"
            Exit Sub
            End If
            
            
            
            
            
            
            
'STU2CLASS.Refresh
'STU2CLASS.Recordset.AddNew
'STU2CLASS.Recordset.Fields("parvande") = Label8.Caption
'STU2CLASS.Recordset.Fields("kodclass") = lkodclass.Caption
'STU2CLASS.Recordset.Fields("tshoro") = ttshoro.Text
'STU2CLASS.Recordset.Fields("tpayan") = ttpayan.Text
'STU2CLASS.Recordset.Fields("elat") = telat.Text
'STU2CLASS.Recordset.Fields("tozih") = "ﬁ»·« œ— «Ì‰ ò·«” »ÊœÂ Ê Õ–› ‘œÂ"
'STU2CLASS.Recordset.Fields("d") = Taqvim.Label1.Caption

'STU2CLASS.Recordset.Update
'STU2CLASS.Refresh
        
            
            
            
            
            
            
            
            
            
            
            
                Student.Recordset.Fields("ostad") = lostad.Caption
                
                If Student.Recordset.Fields("clas1") = "‰œ«—œ" Then
                Student.Recordset.Fields("clas1") = lkodclass.Caption
                Else
                If Student.Recordset.Fields("clas2") = "‰œ«—œ" Then
                Student.Recordset.Fields("clas2") = lkodclass.Caption
                Else
                If Student.Recordset.Fields("clas3") = "‰œ«—œ" Then
                Student.Recordset.Fields("clas3") = lkodclass.Caption
                Else
                If Student.Recordset.Fields("clas4") = "‰œ«—œ" Then
                Student.Recordset.Fields("clas4") = lkodclass.Caption
                Else
                If Student.Recordset.Fields("clas5") = "‰œ«—œ" Then
                Student.Recordset.Fields("clas5") = lkodclass.Caption
                Else
                MsgBox "«Ì‰ ﬁ—¬‰ ¬„Ê“ œ— Õ«· Õ«Ÿ— œ— 5 ò·«” »Â ’Ê—  Â„“„«‰ ‘—ò  „Ì ò‰œ Ê „Ã«“ »Â ‘—ò  œ— ò·«” œÌê—Ì ‰„Ì »«‘œ. ·ÿ›« Ìò Ì« ç‰œ ò·«” «Ì‘«‰ —«Õ–› ò‰Ìœ", vbCritical + vbOKOnly, "Œÿ«"
                Exit Sub
                End If
                End If
                End If
                End If
                End If
                STU2CLASS.Recordset.Fields("tpayan") = ttpayan.Text
                STU2CLASS.Recordset.Fields("elat") = telat.Text
               STU2CLASS.Recordset.Fields("tozih") = ttozih.Text & "ﬁ»·« œ— «Ì‰ ò·«” »ÊœÂ Ê Õ–› ‘œÂ «” "
               Student.Recordset.Update
                STU2CLASS.Recordset.Update
                MsgBox "«ÿ·«⁄«  »Â —Ê“ ‘œ", vbInformation + vbOKOnly, "„œ—Ì—  ·Ì”  ò·«”Ì"
                Exit Sub
        Else
            Exit Sub
        End If
End If
Exit Sub
1  'ﬁ»·«  Ê «Ì‰ ò·«” ‰»ÊœÂ


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''







'«Ì‰ òÂ ÂÌÃÌ ‰Ì”  ›ﬁÿ ⁄œœ «” 

X = 1
TC = 0
'«Ì‰ òÂ ÂÌÃÌ ‰Ì” 

'«Ì‰ —« ‰„Ì œ«„ ç—« ‰Ê‘ „ ›ò— ò‰„ «·òÌ «”  Å «” “ ¬‰ òÌ ê“œÌ„


GoTo 19

For I = 1 To 5   ' Å‰Ã ò·«” —« çò „Ì ò‰œ

C = "clas" & X
If Student.Recordset.Fields(C) <> "‰œ«—œ" Then
TC = TC + 1 '‘„«—‘ ò·«” Â«

End If
X = X + 1
Next I
If TC > 1 Then  'ÿ—› »Ì” — «“ 1 ‘—ò  „Ì ò‰œ

If MsgBox("«Ì‰ ﬁ—¬‰ ¬„Ê“ Â„ «ò‰Ê‰ œ— ò·«” œÌê—Ì ‘—ò  „Ì ò‰œ  ¬Ì« „Ì ŒÊ«ÂÌœ ‰«„ «Ì‘«‰ œ— «Ì‰ ò·«” Â„ À»  ‘Êœ", vbQuestion + vbYesNo, " „œÌ—Ì  ·Ì”  ò·«”Ì") = vbYes Then

GoTo 19 '«ÿ„Ì‰«‰ œ«—œ «“ òœ ò·«”


Else  ' «“ ’Õ  òœ ò·«” «ÿ„Ì‰«‰ ‰œ«—œ

Exit Sub
End If
End If

19:















'STU2CLASS.Refresh
'STU2CLASS.RecordSource = " select * from stu2class where  parvande like ('%" + Label8.Caption + "%') "
'STU2CLASS.Refresh
'If STU2CLASS.Recordset.BOF = True Or STU2CLASS.Recordset.EOF = True Then  '¬Ì« ò·«” œÌê— œ«—œ Ì« ‰Â
'GoTo 2 'ò·«” œÌê—Ì ‰œ«—œ
'Else
'If MsgBox("«Ì‰ ﬁ—¬‰ ¬„Ê“ ﬁ»·« œ— ò·«”Ì œÌê— À»  ‘œÂ «”    ¬Ì« „Ì ŒÊ«ÂÌœ œ— ò·«” ÃœÌœ À»  ‘Êœ", vbQuestion + vbYesNo, "„œÌ—Ì  ·Ì”  ò·«”Ì") = vbYes Then

'GoTo 3  'ò·«” œÌê—Ì œ«—œ Ê·Ì „Ì ŒÊ«Âœ œ— ò·«” ÃœÌœ ‘—ò  ò‰œ
'Else
'Exit Sub
'End If


'End If
'2
'3

If MsgBox("  ¬Ì« „Ì ŒÊ«ÂÌœ ¬ﬁ«Ì  " & Label10.Caption & "  œ— ò·«” ‘„«—Â   " & lkodclass.Caption & "   Ê«—œ ‘Êœ   ", vbQuestion + vbYesNo, "„œÌ—Ì  ·Ì”  ﬂ·«”Ì") = vbYes Then



Student.Recordset.Fields("ostad") = lostad.Caption

If Student.Recordset.Fields("clas1") = "‰œ«—œ" Then
Student.Recordset.Fields("clas1") = lkodclass.Caption
Else
If Student.Recordset.Fields("clas2") = "‰œ«—œ" Then
Student.Recordset.Fields("clas2") = lkodclass.Caption
Else
If Student.Recordset.Fields("clas3") = "‰œ«—œ" Then
Student.Recordset.Fields("clas3") = lkodclass.Caption
Else
If Student.Recordset.Fields("clas4") = "‰œ«—œ" Then
Student.Recordset.Fields("clas4") = lkodclass.Caption
Else
If Student.Recordset.Fields("clas5") = "‰œ«—œ" Then
Student.Recordset.Fields("clas5") = lkodclass.Caption
Else
MsgBox "«Ì‰ ﬁ—¬‰ ¬„Ê“ œ— Õ«· Õ«Ÿ— œ— 5 ò·«” »Â ’Ê—  Â„“„«‰ ‘—ò  „Ì ò‰œ Ê „Ã«“ »Â ‘—ò  œ— ò·«” œÌê—Ì ‰„Ì »«‘œ. ·ÿ›« Ìò Ì« ç‰œ ò·«” «Ì‘«‰ —«Õ–› ò‰Ìœ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
End If
End If
End If
End If
End If



STU2CLASS.Refresh
STU2CLASS.Recordset.AddNew
STU2CLASS.Recordset.Fields("parvande") = Label8.Caption
STU2CLASS.Recordset.Fields("kodclass") = lkodclass.Caption
STU2CLASS.Recordset.Fields("tshoro") = ttshoro.Text
'STU2CLASS.Recordset.Fields("tpayan") = ttpayan.Text
'STU2CLASS.Recordset.Fields("elat") = telat.Text
STU2CLASS.Recordset.Fields("tozih") = ttozih.Text
STU2CLASS.Recordset.Fields("d") = Taqvim.Tarikh.Caption

STU2CLASS.Recordset.Update
STU2CLASS.Refresh




Student.Recordset.Update
MsgBox "«ÿ·«⁄«  œ— ”Ì” „ À»  ‘œ", vbInformation + vbOKOnly, "„œÌ—Ì  ·Ì”  ﬂ·«”Ì"

Else
Student.Recordset.CancelUpdate

End If


Text1.Text = ""
Text1.SetFocus






' ÂÌÂ ·Ì”  »«ò” »—«Ì ò·«”
Student.Refresh
Student.RecordSource = "select * from student where clas1 like ('%" + lkodclass.Caption + "%') or clas2 like ('%" + lkodclass.Caption + "%') or clas3 like ('%" + lkodclass.Caption + "%') or clas4 like ('%" + lkodclass.Caption + "%') or clas5 like ('%" + lkodclass.Caption + "%')"
Student.Refresh
List1.Clear
Label33.Caption = Student.Recordset.RecordCount

For I = 1 To Student.Recordset.RecordCount
List1.AddItem (Student.Recordset.Fields("parvande") & "  -  " & Student.Recordset.Fields("name") & "  -  " & Student.Recordset.Fields("famil"))
Student.Recordset.MoveNext
Next I




End Sub

Private Sub Command10_Click()


Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim AD As String
Student.Recordset.MoveFirst
On Error GoTo 1
GoTo 2

1: MsgBox "›«Ì· ÅÌœ« ‰‘œ∫ »—«Ì «Ã—«Ì «Ì‰ ›—„«‰ »«Ìœ «“ ”Ì” „ ”—Ê— «” ›«œÂ ò‰Ìœ", vbCritical, "Œÿ«"

Exit Sub

2:


' »—«Ì «Ì‰òÂ ﬁ« Ì ‰‘Â »«Ìœ ò«„·« Ãœ« »«‘œÂ
'·Ì”  «⁄·«‰ A4
'A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A44A4A4A4A4A4A44A4A44A4A4A44A









Set oExcel = GetObject(Entekhab.AdressLabel.Caption & "QeybatClass.xlsx")
oExcel.ActiveSheet.Range("b3").Value = lkodclass.Caption
oExcel.ActiveSheet.Range("d3").Value = ltarh.Caption
oExcel.ActiveSheet.Range("f3").Value = lmaqta.Caption
oExcel.ActiveSheet.Range("h3").Value = lzsho.Caption & " «·Ì  " & lzpa.Caption
oExcel.ActiveSheet.Range("b4").Value = lostad.Caption
oExcel.ActiveSheet.Range("E4").Value = ltsho.Caption
oExcel.ActiveSheet.Range("f4").Value = ltpa.Caption

oExcel.ActiveSheet.Range("h4").Value = lmadras.Caption
oExcel.ActiveSheet.Range("b5").Value = mclass.Recordset.Fields("tozih")


STU2CLASS.Refresh
STU2CLASS.RecordSource = "select * from STU2CLASS where kodclass like ('%" + lkodclass.Caption + "%') and elat like ('%" + "€Ì» " + "%') "
STU2CLASS.Refresh


Dim NumberOfRows As Integer
NumberOfRows = STU2CLASS.Recordset.RecordCount
For r = 8 To NumberOfRows + 7







Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + STU2CLASS.Recordset.Fields("parvande") + "%') "
Student.Refresh


oExcel.ActiveSheet.Range("b" & r).Value = Student.Recordset.Fields("PARVANDE")
oExcel.ActiveSheet.Range("c" & r).Value = Student.Recordset.Fields("NAME")
oExcel.ActiveSheet.Range("d" & r).Value = Student.Recordset.Fields("FAMIL")
oExcel.ActiveSheet.Range("g" & r).Value = Student.Recordset.Fields("tell") & "-" & Student.Recordset.Fields("mob")

oExcel.ActiveSheet.Range("e" & r).Value = STU2CLASS.Recordset.Fields("elat")
oExcel.ActiveSheet.Range("f" & r).Value = STU2CLASS.Recordset.Fields("tpayan")
oExcel.ActiveSheet.Range("h" & r).Value = STU2CLASS.Recordset.Fields("tozih")


STU2CLASS.Recordset.MoveNext
Next

MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  Ê«—œ »—‰«„Â «ò”· ‘œ‰œ", 64, "À»  «ÿ·«⁄« "
AD = lkodclass.Caption & "Õ–› »Â œ·Ì· €Ì» "
oExcel.SaveAs AD
oExcel.Application.Visible = True
oExcel.Parent.Windows(1).Visible = True



'A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A44A4A4A4A4A4A44A4A44A4A4A44A

End Sub

Private Sub Command11_Click()




Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim AD As String
Student.Recordset.MoveFirst
On Error GoTo 1
GoTo 2

1: MsgBox "›«Ì· ÅÌœ« ‰‘œ∫ »—«Ì «Ã—«Ì «Ì‰ ›—„«‰ »«Ìœ «“ ”Ì” „ ”—Ê— «” ›«œÂ ò‰Ìœ", vbCritical, "Œÿ«"

Exit Sub

2:


' »—«Ì «Ì‰òÂ ﬁ« Ì ‰‘Â »«Ìœ ò«„·« Ãœ« »«‘œÂ
'·Ì”  «⁄·«‰ A4
'A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A44A4A4A4A4A4A44A4A44A4A4A44A









Set oExcel = GetObject(Entekhab.AdressLabel.Caption & "allclass.xlsx")
oExcel.ActiveSheet.Range("b3").Value = lkodclass.Caption
oExcel.ActiveSheet.Range("d3").Value = ltarh.Caption
oExcel.ActiveSheet.Range("f3").Value = lmaqta.Caption
oExcel.ActiveSheet.Range("h3").Value = lzsho.Caption & " «·Ì  " & lzpa.Caption
oExcel.ActiveSheet.Range("b4").Value = lostad.Caption
oExcel.ActiveSheet.Range("E4").Value = ltsho.Caption
oExcel.ActiveSheet.Range("f4").Value = ltpa.Caption

oExcel.ActiveSheet.Range("h4").Value = lmadras.Caption
oExcel.ActiveSheet.Range("b5").Value = mclass.Recordset.Fields("tozih")


STU2CLASS.Refresh
STU2CLASS.RecordSource = "select * from STU2CLASS where kodclass like ('%" + lkodclass.Caption + "%')  "
STU2CLASS.Refresh


Dim NumberOfRows As Integer
NumberOfRows = STU2CLASS.Recordset.RecordCount
For r = 8 To NumberOfRows + 7







Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + STU2CLASS.Recordset.Fields("parvande") + "%') "
Student.Refresh


oExcel.ActiveSheet.Range("b" & r).Value = Student.Recordset.Fields("PARVANDE")
oExcel.ActiveSheet.Range("c" & r).Value = Student.Recordset.Fields("NAME")
oExcel.ActiveSheet.Range("d" & r).Value = Student.Recordset.Fields("FAMIL")
oExcel.ActiveSheet.Range("g" & r).Value = Student.Recordset.Fields("tell") & "-" & Student.Recordset.Fields("mob")

oExcel.ActiveSheet.Range("e" & r).Value = STU2CLASS.Recordset.Fields("elat")
oExcel.ActiveSheet.Range("f" & r).Value = STU2CLASS.Recordset.Fields("tpayan")
oExcel.ActiveSheet.Range("h" & r).Value = STU2CLASS.Recordset.Fields("tozih")


STU2CLASS.Recordset.MoveNext
Next

MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  Ê«—œ »—‰«„Â «ò”· ‘œ‰œ", 64, "À»  «ÿ·«⁄« "
AD = lkodclass.Caption & "·Ì”  ò«„· ò·«”"
oExcel.SaveAs AD
oExcel.Application.Visible = True
oExcel.Parent.Windows(1).Visible = True



'A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A44A4A4A4A4A4A44A4A44A4A4A44A



End Sub

Private Sub Command12_Click()


ProgressBar1.Visible = True
ProgressBar1.Value = 0
List1.Clear
List2.Clear
List3.Clear
List4.Clear

On Error Resume Next

ProgressBar1.Value = 1

'ﬁ—¬ ‰¬„Ê“«‰ Õ«÷—
Student.Refresh
Student.RecordSource = "select * from student where clas1 like ('%" + lkodclass.Caption + "%') or clas2 like ('%" + lkodclass.Caption + "%') or clas3 like ('%" + lkodclass.Caption + "%') or clas4 like ('%" + lkodclass.Caption + "%') or clas5 like ('%" + lkodclass.Caption + "%')"
Student.Refresh
List1.Clear
Label33.Caption = Student.Recordset.RecordCount

For I = 1 To Student.Recordset.RecordCount
List1.AddItem (Student.Recordset.Fields("parvande") & "  -  " & Student.Recordset.Fields("name") & "  -  " & Student.Recordset.Fields("famil"))
Student.Recordset.MoveNext
Next I



GoTo 1

Exit Sub

ProgressBar1.Value = 2
'« „«„
STU2CLASS.Refresh
STU2CLASS.RecordSource = "select * from STU2CLASS where kodclass like ('%" + lkodclass.Caption + "%') and elat like ('%" + "« „«„" + "%') "
STU2CLASS.Refresh
Label31.Caption = STU2CLASS.Recordset.RecordCount

For I = 1 To STU2CLASS.Recordset.RecordCount

Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + STU2CLASS.Recordset.Fields("parvande") + "%') "
Student.Refresh

List2.AddItem (Student.Recordset.Fields("parvande") & "  -  " & Student.Recordset.Fields("name") & "  -  " & Student.Recordset.Fields("famil"))
STU2CLASS.Recordset.MoveNext
Next I

ProgressBar1.Value = 3
' €Ì» 
STU2CLASS.Refresh
STU2CLASS.RecordSource = "select * from STU2CLASS where kodclass like ('%" + lkodclass.Caption + "%') and elat like ('%" + "€Ì» " + "%') "
STU2CLASS.Refresh
Label29.Caption = STU2CLASS.Recordset.RecordCount

For I = 1 To STU2CLASS.Recordset.RecordCount

Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + STU2CLASS.Recordset.Fields("parvande") + "%') "
Student.Refresh

List3.AddItem (Student.Recordset.Fields("parvande") & "  -  " & Student.Recordset.Fields("name") & "  -  " & Student.Recordset.Fields("famil"))
STU2CLASS.Recordset.MoveNext
Next I

'ò«„·
ProgressBar1.Value = 4

STU2CLASS.Refresh
STU2CLASS.RecordSource = "select * from STU2CLASS where kodclass like ('%" + lkodclass.Caption + "%') "
STU2CLASS.Refresh

Label37.Caption = STU2CLASS.Recordset.RecordCount

For I = 1 To STU2CLASS.Recordset.RecordCount

Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + STU2CLASS.Recordset.Fields("parvande") + "%') "
Student.Refresh

List4.AddItem (Student.Recordset.Fields("parvande") & "  -  " & Student.Recordset.Fields("name") & "  -  " & Student.Recordset.Fields("famil"))
STU2CLASS.Recordset.MoveNext
Next I

1
ProgressBar1.Value = 5
ProgressBar1.Visible = False

End Sub

Private Sub Command2_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "fclass-delete" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513

        If ttpayan.Text = "" Or telat.Text = "" Then
            MsgBox "·ÿ›«  «—ÌŒ Ê ⁄·  Õ–› —« œ—Ã ò‰Ìœ", vbInformation + vbOKOnly, "Õ–› ﬁ—¬‰ ¬„Ê“"
        Else




            STU2CLASS.Refresh
            STU2CLASS.RecordSource = " select * from stu2class where  parvande like ('%" + Label8.Caption + "%') and kodclass like ('%" + lkodclass.Caption + "%')"
            STU2CLASS.Refresh
            If STU2CLASS.Recordset.BOF = True Or STU2CLASS.Recordset.EOF = True Then
                MsgBox "‰«„ «Ì‰ ﬁ—¬‰ ¬„Ê“ œ— «Ì‰ ò·«” À»  ‰‘œÂ «”  ", vbCritical + vbOKOnly, "Œÿ«"
                Exit Sub

            Else
                Student.RecordSource = "select * from student where parvande like ('%" + Label8.Caption + "%')"
                Student.Refresh
               

                If MsgBox("  ¬Ì« „Ì ŒÊ«ÂÌœ ¬ﬁ«Ì  " & Label10.Caption & "  «“ ·Ì”  ò·«”Ì œ«—«Ì òœ    " & lkodclass.Caption & "  Õ–› ò‰Ìœ", vbQuestion + vbYesNo, "Õ–› ò‰Ìœ") = vbYes Then

                    '??? ?? ????? ?? ????? ???? ??? ???
                    If Student.Recordset.Fields("clas1") = lkodclass.Caption Then
                        Student.Recordset.Fields("clas1") = "‰œ«—œ"
                    Else
                        If Student.Recordset.Fields("clas2") = lkodclass.Caption Then
                            Student.Recordset.Fields("clas2") = "‰œ«—œ"
                        Else
                            If Student.Recordset.Fields("clas3") = lkodclass.Caption Then
                                Student.Recordset.Fields("clas3") = "‰œ«—œ"
                            Else
                                If Student.Recordset.Fields("clas4") = lkodclass.Caption Then
                                    Student.Recordset.Fields("clas4") = "‰œ«—œ"
                                Else
                                    If Student.Recordset.Fields("clas5") = lkodclass.Caption Then
                                        Student.Recordset.Fields("clas5") = "‰œ«—œ"
                                    Else
                                        MsgBox "ﬁ—¬‰ ¬„Ê“ ﬁ»·« «“ ò·«” Õ–› ‘œÂ «” ", vbExclamation + vbOKOnly, "Õ–› ﬁ—¬‰ ¬„Ê“"
                                        Exit Sub

                                    End If
                                End If
                            End If
                        End If
                    End If
                Else  ' ¬Ì« „Ì ŒÊ«ÂÌœ ﬁ—¬‰ ¬„Ê“ «— ·Ì” ò·«” Õ–› ‘Êœ   ‰Â ‰„Ì ŒÊ«ÂÌ„ Õ–› ‘Êœ
                Exit Sub
                
                
                End If





                STU2CLASS.Refresh
                STU2CLASS.RecordSource = "select * from stu2class where parvande like ('%" + Label8.Caption + "%') and kodclass like ('%" + lkodclass.Caption + "%')"
                STU2CLASS.Refresh


                STU2CLASS.Recordset.Fields("tpayan") = ttpayan.Text
                STU2CLASS.Recordset.Fields("elat") = telat.Text
                STU2CLASS.Recordset.Fields("tozih") = ttozih.Text
                Student.Recordset.Update
                STU2CLASS.Recordset.Update

                MsgBox "ﬁ—¬‰ ¬„Ê“ Õ–› ‘œ", vbInformation + vbOKOnly, "Õ–› ﬁ—¬‰ ¬„Ê“"
     

           End If
      
            
        End If '·Ì›«  «—ÌŒ Ê ⁄·Ì  Õ–› —« »Ì«‰ òÌ‰œ 
        
        
        
        
        Student.Refresh
        Student.RecordSource = "select * from student where clas1 like ('%" + lkodclass.Caption + "%') or clas2 like ('%" + lkodclass.Caption + "%') or clas3 like ('%" + lkodclass.Caption + "%') or clas4 like ('%" + lkodclass.Caption + "%') or clas5 like ('%" + lkodclass.Caption + "%')"
        Student.Refresh
        List1.Clear
        Label33.Caption = Student.Recordset.RecordCount

        For I = 1 To Student.Recordset.RecordCount
        List1.AddItem (Student.Recordset.Fields("parvande") & "  -  " & Student.Recordset.Fields("name") & "  -  " & Student.Recordset.Fields("famil"))
        Student.Recordset.MoveNext
        Next I




' ÂÌÂ ·Ì”  »«ò” »—«Ì ò·«”
Student.Refresh
Student.RecordSource = "select * from student where clas1 like ('%" + lkodclass.Caption + "%') or clas2 like ('%" + lkodclass.Caption + "%') or clas3 like ('%" + lkodclass.Caption + "%') or clas4 like ('%" + lkodclass.Caption + "%') or clas5 like ('%" + lkodclass.Caption + "%')"
Student.Refresh
List1.Clear
Label33.Caption = Student.Recordset.RecordCount

For I = 1 To Student.Recordset.RecordCount
List1.AddItem (Student.Recordset.Fields("parvande") & "  -  " & Student.Recordset.Fields("name") & "  -  " & Student.Recordset.Fields("famil"))
Student.Recordset.MoveNext
Next I




        
End Sub









Private Sub Command4_Click()
Call Text1_Change


End Sub

Private Sub Command5_Click()
Beep


'Exit Sub



If MsgBox(" „«„ ﬁ—¬‰ ¬„Ê“«‰Ì òÂ œ— «Ì‰ ò·«” ‘—ò  „Ì ò‰‰œ «“ ·Ì”  ò·«”Ì Õ–› ‘œÂ Ê «Ì‰ ò·«” œÌê— Å–Ì—«Ì ﬁ—¬‰ ¬„Ê“ ‰„Ì »«‘œ   ¬Ì« „ÿ„∆‰ Â” Ìœ", vbExclamation + vbYesNo, "Â‘œ«—") = vbYes Then
GoTo 11
End If
Exit Sub
11:


            STU2CLASS.Refresh
            STU2CLASS.RecordSource = " select * from stu2class where kodclass like ('%" + lkodclass.Caption + "%')"
            STU2CLASS.Refresh
            If STU2CLASS.Recordset.BOF = True Or STU2CLASS.Recordset.EOF = True Then
                MsgBox "‰«„ ÂÌÃ ﬁ—¬‰ ¬„Ê“Ì œ— «Ì‰ ò·«” À»  ‰‘œÂ «” ", vbCritical + vbOKOnly, "Œÿ«"
                Exit Sub

            Else
            'œ— «Ì‰Ã«  „«„ ò”«‰ÌòÂ œ— «Ì‰ ò·«” Â” ‰œ —« „Ì “‰œ « „«„ ò·«”
            '
Student.Refresh
Student.RecordSource = "select * from student where clas1 like ('%" + lkodclass.Caption + "%') or clas2 like ('%" + lkodclass.Caption + "%') or clas3 like ('%" + lkodclass.Caption + "%') or clas4 like ('%" + lkodclass.Caption + "%') or clas5 like ('%" + lkodclass.Caption + "%')"
Student.Refresh
            
       For I = 1 To Student.Recordset.RecordCount
            STU2CLASS.Refresh
            STU2CLASS.RecordSource = " select * from stu2class where parvande like ('%" + Student.Recordset.Fields("parvande") + "%') and kodclass like ('%" + lkodclass.Caption + "%')"
            STU2CLASS.Refresh
            
            STU2CLASS.Recordset.Fields("elat") = "« „«„ ò·«”"
            STU2CLASS.Recordset.Fields("tpayan") = Taqvim.Tarikh.Caption
            STU2CLASS.Recordset.Update
        STU2CLASS.Refresh

Student.Recordset.MoveNext




         Next I
              
            
            'Å«Ì«‰ À»  « „«„ ò·«” »—«Ì  „«„ ﬁ—¬‰ ¬„“Ê«‰
            
            '
            Student.Refresh
            Student.RecordSource = "select * from student where clas1 like ('%" + lkodclass.Caption + "%') or clas2 like ('%" + lkodclass.Caption + "%') or clas3 like ('%" + lkodclass.Caption + "%') or clas4 like ('%" + lkodclass.Caption + "%') or clas5 like ('%" + lkodclass.Caption + "%')"
            Student.Refresh
            
            ' »«Ìœ Â— ò”Ì œ— «Ì‰ ò·«” »ÊœÂ  Õ–› ò‰œ «“ «‰ÌÃ« ‘—Ê⁄ „Ì ‘Êœ
            For I = 1 To Student.Recordset.RecordCount
            



                    
                    If Student.Recordset.Fields("clas1") = lkodclass.Caption Then
                        Student.Recordset.Fields("clas1") = "‰œ«—œ"
                    Else
                        If Student.Recordset.Fields("clas2") = lkodclass.Caption Then
                            Student.Recordset.Fields("clas2") = "‰œ«—œ"
                        Else
                            If Student.Recordset.Fields("clas3") = lkodclass.Caption Then
                                Student.Recordset.Fields("clas3") = "‰œ«—œ"
                            Else
                                If Student.Recordset.Fields("clas4") = lkodclass.Caption Then
                                    Student.Recordset.Fields("clas4") = "‰œ«—œ"
                                Else
                                    If Student.Recordset.Fields("clas5") = lkodclass.Caption Then
                                        Student.Recordset.Fields("clas5") = "‰œ«—œ"
                                   
                                       

                                    End If
                                End If
                            End If
                        End If
                    End If
                
                
          Student.Recordset.Update
          Student.Recordset.MoveNext
          


        Next I
        
        
   mclass.Refresh
   mclass.RecordSource = "select * from mclass where kodclass like ('%" + lkodclass.Caption + "%')"
   mclass.Refresh
   mclass.Recordset.Fields("tpayan") = Taqvim.Tarikh.Caption
   mclass.Recordset.Fields("tozih") = "« „«„ ò·«”"
   
        mclass.Recordset.Update
        
         mclass.Refresh
        
        
        
        
        
        
        
MsgBox "⁄„·Ì«  À»  « „«„ ò·«” »« „Ê›ﬁÌ  «‰Ã«„ ‘œ", vbInformation + vbOKOnly, "«À»  « „«„ ò·«”"
        
        
End If ' ‰«„ ÂÌç ﬁ—¬‰ ¬„Ê“Ì œ — «Ì‰ ò·«”Ì À»  ‰‘œÂ «” 







 
End Sub

Private Sub Command6_Click()
List1.Clear
List2.Clear
List3.Clear
List4.Clear

End Sub

Private Sub Command7_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "fclass-print" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub

14082513

Dim oExcel As Object
'Dim oBook As Object
'Dim oSheet As Object
Dim AD As String
'Student.Recordset.MoveFirst
'On Error GoTo 1
GoTo 2

1: MsgBox "›«Ì· ÅÌœ« ‰‘œ∫ »—«Ì «Ã—«Ì «Ì‰ ›—„«‰ »«Ìœ «“ ”Ì” „ ”—Ê— «” ›«œÂ ò‰Ìœ", vbCritical, "Œÿ«"

Exit Sub

2:


' »—«Ì «Ì‰òÂ ﬁ« Ì ‰‘Â »«Ìœ ò«„·« Ãœ« »«‘œÂ
'·Ì”  «⁄·«‰ A4
'A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A44A4A4A4A4A4A44A4A44A4A4A44A

If elana4.Checked = True Then   ' ‘—ÿ «’·Ì »—«Ì «‰ Œ«» ·Ì” 


Student.Refresh
Student.RecordSource = "select * from student where clas1 like ('%" + lkodclass.Caption + "%') or clas2 like ('%" + lkodclass.Caption + "%') or clas3 like ('%" + lkodclass.Caption + "%') or clas4 like ('%" + lkodclass.Caption + "%') or clas5 like ('%" + lkodclass.Caption + "%')"
Student.Refresh

If Entekhab.Pc.Checked = True Then
Set oExcel = GetObject(Entekhab.AdressLabel.Caption & "ElanA4.xlsx")

End If

If Entekhab.net.Checked = True Then
Set oExcel = GetObject(Entekhab.NetAdresslabel.Caption & "ElanA4.xlsx")
End If







oExcel.ActiveSheet.Range("A3").Value = "·Ì”  ò·«” " & ltarh.Caption & "(" & lmaqta.Caption & ")"

oExcel.ActiveSheet.Range("G5").Value = lostad.Caption
oExcel.ActiveSheet.Range("G6").Value = lzsho.Caption
oExcel.ActiveSheet.Range("I6").Value = lzpa.Caption

oExcel.ActiveSheet.Range("G7").Value = lmadras.Caption

oExcel.ActiveSheet.Range("E8").Value = ltsho.Caption
oExcel.ActiveSheet.Range("J8").Value = lkodclass.Caption


Dim NumberOfRows As Integer
NumberOfRows = Student.Recordset.RecordCount
For r = 11 To NumberOfRows + 10
oExcel.ActiveSheet.Range("C" & r).Value = Student.Recordset.Fields("PARVANDE")
oExcel.ActiveSheet.Range("D" & r).Value = Student.Recordset.Fields("NAME")
oExcel.ActiveSheet.Range("G" & r).Value = Student.Recordset.Fields("FAMIL")
'oExcel.ActiveSheet.Range("X" & r).Value = Student.Recordset.Fields("tell") & "-" & Student.Recordset.Fields("mob")
Student.Recordset.MoveNext
Next

MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  Ê«—œ »—‰«„Â «ò”· ‘œ‰œ", 64, "À»  «ÿ·«⁄« "
AD = lkodclass.Caption & "  ·Ì”  «⁄·«‰ a4"
'oExcel.SaveAs AD
'oExcel.Application.Visible = True
'oExcel.Parent.Windows(1).Visible = True
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
oExcel.Application.Visible = True
On Error GoTo 722


oExcel.Parent.Windows(2).Visible = True
GoTo 910
722:

oExcel.Parent.Windows(1).Visible = True
910:
''''''

oExcel.SaveAs AD
'oExcel.Close
'
'
'Set oExcel = Nothing ' Remove object variable.
''''''''
'Shell "Explorer.exe " & "c:\" & KodEnhesariPrint & ".xlsx"

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End If

'A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A44A4A4A4A4A4A44A4A44A4A4A44A




'A5A5A5A5A5A5A5A5A5A5A5A55A5A55A5A5A55A5A55A5A5A5A5A5A55A5A5A5A5A5A5A55A5A5AA
If elana5.Checked = True Then




Student.Refresh
Student.RecordSource = "select * from student where clas1 like ('%" + lkodclass.Caption + "%') or clas2 like ('%" + lkodclass.Caption + "%') or clas3 like ('%" + lkodclass.Caption + "%') or clas4 like ('%" + lkodclass.Caption + "%') or clas5 like ('%" + lkodclass.Caption + "%')"
Student.Refresh


If Entekhab.Pc.Checked = True Then
Set oExcel = GetObject(Entekhab.AdressLabel.Caption & "ElanA5.xlsx")

End If

If Entekhab.net.Checked = True Then
Set oExcel = GetObject(Entekhab.NetAdresslabel.Caption & "ElanA5.xlsx")
End If






'Set oExcel = GetObject("F:\Markaz Quran & Hadis\FORMXLS\ElanA5.xlsx")

oExcel.ActiveSheet.Range("A3").Value = "·Ì”  ò·«” " & ltarh.Caption & "(" & lmaqta.Caption & ")"

oExcel.ActiveSheet.Range("G5").Value = lostad.Caption
oExcel.ActiveSheet.Range("G6").Value = lzsho.Caption
oExcel.ActiveSheet.Range("I6").Value = lzpa.Caption

oExcel.ActiveSheet.Range("G7").Value = lmadras.Caption

oExcel.ActiveSheet.Range("E8").Value = ltsho.Caption
oExcel.ActiveSheet.Range("J8").Value = lkodclass.Caption


'Dim NumberOfRows As Integer
NumberOfRows = Student.Recordset.RecordCount
For r = 11 To NumberOfRows + 10
oExcel.ActiveSheet.Range("C" & r).Value = Student.Recordset.Fields("PARVANDE")
oExcel.ActiveSheet.Range("D" & r).Value = Student.Recordset.Fields("NAME")
oExcel.ActiveSheet.Range("G" & r).Value = Student.Recordset.Fields("FAMIL")
'oExcel.ActiveSheet.Range("X" & r).Value = Student.Recordset.Fields("tell") & "-" & Student.Recordset.Fields("mob")
Student.Recordset.MoveNext
Next

MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  Ê«—œ »—‰«„Â «ò”· ‘œ‰œ", 64, "À»  «ÿ·«⁄« "
AD = lkodclass.Caption & "  ·Ì”  «⁄·«‰ a5"
'oExcel.SaveAs AD
'oExcel.Application.Visible = True
'oExcel.Parent.Windows(1).Visible = True


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
oExcel.Application.Visible = True
On Error GoTo 7222


oExcel.Parent.Windows(2).Visible = True
GoTo 9102
7222

oExcel.Parent.Windows(1).Visible = True
9102
''''''

oExcel.SaveAs AD
'oExcel.Close
'
'
'Set oExcel = Nothing ' Remove object variable.
''''''''
'Shell "Explorer.exe " & "c:\" & KodEnhesariPrint & ".xlsx"

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<



End If

'A5A5A5A5A5A5A5A5A5A5A5A55A5A55A5A5A55A5A55A5A5A5A5A5A55A5A5A5A5A5A5A55A5A5AA


'ò «» ò «» ò «» ò «» ò «» ò «» ò «» ò »« ò «» ò «» ò «» ò «» ò «» «» ò «» ò «» ò «» ò »« ò «» ò «» ò «» ò «» ò «» «» ò «» ò «» ò «» ò »« ò «» ò «» ò «» ò «» ò «» ê«
If foroshlist.Checked = True Then



Student.Refresh
Student.RecordSource = "select * from student where clas1 like ('%" + lkodclass.Caption + "%') or clas2 like ('%" + lkodclass.Caption + "%') or clas3 like ('%" + lkodclass.Caption + "%') or clas4 like ('%" + lkodclass.Caption + "%') or clas5 like ('%" + lkodclass.Caption + "%')"
Student.Refresh


If Entekhab.Pc.Checked = True Then
Set oExcel = GetObject(Entekhab.AdressLabel.Caption & "foroshlist.xlsx")

End If

If Entekhab.net.Checked = True Then
Set oExcel = GetObject(Entekhab.NetAdresslabel.Caption & "foroshlist.xlsx")
End If






'Set oExcel = GetObject("F:\Markaz Quran & Hadis\FORMXLS\foroshlist.xlsx")

oExcel.ActiveSheet.Range("A3").Value = "·Ì”  ò·«” " & ltarh.Caption & "(" & lmaqta.Caption & ")"

oExcel.ActiveSheet.Range("G5").Value = lostad.Caption
oExcel.ActiveSheet.Range("G6").Value = lzsho.Caption
oExcel.ActiveSheet.Range("I6").Value = lzpa.Caption

oExcel.ActiveSheet.Range("G7").Value = lmadras.Caption

oExcel.ActiveSheet.Range("E8").Value = ltsho.Caption
oExcel.ActiveSheet.Range("J8").Value = lkodclass.Caption


'Dim NumberOfRows As Integer
NumberOfRows = Student.Recordset.RecordCount
For r = 11 To NumberOfRows + 10
oExcel.ActiveSheet.Range("C" & r).Value = Student.Recordset.Fields("PARVANDE")
oExcel.ActiveSheet.Range("D" & r).Value = Student.Recordset.Fields("NAME")
oExcel.ActiveSheet.Range("G" & r).Value = Student.Recordset.Fields("FAMIL")
'oExcel.ActiveSheet.Range("X" & r).Value = Student.Recordset.Fields("tell") & "-" & Student.Recordset.Fields("mob")
Student.Recordset.MoveNext
Next

MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  Ê«—œ »—‰«„Â «ò”· ‘œ‰œ", 64, "À»  «ÿ·«⁄« "
AD = lkodclass.Caption & " ·Ì”  ›—Ê‘"
'oExcel.SaveAs AD
'oExcel.Application.Visible = True
'oExcel.Parent.Windows(1).Visible = True


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
oExcel.Application.Visible = True
On Error GoTo 7224


oExcel.Parent.Windows(2).Visible = True
GoTo 9104
7224:

oExcel.Parent.Windows(1).Visible = True
9104:
''''''

oExcel.SaveAs AD
'oExcel.Close
'
'
'Set oExcel = Nothing ' Remove object variable.
''''''''
'Shell "Explorer.exe " & "c:\" & KodEnhesariPrint & ".xlsx"

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End If

'ò «» ò «» ò «» ò «» ò «» ò «» ò «» ò »« ò «» ò «» ò «» ò «» ò «» «» ò «» ò «» ò «» ò »« ò «» ò «» ò «» ò «» ò «» «» ò «» ò «» ò «» ò »« ò «» ò «» ò «


'ò·«”Ì ò·«”Ìò·«”Ì ò·«”Ìò·«”Ì ò·«”Ìò·«”Ì ò·«”Ìò·«”Ì ò·«”Ìò·«”Ì ò·«”Ìò·«”Ì ò·«”Ìò·«”Ì ò·«”Ìò·«”Ì ò·«”Ìò·«”Ì ò·«”Ìò·«”Ì ò·«”Ìò·«”Ì ò·«”Ìò·«”Ì ò·«”Ì

If lclasslist.Checked = True Then

'»«Ìœ €Ì  Â«—« Ê«—œ «ò”· ò‰œ
Student.Refresh
Student.RecordSource = "select * from student where clas1 like ('%" + lkodclass.Caption + "%') or clas2 like ('%" + lkodclass.Caption + "%') or clas3 like ('%" + lkodclass.Caption + "%') or clas4 like ('%" + lkodclass.Caption + "%') or clas5 like ('%" + lkodclass.Caption + "%')"
Student.Refresh



'Dim oExcel As Object
'Dim oBook As Object
'Dim oSheet As Object
'Dim AD As String
Student.Recordset.MoveFirst
'On Error GoTo 1
'GoTo 2

'1: MsgBox "›«Ì· ÅÌœ« ‰‘œ∫ »—«Ì «Ã—«Ì «Ì‰ ›—„«‰ »«Ìœ «“ ”Ì” „ ”—Ê— «” ›«œÂ ò‰Ìœ", vbCritical, "Œÿ«"

'Exit Sub

'2:
If Entekhab.Pc.Checked = True Then
Set oExcel = GetObject(Entekhab.AdressLabel.Caption & "lclas-jadid.xlsx")

End If

If Entekhab.net.Checked = True Then
Set oExcel = GetObject(Entekhab.NetAdresslabel.Caption & "lclas-jadid.xlsx")
End If



'Set oExcel = GetObject("F:\Markaz Quran & Hadis\FORMXLS\lclas-jadid.xlsx")
oExcel.ActiveSheet.Range("C2").Value = ltarh.Caption
oExcel.ActiveSheet.Range("G2").Value = lmaqta.Caption
oExcel.ActiveSheet.Range("M2").Value = lostad
oExcel.ActiveSheet.Range("AC2").Value = lmadras.Caption
oExcel.ActiveSheet.Range("R1").Value = ltsho.Caption
oExcel.ActiveSheet.Range("X1").Value = lkodclass.Caption
oExcel.ActiveSheet.Range("T2").Value = lzsho.Caption
oExcel.ActiveSheet.Range("V2").Value = lzpa.Caption
'Dim NumberOfRows As Integer
NumberOfRows = Student.Recordset.RecordCount
For r = 6 To NumberOfRows + 5
oExcel.ActiveSheet.Range("B" & r).Value = Student.Recordset.Fields("NAME")
oExcel.ActiveSheet.Range("C" & r).Value = Student.Recordset.Fields("FAMIL")
oExcel.ActiveSheet.Range("X" & r).Value = Student.Recordset.Fields("tell") & "-" & Student.Recordset.Fields("mob")
Student.Recordset.MoveNext
Next

MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  Ê«—œ »—‰«„Â «ò”· ‘œ‰œ", 64, "À»  «ÿ·«⁄« "
AD = lkodclass.Caption
'oExcel.SaveAs AD
'oExcel.Application.Visible = True
'oExcel.Parent.Windows(1).Visible = True
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
oExcel.Application.Visible = True
On Error GoTo 7225


oExcel.Parent.Windows(2).Visible = True
GoTo 9105
7225:

oExcel.Parent.Windows(1).Visible = True
9105:
''''''

oExcel.SaveAs AD
'oExcel.Close
'
'
'Set oExcel = Nothing ' Remove object variable.
''''''''
'Shell "Explorer.exe " & "c:\" & KodEnhesariPrint & ".xlsx"

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End If



'·Ì”  ò·«” ÃœÌœ ·Ì”  ò·«” ÃœÌœ ·Ì ” ò·«” ÃœÌœ ·Ì”  ò·«” ÃœÌœ


If lclasslist1.Checked = True Then

'»«Ìœ €Ì  Â«—« Ê«—œ «ò”· ò‰œ
Student.Refresh
Student.RecordSource = "select * from student where clas1 like ('%" + lkodclass.Caption + "%') or clas2 like ('%" + lkodclass.Caption + "%') or clas3 like ('%" + lkodclass.Caption + "%') or clas4 like ('%" + lkodclass.Caption + "%') or clas5 like ('%" + lkodclass.Caption + "%')"
Student.Refresh



'Dim oExcel As Object
'Dim oBook As Object
'Dim oSheet As Object
'Dim AD As String
'Student.Recordset.MoveFirst
On Error GoTo 9898
GoTo 9999

9898:
MsgBox "›«Ì· ÅÌœ« ‰‘œ∫ »—«Ì «Ã—«Ì «Ì‰ ›—„«‰ »«Ìœ «“ ”Ì” „ ”—Ê— «” ›«œÂ ò‰Ìœ", vbCritical, "Œÿ«"

Exit Sub

9999:

If Entekhab.Pc.Checked = True Then
Set oExcel = GetObject(Entekhab.AdressLabel.Caption & "ListclassJadid.xlsx")

End If

If Entekhab.net.Checked = True Then
Set oExcel = GetObject(Entekhab.NetAdresslabel.Caption & "ListclassJadid.xlsx")
End If



'Set oExcel = GetObject("F:\Markaz Quran & Hadis\FORMXLS\ListclassJadid.xlsx")
oExcel.ActiveSheet.Range("C2").Value = ltarh.Caption
oExcel.ActiveSheet.Range("f2").Value = lmaqta.Caption
oExcel.ActiveSheet.Range("i2").Value = lostad
oExcel.ActiveSheet.Range("t2").Value = lmadras.Caption
oExcel.ActiveSheet.Range("m1").Value = ltsho.Caption
oExcel.ActiveSheet.Range("t1").Value = lkodclass.Caption
oExcel.ActiveSheet.Range("m2").Value = lzsho.Caption
oExcel.ActiveSheet.Range("o2").Value = lzpa.Caption
'Dim NumberOfRows As Integer

NumberOfRows = Student.Recordset.RecordCount
For r = 6 To NumberOfRows + 11 Step 1
On Error Resume Next
GoTo 14

oExcel.ActiveSheet.Range("B" & r).Value = Student.Recordset.Fields("NAME") & " " & Student.Recordset.Fields("FAMIL")
'oExcel.ActiveSheet.Range("C" & r).Value = Student.Recordset.Fields("FAMIL")
'oExcel.ActiveSheet.Range("X" & r).Value = Student.Recordset.Fields("tell") & "-" & Student.Recordset.Fields("mob")



oExcel.ActiveSheet.Range("s" & r).Value = Student.Recordset.Fields("tell")
oExcel.ActiveSheet.Range("s" & r + 1).Value = Student.Recordset.Fields("mob")

14:

oExcel.ActiveSheet.Range("AB" & r + 40).Value = Student.Recordset.Fields("tell")
oExcel.ActiveSheet.Range("AC" & r + 40).Value = Student.Recordset.Fields("mob")

oExcel.ActiveSheet.Range("z" & r + 40).Value = Student.Recordset.Fields("NAME")
oExcel.ActiveSheet.Range("aa" & r + 40).Value = Student.Recordset.Fields("FAMIL")


Student.Recordset.MoveNext
Next





MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  Ê«—œ »—‰«„Â «ò”· ‘œ‰œ", 64, "À»  «ÿ·«⁄« "
AD = lkodclass.Caption
'oExcel.SaveAs AD
'oExcel.Application.Visible = True
'oExcel.Parent.Windows(1).Visible = True
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
oExcel.Application.Visible = True
On Error GoTo 7226


oExcel.Parent.Windows(2).Visible = True
GoTo 9106
7226:

oExcel.Parent.Windows(1).Visible = True
9106:
''''''

oExcel.SaveAs AD
'oExcel.Close
'
'
'Set oExcel = Nothing ' Remove object variable.
''''''''
'Shell "Explorer.exe " & "c:\" & KodEnhesariPrint & ".xlsx"

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End If


'list_nomre_&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

If mnu_sabt_nomre.Checked = True Then

'»«Ìœ €Ì  Â«—« Ê«—œ «ò”· ò‰œ
Student.Refresh
Student.RecordSource = "select * from student where clas1 like ('%" + lkodclass.Caption + "%') or clas2 like ('%" + lkodclass.Caption + "%') or clas3 like ('%" + lkodclass.Caption + "%') or clas4 like ('%" + lkodclass.Caption + "%') or clas5 like ('%" + lkodclass.Caption + "%')"
Student.Refresh



'Dim oExcel As Object
'Dim oBook As Object
'Dim oSheet As Object
'Dim AD As String
'Student.Recordset.MoveFirst
On Error GoTo 982298
GoTo 992299

982298:
MsgBox "›«Ì· ÅÌœ« ‰‘œ∫ »—«Ì «Ã—«Ì «Ì‰ ›—„«‰ »«Ìœ «“ ”Ì” „ ”—Ê— «” ›«œÂ ò‰Ìœ", vbCritical, "Œÿ«"

Exit Sub

992299:

If Entekhab.Pc.Checked = True Then
Set oExcel = GetObject(Entekhab.AdressLabel.Caption & "list_nomre.xlsx")

End If

If Entekhab.net.Checked = True Then
Set oExcel = GetObject(Entekhab.NetAdresslabel.Caption & "list_nomre.xlsx")
End If



'Set oExcel = GetObject("F:\Markaz Quran & Hadis\FORMXLS\ListclassJadid.xlsx")
oExcel.ActiveSheet.Range("C2").Value = ltarh.Caption
oExcel.ActiveSheet.Range("f2").Value = lmaqta.Caption
oExcel.ActiveSheet.Range("i2").Value = lostad
oExcel.ActiveSheet.Range("t2").Value = lmadras.Caption
oExcel.ActiveSheet.Range("m1").Value = ltsho.Caption
oExcel.ActiveSheet.Range("s1").Value = lkodclass.Caption
oExcel.ActiveSheet.Range("m2").Value = lzsho.Caption
oExcel.ActiveSheet.Range("o2").Value = lzpa.Caption
'Dim NumberOfRows As Integer

NumberOfRows = Student.Recordset.RecordCount
For r = 6 To NumberOfRows + 11 Step 1
On Error Resume Next
GoTo 1224

oExcel.ActiveSheet.Range("B" & r).Value = Student.Recordset.Fields("NAME") & " " & Student.Recordset.Fields("FAMIL")
'oExcel.ActiveSheet.Range("C" & r).Value = Student.Recordset.Fields("FAMIL")
'oExcel.ActiveSheet.Range("X" & r).Value = Student.Recordset.Fields("tell") & "-" & Student.Recordset.Fields("mob")



oExcel.ActiveSheet.Range("s" & r).Value = Student.Recordset.Fields("tell")
oExcel.ActiveSheet.Range("s" & r + 1).Value = Student.Recordset.Fields("mob")

1224:

oExcel.ActiveSheet.Range("AB" & r + 40).Value = Student.Recordset.Fields("tell")
oExcel.ActiveSheet.Range("AC" & r + 40).Value = Student.Recordset.Fields("mob")

oExcel.ActiveSheet.Range("z" & r + 40).Value = Student.Recordset.Fields("NAME")
oExcel.ActiveSheet.Range("aa" & r + 40).Value = Student.Recordset.Fields("FAMIL")


Student.Recordset.MoveNext
Next





MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  Ê«—œ »—‰«„Â «ò”· ‘œ‰œ", 64, "À»  «ÿ·«⁄« "
AD = lkodclass.Caption
'oExcel.SaveAs AD
'oExcel.Application.Visible = True
'oExcel.Parent.Windows(1).Visible = True
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
oExcel.Application.Visible = True
On Error GoTo 722226


oExcel.Parent.Windows(2).Visible = True
GoTo 912206
722226:

oExcel.Parent.Windows(1).Visible = True
912206:
''''''

oExcel.SaveAs AD
'oExcel.Close
'
'
'Set oExcel = Nothing ' Remove object variable.
''''''''
'Shell "Explorer.exe " & "c:\" & KodEnhesariPrint & ".xlsx"

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End If



'list_nomre&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

End Sub



Private Sub Command8_Click()
Call Command7_Click

End Sub

Private Sub Command9_Click()



Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim AD As String
Student.Recordset.MoveFirst
On Error GoTo 1
GoTo 2

1: MsgBox "›«Ì· ÅÌœ« ‰‘œ∫ »—«Ì «Ã—«Ì «Ì‰ ›—„«‰ »«Ìœ «“ ”Ì” „ ”—Ê— «” ›«œÂ ò‰Ìœ", vbCritical, "Œÿ«"

Exit Sub

2:


' »—«Ì «Ì‰òÂ ﬁ« Ì ‰‘Â »«Ìœ ò«„·« Ãœ« »«‘œÂ
'·Ì”  «⁄·«‰ A4
'A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A44A4A4A4A4A4A44A4A44A4A4A44A









Set oExcel = GetObject(Entekhab.AdressLabel.Caption & "Etmamclass.xlsx")
oExcel.ActiveSheet.Range("b3").Value = lkodclass.Caption
oExcel.ActiveSheet.Range("d3").Value = ltarh.Caption
oExcel.ActiveSheet.Range("f3").Value = lmaqta.Caption
oExcel.ActiveSheet.Range("h3").Value = lzsho.Caption & " «·Ì  " & lzpa.Caption
oExcel.ActiveSheet.Range("b4").Value = lostad.Caption
oExcel.ActiveSheet.Range("E4").Value = ltsho.Caption
oExcel.ActiveSheet.Range("f4").Value = ltpa.Caption

oExcel.ActiveSheet.Range("h4").Value = lmadras.Caption
oExcel.ActiveSheet.Range("b5").Value = mclass.Recordset.Fields("tozih")


STU2CLASS.Refresh
STU2CLASS.RecordSource = "select * from STU2CLASS where kodclass like ('%" + lkodclass.Caption + "%') and elat like ('%" + "« „«„" + "%') "
STU2CLASS.Refresh


Dim NumberOfRows As Integer
NumberOfRows = STU2CLASS.Recordset.RecordCount
For r = 8 To NumberOfRows + 7







Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + STU2CLASS.Recordset.Fields("parvande") + "%') "
Student.Refresh


oExcel.ActiveSheet.Range("b" & r).Value = Student.Recordset.Fields("PARVANDE")
oExcel.ActiveSheet.Range("c" & r).Value = Student.Recordset.Fields("NAME")
oExcel.ActiveSheet.Range("d" & r).Value = Student.Recordset.Fields("FAMIL")
oExcel.ActiveSheet.Range("g" & r).Value = Student.Recordset.Fields("tell") & "-" & Student.Recordset.Fields("mob")

oExcel.ActiveSheet.Range("e" & r).Value = STU2CLASS.Recordset.Fields("elat")
oExcel.ActiveSheet.Range("f" & r).Value = STU2CLASS.Recordset.Fields("tpayan")
oExcel.ActiveSheet.Range("h" & r).Value = STU2CLASS.Recordset.Fields("tozih")


STU2CLASS.Recordset.MoveNext
Next

MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  Ê«—œ »—‰«„Â «ò”· ‘œ‰œ", 64, "À»  «ÿ·«⁄« "
AD = lkodclass.Caption & "« „«„ ò·«”"
oExcel.SaveAs AD
oExcel.Application.Visible = True
oExcel.Parent.Windows(1).Visible = True



'A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A4A44A4A4A4A4A4A44A4A44A4A4A44A




End Sub

Private Sub DataGridSTUDENT_DblClick()
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>


Dim A As String

On Error GoTo 7
GoTo 8
7:
MsgBox "›«Ì· ÅÌœ« ‰‘œ", vbCritical + vbOKOnly, "Œÿ«"
'Scan.Hide

Exit Sub

8:

If Entekhab.Pc.Checked = True Then
A = Student.Recordset.Fields("PARVANDE")

Scan.Text1.Text = A

Scan.Show
A = SettingF.ScanAdress.Caption & A & "\" & A & ".jpg"
'A = Student.Recordset.Fields("scan")
Scan.Im1.Picture = LoadPicture(A)

Exit Sub
End If

If Entekhab.net.Checked = True Then
A = Student.Recordset.Fields("PARVANDE")

Scan.Text1.Text = A

Scan.Show
'\\Yafatemeh2-pc\f\Markaz Quran & Hadis\FormScan\Pic\9020204\9020204.jpg
A = SettingF.NetScanAdress.Caption & A & "\" & A & ".jpg"
'A = Student.Recordset.Fields("scan")
Scan.Im1.Picture = LoadPicture(A)
Exit Sub
End If

Scan.Text1.Text = A


'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Sub

Private Sub elana4_Click()
lclasslist1.Checked = False
lclasslist.Checked = False
elana4.Checked = True
elana5.Checked = False
foroshlist.Checked = False
mnu_sabt_nomre.Checked = False
End Sub

Private Sub elana5_Click()
lclasslist1.Checked = False
lclasslist.Checked = False
elana4.Checked = False
elana5.Checked = True
foroshlist.Checked = False
mnu_sabt_nomre.Checked = False
End Sub


Private Sub Form_Load()
Student.Refresh

'SB1.Panels(1).Text = user.OP.Text & Start.LD.Caption
List1.RightToLeft = True

Me.stb1.Panels(1).Text = user.OP.Text
Me.stb1.Panels(3).Text = Taqvim.Tarikh.Caption


Combo2.AddItem ("ﬁ—¬‰ ¬„Ê“«‰ Õ–› ‘œÂ »Â œ·Ì· €Ì» ")
Combo2.AddItem ("·Ì”  « „«„ ò·«”")
Combo2.AddItem ("·Ì”  ò«„· ò·«”")

telat.AddItem ("€Ì»  »Ì‘ «“ Õœ")
telat.AddItem ("« „«„ ò·«”")
telat.AddItem ("«‰ ﬁ«· »Â ò·«” œÌê—")
telat.AddItem ("«‘ »«Â œ— À» ")
telat.AddItem ("«‰’—«›")





End Sub

Private Sub Form_Unload(Cancel As Integer)
Entekhab.Show
FClassroom.Hide

End Sub

Private Sub foroshlist_Click()
lclasslist1.Checked = False
lclasslist.Checked = False
elana4.Checked = False
elana5.Checked = False
foroshlist.Checked = True
mnu_sabt_nomre.Checked = False
End Sub

Private Sub lclass1_Click()

End Sub

Private Sub Label8_Change()
On Error Resume Next

Combo1.Clear


Combo1.AddItem (Me.Student.Recordset.Fields("clas1"))
Combo1.AddItem (Me.Student.Recordset.Fields("clas2"))
Combo1.AddItem (Me.Student.Recordset.Fields("clas3"))
Combo1.AddItem (Me.Student.Recordset.Fields("clas4"))
Combo1.AddItem (Me.Student.Recordset.Fields("clas5"))

Combo1.Text = Combo1.List(0)



End Sub

Private Sub lCLASSLIST_Click()
lclasslist1.Checked = False
mnu_sabt_nomre.Checked = False
lclasslist.Checked = True
elana4.Checked = False
elana5.Checked = False
foroshlist.Checked = False

End Sub

Private Sub lclasslist1_Click()
lclasslist1.Checked = True
mnu_sabt_nomre.Checked = False
lclasslist.Checked = False
elana4.Checked = False
elana5.Checked = False
foroshlist.Checked = False
End Sub

Private Sub List1_Click()
On Error Resume Next
Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + mid(List1.Text, 1, 7) + "%')"
Student.Refresh

End Sub

Private Sub List2_Click()
On Error Resume Next
Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + mid(List2.Text, 1, 7) + "%')"
Student.Refresh

End Sub

Private Sub List3_Click()
On Error Resume Next
Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + mid(List3.Text, 1, 7) + "%')"
Student.Refresh

End Sub

Private Sub List4_Click()
On Error Resume Next
Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + mid(List4.Text, 1, 7) + "%')"
Student.Refresh

End Sub

Private Sub ltsho_Change()
ttshoro.Text = ltsho.Caption
End Sub

Private Sub mnubank_Click()
BankStudent.Show

End Sub

Private Sub mnuclasjadid_Click()
ModiriyatCLASS.Show

End Sub

Private Sub mnu_sabt_nomre_Click()
lclasslist1.Checked = False


lclasslist.Checked = False
elana4.Checked = False
elana5.Checked = False
foroshlist.Checked = False
mnu_sabt_nomre.Checked = True

End Sub

Private Sub mnudellall_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "fclass-realy-delete" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
        If ttpayan.Text = "" Or telat.Text = "" Then
           ' MsgBox "·ÿ›«  «—ÌŒ Ê ⁄·  Õ–› —« œ—Ã ò‰Ìœ", vbInformation + vbOKOnly, "Õ–› ﬁ—¬‰ ¬„Ê“"
           GoTo 9996
        Else
9996:



            STU2CLASS.Refresh
            STU2CLASS.RecordSource = " select * from stu2class where  parvande like ('%" + Label8.Caption + "%') and kodclass like ('%" + lkodclass.Caption + "%')"
            STU2CLASS.Refresh
            If STU2CLASS.Recordset.BOF = True Or STU2CLASS.Recordset.EOF = True Then
                MsgBox "‰«„ «Ì‰ ﬁ—¬‰ ¬„Ê“ œ— «Ì‰ ò·«” À»  ‰‘œÂ «”  ", vbCritical + vbOKOnly, "Œÿ«"
                Exit Sub

            Else
                Student.RecordSource = "select * from student where parvande like ('%" + Label8.Caption + "%')"
                Student.Refresh
               

                If MsgBox("  ¬Ì« „Ì ŒÊ«ÂÌœ ¬ﬁ«Ì  " & Label10.Caption & "  «“ ·Ì”  ò·«”Ì œ«—«Ì òœ    " & lkodclass.Caption & "  Õ–› ò‰Ìœ", vbQuestion + vbYesNo, "Õ–› ò‰Ìœ") = vbYes Then

                    '??? ?? ????? ?? ????? ???? ??? ???
                    If Student.Recordset.Fields("clas1") = lkodclass.Caption Then
                        Student.Recordset.Fields("clas1") = "‰œ«—œ"
                    Else
                        If Student.Recordset.Fields("clas2") = lkodclass.Caption Then
                            Student.Recordset.Fields("clas2") = "‰œ«—œ"
                        Else
                            If Student.Recordset.Fields("clas3") = lkodclass.Caption Then
                                Student.Recordset.Fields("clas3") = "‰œ«—œ"
                            Else
                                If Student.Recordset.Fields("clas4") = lkodclass.Caption Then
                                    Student.Recordset.Fields("clas4") = "‰œ«—œ"
                                Else
                                    If Student.Recordset.Fields("clas5") = lkodclass.Caption Then
                                        Student.Recordset.Fields("clas5") = "‰œ«—œ"
                                    Else
                                        MsgBox "ﬁ—¬‰ ¬„Ê“ ﬁ»·« «“ ò·«” Õ–› ‘œÂ «” ", vbExclamation + vbOKOnly, "Õ–› ﬁ—¬‰ ¬„Ê“"
                                        Exit Sub

                                    End If
                                End If
                            End If
                        End If
                    End If
                Else  ' ¬Ì« „Ì ŒÊ«ÂÌœ ﬁ—¬‰ ¬„Ê“ «— ·Ì” ò·«” Õ–› ‘Êœ   ‰Â ‰„Ì ŒÊ«ÂÌ„ Õ–› ‘Êœ
                Exit Sub
                
                
                End If





                STU2CLASS.Refresh
                STU2CLASS.RecordSource = "select * from stu2class where parvande like ('%" + Label8.Caption + "%') and kodclass like ('%" + lkodclass.Caption + "%')"
                STU2CLASS.Refresh

STU2CLASS.Recordset.Delete

              '  STU2CLASS.Recordset.Fields("tpayan") = ttpayan.Text
              '  STU2CLASS.Recordset.Fields("elat") = telat.Text
              '  STU2CLASS.Recordset.Fields("tozih") = ttozih.Text
                Student.Recordset.Update
              '  STU2CLASS.Recordset.Update

                MsgBox "ﬁ—¬‰ ¬„Ê“ »Â ’Ê—  ò«„· «“ ·Ì”  ò·«” Õ–› ‘œ", vbInformation + vbOKOnly, "Õ–› ﬁ—¬‰ ¬„Ê“"
     

           End If
      
            
        End If '·Ì›«  «—ÌŒ Ê ⁄·Ì  Õ–› —« »Ì«‰ òÌ‰œ 
        
        
        
        
        Student.Refresh
        Student.RecordSource = "select * from student where clas1 like ('%" + lkodclass.Caption + "%') or clas2 like ('%" + lkodclass.Caption + "%') or clas3 like ('%" + lkodclass.Caption + "%') or clas4 like ('%" + lkodclass.Caption + "%') or clas5 like ('%" + lkodclass.Caption + "%')"
        Student.Refresh
        List1.Clear
        Label33.Caption = Student.Recordset.RecordCount

        For I = 1 To Student.Recordset.RecordCount
        List1.AddItem (Student.Recordset.Fields("parvande") & "  -  " & Student.Recordset.Fields("name") & "  -  " & Student.Recordset.Fields("famil"))
        Student.Recordset.MoveNext
        Next I




' ÂÌÂ ·Ì”  »«ò” »—«Ì ò·«”
Student.Refresh
Student.RecordSource = "select * from student where clas1 like ('%" + lkodclass.Caption + "%') or clas2 like ('%" + lkodclass.Caption + "%') or clas3 like ('%" + lkodclass.Caption + "%') or clas4 like ('%" + lkodclass.Caption + "%') or clas5 like ('%" + lkodclass.Caption + "%')"
Student.Refresh
List1.Clear
Label33.Caption = Student.Recordset.RecordCount

For I = 1 To Student.Recordset.RecordCount
List1.AddItem (Student.Recordset.Fields("parvande") & "  -  " & Student.Recordset.Fields("name") & "  -  " & Student.Recordset.Fields("famil"))
Student.Recordset.MoveNext
Next I



End Sub

Private Sub mnuetmam_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "fclass-endclass" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
Call Command5_Click

End Sub

Private Sub mnufclass_Click()
Beep

End Sub

Private Sub mnugozaresh_Click()
Gozaresh.Show

End Sub

Private Sub mnukarname_Click()
Karname.Show

End Sub

Private Sub mnuplclas_Click()
FClassroom.Show

End Sub

Private Sub mnuqeybat_Click()
QeybatF.Show

End Sub

Private Sub MNUSABTNOMARAT_Click()
EmtahanF.Show

End Sub

Private Sub mnuhiome_Click()
Entekhab.Show

End Sub

Private Sub mnusabtdar_Click()
Call Command1_Click

End Sub

Private Sub stb1_PanelClick(ByVal Panel As ComctlLib.Panel)
ttpayan.Text = Me.stb1.Panels(3).Text

'If ttshoro.ZOrder = True Then
ttshoro.Text = Me.stb1.Panels(3).Text
'End If


End Sub

Private Sub Text1_Change()
On Error Resume Next



'Command6.Enabled = False  ' ﬂ«„«‰œ ﬁ»·Ì

'If Option6.Value = True Then 'Ã” ÊÃÊ œ— «”«„Ì ﬁ—¬‰ ¬„Ê“«‰

If Option1.Value = True Then 'Ã” ÊÃÊ »— «”«” ‘„«—Â Å—Ê‰œÂ
Student.RecordSource = "select * from student where parvande like ('%" + Text1 + "%')"
Student.Refresh
DataGrid1.Refresh
End If
If Option2.Value = True Then  ' Ã” ÊÃÊ »— «”«” ‰«„ Œ«‰Ê«œêÌ
Student.RecordSource = "select * from student where famil like ('%" + Text1 + "%') or name like ('%" + Text1 + "%')  or parvande like ('%" + Text1 + "%') or nf like ('%" + Text1 + "%')"
Student.Refresh
DataGrid1.Refresh
End If
If Option3.Value = True Then  'Ã” ÊÃÊ »— «”«” ﬂœ „·Ì
Student.RecordSource = "select * from student where kodMeli like ('%" + Text1 + "%')"
Student.Refresh
DataGrid1.Refresh
End If
If Option4.Value = True Then  '»— «”«” ‰«„ «” «œ
Student.RecordSource = "select * from student where ostad like ('%" + Text1 + "%')"
Student.Refresh
DataGrid1.Refresh
End If
If Option5.Value = True Then  'ﬂœ ﬂ·«”
Student.RecordSource = "select * from student where clas like ('%" + Text1 + "%') or clas2 like ('%" + Text1 + "%')or clas3 like ('%" + Text1 + "%')or clas4 like ('%" + Text1 + "%')or clas5 like ('%" + Text1 + "%')"

Student.Refresh
DataGrid1.Refresh
End If
If Option8.Value = True Then  'Â„—«Â
Student.RecordSource = "select * from student where mob like ('%" + Text1 + "%')"
Student.Refresh
DataGrid1.Refresh
End If






End Sub

Private Sub Text2_Change()
STU2CLASS.Refresh
STU2CLASS.RecordSource = "select * from STU2CLASS where kodclass like ('%" + Text2.Text + "%')"
STU2CLASS.Refresh

End Sub

Private Sub Text1_Click()
Text1.Text = ""

End Sub

Private Sub Text3_Change()
mclass.Refresh
mclass.RecordSource = "select * from mclass where tarh like ('%" + Text3.Text + "%') or maqta like ('%" + Text3.Text + "%')or ostad like ('%" + Text3.Text + "%')or kodclass like ('%" + Text3.Text + "%') or tozih like ('%" + Text3.Text + "%')"
mclass.Refresh

End Sub
