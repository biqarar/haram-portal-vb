VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Gozaresh 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ê“«—‘ êÌ—Ì «“ Ê÷⁄Ì  ﬁ—¬‰ ¬„Ê“«‰"
   ClientHeight    =   10245
   ClientLeft      =   4935
   ClientTop       =   2595
   ClientWidth     =   12840
   BeginProperty Font 
      Name            =   "B Homa"
      Size            =   9
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "qozaresh.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10245
   ScaleWidth      =   12840
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      BackColor       =   &H0080FFFF&
      Caption         =   "Õ–› «“ ·Ì”  «Œÿ«—"
      Height          =   375
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   70
      ToolTipText     =   "»« «Ì‰ œò„Â  ‰Â« ‰«„ ﬁ—¬‰ ¬„Ê“ «“ ·Ì”  Å«ò „Ì ‘Êœ Ê œ— ê“«—‘ êÌ—Ì »⁄œÌ œÊ»«—Â Ê«—œ ·Ì”  „Ì ‘Êœ"
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "‰„«Ì‘ Ã“∆Ì« "
      Height          =   330
      Left            =   2160
      TabIndex        =   69
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame Frame7 
      Caption         =   "«ÿ·«⁄«   ⁄Âœ"
      Height          =   2535
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   68
      Top             =   0
      Width           =   3255
      Begin VB.ComboBox Combo1 
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
         ItemData        =   "qozaresh.frx":08CA
         Left            =   1200
         List            =   "qozaresh.frx":08CC
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Text            =   "›—Ê—œÌ‰"
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox Combo3 
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
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Text            =   "1"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Text            =   "»Â ⁄·  €Ì»  »Ì‘ «“ Õœ"
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Text            =   "1390"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Text            =   " Ê÷ÌÕ« "
         Top             =   1920
         Width           =   2895
      End
      Begin VB.ComboBox Combo2 
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
         ItemData        =   "qozaresh.frx":08CE
         Left            =   240
         List            =   "qozaresh.frx":08D8
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Text            =   " ⁄Âœ ‘›«ÂÌ"
         Top             =   840
         Width           =   2895
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "ò·«” Â«Ì ﬁ—¬‰ ¬„Ê“"
      Height          =   2535
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   0
      Width           =   2175
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         DataField       =   "Clas5"
         DataSource      =   "ekhtar"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   240
         TabIndex        =   66
         ToolTipText     =   "»— —ÊÌ òœ ò·«” ò·Ìò ò‰Ìœ  « „‘Œ’«  ¬‰ ‰„«Ì‘ œ«œÂ ‘Êœ"
         Top             =   2160
         Width           =   45
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         DataField       =   "Clas4"
         DataSource      =   "ekhtar"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   240
         TabIndex        =   65
         ToolTipText     =   "»— —ÊÌ òœ ò·«” ò·Ìò ò‰Ìœ  « „‘Œ’«  ¬‰ ‰„«Ì‘ œ«œÂ ‘Êœ"
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         DataField       =   "Clas3"
         DataSource      =   "ekhtar"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1200
         TabIndex        =   64
         ToolTipText     =   "»— —ÊÌ òœ ò·«” ò·Ìò ò‰Ìœ  « „‘Œ’«  ¬‰ ‰„«Ì‘ œ«œÂ ‘Êœ"
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         DataField       =   "Clas2"
         DataSource      =   "ekhtar"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   240
         TabIndex        =   63
         ToolTipText     =   "»— —ÊÌ òœ ò·«” ò·Ìò ò‰Ìœ  « „‘Œ’«  ¬‰ ‰„«Ì‘ œ«œÂ ‘Êœ"
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         DataField       =   "Clas1"
         DataSource      =   "ekhtar"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1200
         TabIndex        =   62
         ToolTipText     =   "»— —ÊÌ òœ ò·«” ò·Ìò ò‰Ìœ  « „‘Œ’«  ¬‰ ‰„«Ì‘ œ«œÂ ‘Êœ"
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "ò·«”5"
         Height          =   330
         Left            =   1200
         TabIndex        =   61
         Top             =   2160
         Width           =   450
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "ò·«”4"
         Height          =   330
         Left            =   240
         TabIndex        =   60
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "ò·«” 3"
         Height          =   330
         Left            =   1200
         TabIndex        =   59
         Top             =   1200
         Width           =   510
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "ò·«” 2"
         Height          =   330
         Left            =   240
         TabIndex        =   58
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "ò·«” 1"
         Height          =   330
         Left            =   1200
         TabIndex        =   57
         Top             =   360
         Width           =   450
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "qozaresh.frx":08F3
      Height          =   4935
      Left            =   120
      TabIndex        =   13
      Top             =   4920
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   8705
      _Version        =   393216
      BackColor       =   8421631
      DefColWidth     =   80
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
      Caption         =   "·Ì”  «Œÿ«—"
      ColumnCount     =   12
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
         DataField       =   "NamePedar"
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
         DataField       =   "Clas1"
         Caption         =   "ò·«”"
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
         DataField       =   "Tedad"
         Caption         =   " ⁄œ«œ €Ì» "
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
         DataField       =   "VP"
         Caption         =   "Ê÷⁄Ì  Å—Ê‰œÂ"
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
      BeginProperty Column08 
         DataField       =   "Clas2"
         Caption         =   "ò·«” 2"
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
      BeginProperty Column10 
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
      BeginProperty Column11 
         DataField       =   "Clas5"
         Caption         =   "ò·«”5"
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
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1214.929
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame5 
      Caption         =   "„‘Œ’«  ﬁ—¬‰ ¬„Ê“"
      Height          =   3375
      Left            =   9120
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   0
      Width           =   3615
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   " ·›‰"
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
         Left            =   3120
         TabIndex        =   55
         Top             =   2400
         Width           =   270
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Tell"
         DataSource      =   "ekhtar"
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
         Left            =   480
         TabIndex        =   54
         Top             =   2400
         Width           =   135
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "VP"
         DataSource      =   "ekhtar"
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
         Left            =   480
         TabIndex        =   53
         Top             =   1920
         Width           =   135
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   " Ê÷ÌÕ« "
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
         Left            =   2760
         TabIndex        =   52
         Top             =   2040
         Width           =   585
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Tedad"
         DataSource      =   "ekhtar"
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
         Left            =   480
         TabIndex        =   51
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "famil"
         DataSource      =   "ekhtar"
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
         Left            =   480
         TabIndex        =   50
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "name"
         DataSource      =   "ekhtar"
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
         Left            =   480
         TabIndex        =   49
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Parvande"
         DataSource      =   "ekhtar"
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
         Left            =   480
         TabIndex        =   48
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   " ⁄œ«œ €Ì» "
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
         Left            =   2640
         TabIndex        =   47
         Top             =   1560
         Width           =   705
      End
      Begin VB.Label Label4 
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
         Left            =   2400
         TabIndex        =   46
         Top             =   1080
         Width           =   870
      End
      Begin VB.Label Label3 
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
         Left            =   3120
         TabIndex        =   45
         Top             =   720
         Width           =   195
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
         Index           =   2
         Left            =   2520
         TabIndex        =   44
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "„‘Œ’«  ò·«”"
      Height          =   3375
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   0
      Width           =   3255
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ Å«Ì«‰"
         Height          =   345
         Left            =   2040
         TabIndex        =   67
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   " «"
         Height          =   330
         Left            =   720
         TabIndex        =   42
         Top             =   1800
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
         Left            =   1080
         TabIndex        =   41
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label lmadras 
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
         Left            =   240
         TabIndex        =   40
         Top             =   2040
         Width           =   225
      End
      Begin VB.Label lostad 
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
         Left            =   240
         TabIndex        =   39
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
         Left            =   240
         TabIndex        =   38
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label lmaqta 
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
         Left            =   240
         TabIndex        =   37
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label ltarh 
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
         Left            =   240
         TabIndex        =   36
         Top             =   720
         Width           =   135
      End
      Begin VB.Label lkodclass 
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
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "„œ—”"
         Height          =   330
         Left            =   2040
         TabIndex        =   34
         Top             =   2160
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "”«⁄ "
         Height          =   330
         Left            =   2040
         TabIndex        =   33
         Top             =   1800
         Width           =   405
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "‰«„ «” «œ"
         Height          =   330
         Left            =   2040
         TabIndex        =   32
         Top             =   1440
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "„ﬁÿ⁄"
         Height          =   330
         Index           =   1
         Left            =   2040
         TabIndex        =   31
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "ÿ—Õ"
         Height          =   330
         Left            =   2040
         TabIndex        =   30
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "òœ ò·«”"
         Height          =   330
         Index           =   0
         Left            =   2040
         TabIndex        =   29
         Top             =   360
         Width           =   555
      End
      Begin VB.Label ltsho 
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
         Left            =   240
         TabIndex        =   28
         Top             =   2400
         Width           =   135
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ ‘—Ê⁄"
         Height          =   345
         Left            =   2040
         TabIndex        =   27
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label ltpa 
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
         Left            =   240
         TabIndex        =   26
         Top             =   2760
         Width           =   135
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " ⁄œ«œ €Ì»  Â«"
      Height          =   975
      Left            =   4920
      TabIndex        =   22
      Top             =   3360
      Width           =   3135
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         Caption         =   " ⁄œ«œ €Ì»  Â«Ì »——”Ì ‰‘œÂ"
         Height          =   255
         Left            =   720
         TabIndex        =   73
         Top             =   240
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "»«  ÊÃÂ »Â €Ì»  „Ã«“ œ— Â— ò·«”"
         Height          =   255
         Left            =   480
         TabIndex        =   72
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Text            =   "7"
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000002&
      Caption         =   "ê“«—‘ «“ Ê÷⁄Ì  €Ì»  Â«"
      Height          =   375
      Left            =   120
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "‘—Ê⁄ ê“«—‘ êÌ—Ì «“ Ê÷⁄Ì  €Ì»  Â«"
      Top             =   4440
      Width           =   10575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Motor"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   9240
      Visible         =   0   'False
      Width           =   8175
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
         Connect         =   $"qozaresh.frx":0908
         OLEDBString     =   $"qozaresh.frx":0991
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
         Connect         =   $"qozaresh.frx":0A1A
         OLEDBString     =   $"qozaresh.frx":0AA3
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
         Connect         =   $"qozaresh.frx":0B2C
         OLEDBString     =   $"qozaresh.frx":0BB5
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
         Connect         =   $"qozaresh.frx":0C3E
         OLEDBString     =   $"qozaresh.frx":0CC7
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
         Connect         =   $"qozaresh.frx":0D50
         OLEDBString     =   $"qozaresh.frx":0DD9
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
         Connect         =   $"qozaresh.frx":0E62
         OLEDBString     =   $"qozaresh.frx":0EEB
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
         Connect         =   $"qozaresh.frx":0F74
         OLEDBString     =   $"qozaresh.frx":0FFD
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
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080FF80&
      Caption         =   "‰«œÌœÂ ê—› ‰ «Œÿ«—"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "À»   ⁄Âœ"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "Õ–› ﬁ—¬‰ ¬„Ê“ «“ ò·«”"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000E&
      Caption         =   "Å«ò ò—œ‰ ÃœÊ·"
      Height          =   690
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      DisabledPicture =   "qozaresh.frx":1086
      DownPicture     =   "qozaresh.frx":25D00
      DragIcon        =   "qozaresh.frx":4A97A
      Height          =   375
      Left            =   2280
      Picture         =   "qozaresh.frx":6F5F4
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "«‰ ﬁ«· ÃœÊ· »Â »—‰«„Â «ò”·"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox Text2 
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
      Height          =   405
      Left            =   2160
      TabIndex        =   3
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ã” ÊÃÊ œ— ·Ì”  «Œÿ«— »— «”«”"
      Height          =   975
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   3360
      Width           =   4575
      Begin VB.OptionButton omm 
         Alignment       =   1  'Right Justify
         Caption         =   "òœ ò·«”"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton ofamil 
         Alignment       =   1  'Right Justify
         Caption         =   "‰«„ Ê ‰«„ Œ«‰Ê«œêÌ"
         Height          =   255
         Left            =   1320
         TabIndex        =   16
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton okod 
         Alignment       =   1  'Right Justify
         Caption         =   "‘„«—Â Å—Ê‰œÂ"
         Height          =   255
         Left            =   3000
         TabIndex        =   15
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin ComctlLib.ProgressBar PR1 
      Height          =   135
      Left            =   2160
      TabIndex        =   18
      Top             =   4080
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.StatusBar stb1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   71
      Top             =   9870
      Width           =   12840
      _ExtentX        =   22648
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
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
   Begin MSAdodcLib.Adodc userprofiletable 
      Height          =   330
      Left            =   2280
      Top             =   10200
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
      Connect         =   $"qozaresh.frx":9426E
      OLEDBString     =   $"qozaresh.frx":942F7
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
   Begin VB.Line Line4 
      X1              =   2280
      X2              =   4800
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "_"
      Height          =   330
      Left            =   3480
      TabIndex        =   24
      Top             =   2640
      Width           =   120
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   " ⁄œ«œ"
      Height          =   330
      Left            =   4080
      TabIndex        =   23
      Top             =   2640
      Width           =   360
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "Ã” ÃÊ "
      Height          =   330
      Left            =   4080
      TabIndex        =   0
      Top             =   3120
      Width           =   465
   End
   Begin VB.Menu mnuhome 
      Caption         =   "#"
   End
End
Attribute VB_Name = "Gozaresh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
On Error Resume Next
ekhtar.Refresh
ekhtar.Recordset.MoveFirst
For I = 1 To ekhtar.Recordset.RecordCount
ekhtar.Recordset.Delete
ekhtar.Recordset.MoveNext
Next I
End Sub

Private Sub Check1_Click()
If Combo1.Enabled = False Then
Combo1.Enabled = True
Else
Combo1.Enabled = False
End If

End Sub

Private Sub Command1_Click()

Dim X, TC As Integer
Dim C As String


'»«Ìœ €Ì  Â«—« Ê«—œ «ò”· ò‰œ
Dim oExcel As Object
'Dim oBook As Object
'Dim oSheet As Object
Dim AD As String
On Error GoTo 1
GoTo 2
1: MsgBox "›«Ì· ÅÌœ« ‰‘œ∫ »—«Ì «Ã—«Ì «Ì‰ ›—„«‰ »«Ìœ «“ ”Ì” „ ”—Ê— «” ›«œÂ ò‰Ìœ", vbCritical, "Œÿ«"
Exit Sub

2:




If Entekhab.Pc.Checked = True Then
Set oExcel = GetObject(Entekhab.AdressLabel.Caption & "ekhtarxls.xlsx")

End If

If Entekhab.net.Checked = True Then
Set oExcel = GetObject(Entekhab.NetAdresslabel.Caption & "ekhtarxls.xlsx")
End If


'Set oExcel = GetObject("F:\Markaz Quran & Hadis\FORMXLS\ekhtarxls.xlsx")
oExcel.ActiveSheet.Range("f1").Value = "«„Ê— ¬„Ê“‘Ì"
oExcel.ActiveSheet.Range("j1").Value = Text10.Text
oExcel.ActiveSheet.Range("M1").Value = Taqvim.Label1.Caption
ekhtar.Recordset.MoveFirst




Dim NumberOfRows As Integer
NumberOfRows = ekhtar.Recordset.RecordCount
For r = 4 To NumberOfRows + 3
oExcel.ActiveSheet.Range("B" & r).Value = ekhtar.Recordset.Fields("parvande")
oExcel.ActiveSheet.Range("C" & r).Value = ekhtar.Recordset.Fields("name")

oExcel.ActiveSheet.Range("D" & r).Value = ekhtar.Recordset.Fields("famil")
oExcel.ActiveSheet.Range("E" & r).Value = ekhtar.Recordset.Fields("namepedar")

'On Error Resume Next



X = 1
TC = 0
'«Ì‰ òÂ ÂÌÃÌ ‰Ì” 
For J = 1 To 5   ' Å‰Ã ò·«” —« çò „Ì ò‰œ


C = "clas" & X
If ekhtar.Recordset.Fields(C) <> "‰œ«—œ" And ekhtar.Recordset.Fields(C) <> "" Then

QW = ekhtar.Recordset.Fields(C)
mclass.Refresh
mclass.RecordSource = "SELECT * FROM MCLASS WHERE KODCLASs LIKE ('%" + QW + "%')"
mclass.Refresh

'oExcel.ActiveSheet.Range("F" & r).Value = ltarh.Caption & " - " & lmaqta.Caption & " - " & lostad.Caption & " - " & lzsho.Caption & " - " & lzpa.Caption & "  „œ—”  " & lmadras.Caption
oExcel.ActiveSheet.Range("F" & r).Value = mclass.Recordset.Fields("tarh") & " - " & mclass.Recordset.Fields("maqta") & " «” «œ " & mclass.Recordset.Fields("ostad") & " - " & mclass.Recordset.Fields("zamaneshoro") & " - " & mclass.Recordset.Fields("zamanepayan") & " „œ—” " & mclass.Recordset.Fields("madras")
'oExcel.ActiveSheet.Range("F" & r).Value = mclass.Recordset.Fields("tarh")
GoTo 52


End If
X = X + 1
Next J

52:



oExcel.ActiveSheet.Range("G" & r).Value = ekhtar.Recordset.Fields("tell")

oExcel.ActiveSheet.Range("H" & r).Value = ekhtar.Recordset.Fields("vp")

oExcel.ActiveSheet.Range("i" & r).Value = ekhtar.Recordset.Fields("tedad")

ekhtar.Recordset.MoveNext

Next

MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  Ê«—œ »—‰«„Â «ò”· ‘œ‰œ", 64, "À»  «ÿ·«⁄« "
AD = lkodclass.Caption
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


End Sub


Private Sub Command3_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513


userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "gozaresh-delete-stu" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
Beep
    If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ ﬁ—¬‰ ¬„Ê“ —« «“ ·Ì”  ò·«”Ì Õ–› ò‰Ìœ" & Chr(10) & "»⁄œ «“ Õ–› ﬁ—¬‰ ¬„Ê“ «“ ò·«” ‰«„ «Ì‘«‰ «“ ·Ì”  «Œÿ«— Å«ò „Ì ‘Êœ", vbQuestion + vbYesNo, "„œÌ—Ì  ·Ì”  «Œÿ«—") = vbYes Then
    
          '  MsgBox "·ÿ›«  «—ÌŒ Ê ⁄·  Õ–› —« œ—Ã ò‰Ìœ", vbInformation + vbOKOnly, "Õ–› ﬁ—¬‰ ¬„Ê“"
       ' Else




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


                STU2CLASS.Recordset.Fields("tpayan") = Taqvim.Label1.Caption
                STU2CLASS.Recordset.Fields("elat") = "€Ì»  »Ì‘ «“ Õœ"
                STU2CLASS.Recordset.Fields("tozih") = "Õ–› ‘œÂ  Ê”ÿ ·Ì”  «Œÿ«—"
                Student.Recordset.Update
                STU2CLASS.Recordset.Update

                MsgBox "ﬁ—¬‰ ¬„Ê“ Õ–› ‘œ", vbInformation + vbOKOnly, "Õ–› ﬁ—¬‰ ¬„Ê“"
     
            Qeybat.Refresh
            Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%') and vazeyat like ('%" + "0" + "%')"
            Qeybat.Refresh

            For I = 1 To Qeybat.Recordset.RecordCount
            
             Qeybat.Recordset.Fields("vazeyat") = "1"
             Qeybat.Recordset.Fields("natije") = "Õ–› «“ ò·«”"
             Qeybat.Recordset.Update
            Qeybat.Recordset.MoveNext
            
            Next I
            
            ekhtar.Recordset.Delete
            
            
            
           End If
      
         

            
            
            
            
            
            
        End If '????? ????? ???? ? ??? ?? ??? ????????
        
        
        
        
End Sub

Private Sub Command4_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "gozaresh-taahod-stu" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
Beep
'If Entekhab.SB.Panels(1).Text = "„ÌÂ„«‰" Then Exit Sub
If MsgBox("  »—«Ì ¬ﬁ«Ì  " & Label10.Caption & "   " & Combo2.Text & "  œ—  «—ÌŒ  " & Combo3.Text & "   " & Combo1.Text & "   " & Text6.Text & "  À»  ŒÊ«Âœ ‘œ ¬Ì« „ÿ„∆‰ Â” Ìœ  ", vbQuestion + vbYesNo, " „œÌ—Ì  ·Ì”  «Œÿ«—") = vbYes Then






' À»  €Ì  «“ ÿ—Ìﬁ òœ ò·«”Ì


Qeybat.Refresh
Qeybat.Recordset.AddNew
Qeybat.Recordset.Fields("Parvande") = Label8.Caption
Qeybat.Recordset.Fields("name") = Label9.Caption
Qeybat.Recordset.Fields("famil") = Label10.Caption
'Qeybat.Recordset.Fields("Ostad") = Label12.Caption
'Qeybat.Recordset.Fields("tarh") = Label11.Caption
Qeybat.Recordset.Fields("sal") = Text6
Qeybat.Recordset.Fields("mah") = Combo1.Text
Qeybat.Recordset.Fields("rooz") = Combo3.Text
Qeybat.Recordset.Fields("noe") = Combo2.Text
Qeybat.Recordset.Fields("elat") = Text3
Qeybat.Recordset.Fields("tozih") = Text7
Qeybat.Recordset.Fields("clas") = lkodclass.Caption
Qeybat.Recordset.Fields("vazeyat") = "3"

Qeybat.Recordset.Fields("natije") = "‰ ÌÃÂ ·Ì”  «Œÿ«—"
Qeybat.Recordset.Update
Qeybat.Refresh





            Qeybat.Refresh
            Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%') and vazeyat like ('%" + "0" + "%')"
            Qeybat.Refresh

            For I = 1 To Qeybat.Recordset.RecordCount
            
             Qeybat.Recordset.Fields("vazeyat") = "1"
             Qeybat.Recordset.Fields("natije") = "À»   ⁄Âœ"
             Qeybat.Recordset.Update
            Qeybat.Recordset.MoveNext
            
            Next I
            
ekhtar.Recordset.Delete




Beep




Exit Sub




' €Ì»   »« ”«” ›«Âœ «“ òœ ò·”« Å«Ì«‰ À» 




End If



End Sub

Private Sub Command5_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "gozaresh-bikhiyal-stu" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
'If Entekhab.SB.Panels(1).Text = "„ÌÂ„«‰" Then Exit Sub
Beep
If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ €Ì»  Â«Ì ﬁ—¬‰ ¬„Ê“ —« ‰«œÌœÂ »êÌ—Ìœ", vbQuestion + vbYesNo, "„œÌ—Ì  ·Ì”  «Œÿ«—") = vbYes Then


            Qeybat.Refresh
            Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label8.Caption + "%') and vazeyat like ('%" + "0" + "%')"
            Qeybat.Refresh

            For I = 1 To Qeybat.Recordset.RecordCount
            
             Qeybat.Recordset.Fields("vazeyat") = "1"
             Qeybat.Recordset.Fields("natije") = "‰«œÌœÂ ê—› ‰ «Œÿ«—"
             Qeybat.Recordset.Update
            Qeybat.Recordset.MoveNext
            
            Next I
            ekhtar.Recordset.Delete
            Beep
          End If
          
End Sub


Private Sub Command6_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513


userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "gozaresh-start" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
Dim X, TC As Integer
Dim C As String

' „«„Ì ﬁ—¬´ ¬„Ê“ «·‰«Ì ‰” 
'›ﬁÿ ò”«‰Ì òÂ œ— ò·«” ‘—ò  „Ì ò‰‰œ

'Â„Â ﬁ—¬‰ ¬„Ê“«‰ —« « ‰ Œ«» „Ì ò‰œ
Student.Refresh
Student.RecordSource = "select * from student where clas1 like ('%" + "9" + "%') or clas2 like ('%" + "9" + "%') or clas3 like ('%" + "9" + "%') or clas4 like ('%" + "9" + "%') or clas5 like ('%" + "9" + "%')"
Student.Refresh

PR1.Value = 0
PR1.Max = Student.Recordset.RecordCount
PR1.Visible = True






For I = 1 To Student.Recordset.RecordCount '‘—Ê⁄ çò ò—œ‰ ò· ﬁ—¬‰ ¬„Ê“«‰
PR1.Value = PR1.Value + 1



GoTo 30
X = 1
TC = 0
'«Ì‰ òÂ ÂÌÃÌ ‰Ì” 
For J = 1 To 5   ' Å‰Ã ò·«” —« çò „Ì ò‰œ


C = "clas" & X
If Student.Recordset.Fields(C) <> "‰œ«—œ" Then

TC = TC + 1 '‘„«—‘ ò·«” Â«

End If
X = X + 1
Next J
If TC >= 1 Then  'ÿ—› »Ì” — «“ 1 ‘—ò  „Ì ò‰œ



30: '«“ «Ê·  „«„ ò”«‰Ì òÂ œ— ò·«” ‘—ò  „Ì òœ—‰œ —« «‰ ŒÕ« òÌœ—Â «”„ 

'ÿ—› ò·«” „Ì ¬Ìœ «Ì‰ òœ »—«Ì ¬‰ «”  òÂ ò”«‰Ì òÂ ò·«” ‰„Ì ¬Ì‰œ œ— ·Ì”  Ê«—œ ‰‘Ê‰œ

' »«Ìœ Ê÷⁄Ì  €Ì» Â«Ì «Ì‰ ‘Œ’ »——”Ì ‘Êœ
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Student.Recordset.Fields("parvande") + "%') and vazeyat like ('%" + "0" + "%')and noe like ('%" + "„" + "%') "
Qeybat.Refresh


If Option1.Value = True Then


X = 1
TC = 0
'«Ì‰ òÂ ÂÌÃÌ ‰Ì” 
For J = 1 To 5   ' Å‰Ã ò·«” —« çò „Ì ò‰œ


C = "clas" & X
If Student.Recordset.Fields(C) <> "‰œ«—œ" Then
GoTo 35 'Ìò ò·«” ÅÌœ« ò—œÂ «”  òÂ «Ì‰ ﬁ—¬‰ ¬„Ê“ œ— ¬‰ ‘—ò  òÌ œò

TC = TC + 1 '‘„«—‘ ò·«” Â«

End If
X = X + 1
Next J

35
mclass.Refresh
mclass.RecordSource = "select * from mclass where kodclass like('%" & Student.Recordset.Fields(C) & "%')"
mclass.Refresh
If Val(Qeybat.Recordset.RecordCount) >= Val(mclass.Recordset.Fields("qmojaz")) Then
 GoTo 70
 
 Else
  GoTo 71
  
  End If
  







End If


If Option2.Value = True Then
If Qeybat.Recordset.RecordCount >= Text10.Text Then
70:

ekhtar.Refresh
ekhtar.Recordset.AddNew

ekhtar.Recordset.Fields("parvande") = Student.Recordset.Fields("parvande")
ekhtar.Recordset.Fields("name") = Student.Recordset.Fields("name")
ekhtar.Recordset.Fields("famil") = Student.Recordset.Fields("famil")

ekhtar.Recordset.Fields("tedad") = Qeybat.Recordset.RecordCount
ekhtar.Recordset.Fields("namepedar") = Student.Recordset.Fields("namepedar")
ekhtar.Recordset.Fields("vp") = Student.Recordset.Fields("tozih")

ekhtar.Recordset.Fields("clas1") = Student.Recordset.Fields("clas1")

ekhtar.Recordset.Fields("clas2") = Student.Recordset.Fields("clas2")
ekhtar.Recordset.Fields("clas3") = Student.Recordset.Fields("clas3")
ekhtar.Recordset.Fields("clas4") = Student.Recordset.Fields("clas4")
ekhtar.Recordset.Fields("clas5") = Student.Recordset.Fields("clas5")


ekhtar.Recordset.Fields("tell") = Student.Recordset.Fields("tell") & " - " & Student.Recordset.Fields("mob")
'ekhtar.Recordset.Fields("Date") = Taqvim.Label1.Caption
ekhtar.Recordset.Update

ekhtar.Refresh
Label17.Caption = ekhtar.Recordset.RecordCount

End If
End If
GoTo 1

Else 'ÿ—› ò·«”Ì ‰œ«‘ Â «” 
1:
71:


Student.Recordset.MoveNext

End If
Next I



'  Å«Ì«‰ »——”Ì Ê÷⁄Ì  €Ì» Â«Ì ÿ—›

PR1.Value = 0
PR1.Visible = False
On Error Resume Next

ekhtar.Recordset.MoveFirst
Label17.Caption = ekhtar.Recordset.RecordCount


End Sub

Private Sub Command7_Click()
QeybatF.Show
QeybatF.Option7.Value = True
QeybatF.Option1.Value = True
QeybatF.Text1.Text = Label8.Caption


End Sub

Private Sub Command8_Click()
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "gozaresh-delete-fromlist" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
On Error Resume Next

ekhtar.Recordset.Delete

End Sub

Private Sub Form_Load()
Me.stb1.Panels(1).Text = user.OP.Text
Me.stb1.Panels(3).Text = Taqvim.Label1.Caption


Combo1.AddItem ("›—Ê—œÌ‰")
Combo1.AddItem ("«—œÌ»Â‘ ")
Combo1.AddItem ("Œ—œ«œ")
Combo1.AddItem (" Ì—")
Combo1.AddItem ("„—œ«œ")
Combo1.AddItem ("‘Â—ÌÊ—")
Combo1.AddItem ("„Â—")
Combo1.AddItem ("¬»«‰")
Combo1.AddItem ("¬–—")
Combo1.AddItem ("œÌ")
Combo1.AddItem ("»Â„‰")
Combo1.AddItem ("«”›‰œ")


For I = 1 To 31 Step 1  ' »—«Ì Ê«—œ ﬂ—œ‰ ‘„«—Â —Ê“ œ— ÃœÊ· «ÿ·«⁄«  €Ì 
Combo3.AddItem (I)
Next I

' «Ì‰ ﬁ”„  ÿ—Õ Â« —« «“ ÃœÊ· ê—› Â Ê «÷«›Â „Ì ﬂ‰œ »Â ﬂ„»Ê »«ﬂ”
'Tarhha.Refresh

' »—«Ì «÷«›Â ﬂ—œ‰ „«Â Â«Ì ”«· »Â ·Ì”  „Ì »«‘œ



GoTo 2


'«Ì‰ ﬁ”„  ÃœÊ· «Œÿ«— —« Å«ﬂ „Ì ﬂ‰œ
On Error Resume Next

For I = 1 To ekhtar.Recordset.RecordCount
ekhtar.Recordset.Delete
ekhtar.Recordset.MoveNext
Next I


2

'Å«ﬂ ﬂ—œ‰ ÃœÊ· «Œ «—  „«„ ‘œ


 
End Sub

Private Sub Form_Unload(Cancel As Integer)
Entekhab.Show
Gozaresh.Hide

End Sub

Private Sub Label30_Change()
mclass.Refresh
mclass.RecordSource = "SELECT * FROM MCLASS WHERE KODCLASs LIKE ('%" + Label30.Caption + "%')"
mclass.Refresh
End Sub

Private Sub Label30_Click()
mclass.Refresh
mclass.RecordSource = "SELECT * FROM MCLASS WHERE KODCLASs LIKE ('%" + Label30.Caption + "%')"
mclass.Refresh
End Sub

Private Sub Label31_Click()
mclass.Refresh
mclass.RecordSource = "SELECT * FROM MCLASS WHERE KODCLASs LIKE ('%" + Label31.Caption + "%')"
mclass.Refresh
End Sub


Private Sub Label40_Click()
mclass.Refresh
mclass.RecordSource = "SELECT * FROM MCLASS WHERE KODCLASs LIKE ('%" + Label40.Caption + "%')"
mclass.Refresh
End Sub

Private Sub Label41_Click()
mclass.Refresh
mclass.RecordSource = "SELECT * FROM MCLASS WHERE KODCLASs LIKE ('%" + Label41.Caption + "%')"
mclass.Refresh
End Sub


Private Sub Label42_Click()
mclass.Refresh
mclass.RecordSource = "SELECT * FROM MCLASS WHERE KODCLASs LIKE ('%" + Label30.Caption + "%')"
mclass.Refresh
End Sub


Private Sub m10_Click()
EmtahanF.Show
End Sub

Private Sub m11_Click()

Karname.Show

End Sub

Private Sub m2_Click()
BankStudent.Show

End Sub

Private Sub m3_Click()
QeybatF.Show

End Sub

Private Sub m4_Click()
Beep

End Sub

Private Sub m6_Click()
ModiriyatCLASS.Show

End Sub

Private Sub m7_Click()
FClassroom.Show

End Sub

Private Sub m8_Click()
FClassroomv.Show

End Sub

Private Sub Start_Click()
If Val(Text10.Text) <= 1 Then Text10.Text = "2"

'«Ì‰ ﬁ”„  ÃœÊ· «Œÿ«— —« Å«ﬂ „Ì ﬂ‰œ

For I = 1 To ekhtar.Recordset.RecordCount
ekhtar.Recordset.Delete
ekhtar.Recordset.MoveNext
Next I
'Å«ﬂ ﬂ—œ‰ ÃœÊ· «Œ «—  „«„ ‘œ

If Check1.Value = 1 Then




 Student.Refresh
 Student.RecordSource = " select * from student where parvande like ('%" + "9" + "%')"
 Text1.Text = Student.Recordset.Fields("Parvande")


 PR1.Visible = True
 
 PR1.Max = Student.Recordset.RecordCount
 For I = 1 To Student.Recordset.RecordCount
 Qeybat.Refresh
 Qeybat.RecordSource = "select * from qeybat where Parvande like ('%" + Text1.Text + "%') and mah like ('%" + Combo1.Text + "%')"
 
  Qeybat.Refresh
 If Qeybat.Recordset.RecordCount >= Val(Text10.Text) Then
ekhtar.Refresh
ekhtar.Recordset.AddNew
'ekhtar.Recordset.Fields("id") = Qeybat.Recordset.Fields("parvande")

ekhtar.Recordset.Fields("parvande") = Qeybat.Recordset.Fields("parvande")
ekhtar.Recordset.Fields("name") = Qeybat.Recordset.Fields("name")
ekhtar.Recordset.Fields("famil") = Qeybat.Recordset.Fields("famil")

ekhtar.Recordset.Fields("tedad") = Qeybat.Recordset.RecordCount
ekhtar.Recordset.Fields("mah") = Qeybat.Recordset.Fields("mah")
'ekhtar.Recordset.Fields("rooz") = Qeybat.Recordset.Fields("rooz")
ekhtar.Recordset.AddNew
ekhtar.Refresh
DataGrid1.Refresh
Else
GoTo 18

End If
18 If Student.Recordset.BOF = True Or Student.Recordset.EOF = True Then

GoTo 108
Else
 
 Student.Recordset.MoveNext
 
 
28 If Student.Recordset.BOF = True Or Student.Recordset.EOF = True Then

GoTo 108
End If

 
 
 
 Text1.Text = Student.Recordset.Fields("parvande")
End If
PR1.Value = PR1.Value + 1

 Next I
108 MsgBox " ⁄„·Ì«  »« „Ê›ﬁÌ  »Â Å«Ì«‰ —”Ìœ" & Chr(10) & " ‰›— »——”Ì ‘œ " & Student.Recordset.RecordCount & Chr$(10) & "  ⁄œ«œ ﬂ”«‰Ì ﬂÂ €Ì»  ¬‰Â« »Ì‘ «“ " + Text10.Text + "  „Ê—œ œ— „«Â   " & Combo1.Text & "  »ÊœÂ  " & ekhtar.Recordset.RecordCount & "  ‰›— „Ì »«‘‰œ  ", vbInformation, "‰ ÌÃÂ ê“«—‘ "

 

  PR1.Visible = False
 PR1.Value = 0
 
Else '„«Â  ÀÌ— ‰œ«—œ





8 Student.Refresh
 Student.RecordSource = " select * from student where parvande like ('%" + "9" + "%')"
 Text1.Text = Student.Recordset.Fields("parvande")


 PR1.Visible = True
 
 PR1.Max = Student.Recordset.RecordCount
 For I = 1 To Student.Recordset.RecordCount
 Qeybat.Refresh
 Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Text1.Text + "%') "
 
  Qeybat.Refresh
 If Qeybat.Recordset.RecordCount >= Val(Text10.Text) Then
ekhtar.Refresh
ekhtar.Recordset.AddNew
'ekhtar.Recordset.Fields("id") = Qeybat.Recordset.Fields("parvande")

ekhtar.Recordset.Fields("parvande") = Qeybat.Recordset.Fields("parvande")
ekhtar.Recordset.Fields("name") = Qeybat.Recordset.Fields("name")
ekhtar.Recordset.Fields("famil") = Qeybat.Recordset.Fields("famil")

ekhtar.Recordset.Fields("tedad") = Qeybat.Recordset.RecordCount
ekhtar.Recordset.Fields("mah") = Qeybat.Recordset.Fields("mah")
'ekhtar.Recordset.Fields("roz") = Qeybat.Recordset.Fields("roz")
ekhtar.Recordset.AddNew
ekhtar.Refresh
DataGrid1.Refresh
Else
GoTo 1

End If
1 If Student.Recordset.BOF = True Or Student.Recordset.EOF = True Then

GoTo 10
Else
 
 Student.Recordset.MoveNext
 
 
2 If Student.Recordset.BOF = True Or Student.Recordset.EOF = True Then

GoTo 10
End If

 
 
 
 Text1.Text = Student.Recordset.Fields("parvande")
End If
PR1.Value = PR1.Value + 1

 Next I
10 MsgBox " ⁄„·Ì«  »« „Ê›ﬁÌ  »Â Å«Ì«‰ —”Ìœ" & Chr(10) & " ‰›— »——”Ì ‘œ " & Student.Recordset.RecordCount & Chr$(10) & "  ⁄œ«œ ﬂ”«‰Ì ﬂÂ €Ì»  ¬‰Â« »Ì‘ «“ " + Text10.Text + " „Ê—œ »ÊœÂ  " & ekhtar.Recordset.RecordCount & "  ‰›— „Ì »«‘‰œ  ", vbInformation, "‰ ÌÃÂ ê“«—‘ "

 

  PR1.Visible = False
 PR1.Value = 0

End If

  End Sub

Private Sub mnuhome_Click()
Entekhab.Show

End Sub

Private Sub Option1_Click()

Text10.Visible = False


End Sub

Private Sub Option2_Click()
Text10.Visible = True
End Sub

Private Sub Text2_Change()
If okod.Value = True Then

ekhtar.Refresh
ekhtar.RecordSource = "select * from ekhtar where parvande like ('%" + Text2 + "%')"
ekhtar.Refresh
DataGrid1.Refresh
End If

'If oostad.Value = True Then


'ekhtar.Refresh
'ekhtar.RecordSource = "select * from ekhtar where teacher like ('%" + Text2 + "%')"
'ekhtar.Refresh
'DataGrid1.Refresh
'End If

If omm.Value = True Then


ekhtar.Refresh
ekhtar.RecordSource = "select * from ekhtar where clas like ('%" + Text2 + "%')"
ekhtar.Refresh
DataGrid1.Refresh
End If

If ofamil.Value = True Then


ekhtar.Refresh
ekhtar.RecordSource = "select * from ekhtar where famil like ('%" + Text2 + "%') or name like ('%" + Text2 + "%')"
ekhtar.Refresh
DataGrid1.Refresh
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
