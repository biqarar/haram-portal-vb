VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Karname 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ò«—‰«„Â"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   13635
   BeginProperty Font 
      Name            =   "B Homa"
      Size            =   9
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Karname.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   13635
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar PB1 
      Height          =   135
      Left            =   9240
      TabIndex        =   90
      Top             =   1680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
      Max             =   10
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "»Â —Ê“ —”«‰Ì"
      Height          =   375
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      Caption         =   "¬„«—"
      Height          =   1335
      Left            =   7080
      TabIndex        =   81
      Top             =   1080
      Width           =   2055
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "ò· ò«—‰«„Â Â«"
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
         Index           =   7
         Left            =   1080
         TabIndex        =   87
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "’«œ— ‰‘œÂ"
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
         Index           =   3
         Left            =   1080
         TabIndex        =   86
         Top             =   600
         Width           =   690
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         Caption         =   "’«œ— ‘œÂ"
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
         Left            =   1080
         TabIndex        =   85
         Top             =   960
         Width           =   630
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   360
         TabIndex        =   84
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
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
         Left            =   360
         TabIndex        =   83
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
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
         Left            =   360
         TabIndex        =   82
         Top             =   960
         Width           =   135
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   9000
      TabIndex        =   70
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   " ⁄œ«œ «„ Õ«‰«  À»   ‘œÂ »—«Ì «Ì‰ ﬁ—¬‰ ¬„Ê“"
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
         Left            =   840
         TabIndex        =   72
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
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
         TabIndex        =   71
         Top             =   240
         Width           =   135
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "’œÊ— ò«—‰«„Â"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "„‘Œ’«  ﬁ—¬‰ ¬„Ê“"
      Height          =   3015
      Left            =   3360
      TabIndex        =   54
      Top             =   120
      Width           =   3615
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Ostad"
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
         Height          =   315
         Left            =   480
         TabIndex        =   68
         Top             =   1680
         Width           =   135
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "‰«„ «” «œ"
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
         Left            =   2280
         TabIndex        =   67
         Top             =   1800
         Width           =   555
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Namepedar"
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
         Height          =   345
         Left            =   480
         TabIndex        =   66
         Top             =   1320
         Width           =   135
      End
      Begin VB.Label Label27 
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
         Height          =   345
         Left            =   480
         TabIndex        =   65
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label26 
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
         Height          =   345
         Left            =   480
         TabIndex        =   64
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label22 
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
         Left            =   480
         TabIndex        =   63
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "‰«„ Åœ—"
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
         Left            =   2280
         TabIndex        =   62
         Top             =   1320
         Width           =   450
      End
      Begin VB.Label Label8 
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
         Left            =   2280
         TabIndex        =   61
         Top             =   960
         Width           =   870
      End
      Begin VB.Label Label42 
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
         Index           =   1
         Left            =   2280
         TabIndex        =   60
         Top             =   600
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
         Index           =   4
         Left            =   2280
         TabIndex        =   59
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label23 
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
         Left            =   2280
         TabIndex        =   58
         Top             =   2160
         Width           =   585
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Tozih"
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
         Height          =   345
         Left            =   480
         TabIndex        =   57
         Top             =   2040
         Width           =   135
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "‘„«—Â  „«”"
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
         Left            =   2280
         TabIndex        =   56
         Top             =   2520
         Width           =   765
      End
      Begin VB.Label Label30 
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
         Height          =   315
         Left            =   480
         TabIndex        =   55
         Top             =   2400
         Width           =   135
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "„‘Œ’«  ò·«”"
      Height          =   3015
      Left            =   120
      TabIndex        =   39
      Top             =   120
      Width           =   3135
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "òœ ò·«”"
         Height          =   330
         Index           =   3
         Left            =   2040
         TabIndex        =   53
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "ÿ—Õ"
         Height          =   330
         Left            =   2040
         TabIndex        =   52
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "„ﬁÿ⁄"
         Height          =   330
         Left            =   2040
         TabIndex        =   51
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "‰«„ «” «œ"
         Height          =   330
         Left            =   2040
         TabIndex        =   50
         Top             =   1440
         Width           =   585
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "”«⁄ "
         Height          =   330
         Left            =   2040
         TabIndex        =   49
         Top             =   1800
         Width           =   405
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "„œ—”"
         Height          =   330
         Left            =   2040
         TabIndex        =   48
         Top             =   2280
         Width           =   405
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
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   240
         TabIndex        =   47
         Top             =   360
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
         TabIndex        =   46
         Top             =   720
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
         TabIndex        =   45
         Top             =   1080
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
         TabIndex        =   44
         Top             =   1800
         Width           =   135
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
         TabIndex        =   43
         Top             =   1440
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
         TabIndex        =   42
         Top             =   2160
         Width           =   225
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
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   " «"
         Height          =   330
         Left            =   720
         TabIndex        =   40
         Top             =   1800
         Width           =   120
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "„‘Œ’«  ﬁ—¬‰ ¬„Ê“"
      Height          =   8175
      Left            =   10440
      TabIndex        =   4
      Top             =   120
      Width           =   3135
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ — »Â »‰œÌ"
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
         Index           =   6
         Left            =   2040
         TabIndex        =   80
         Top             =   7560
         Width           =   990
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "DateRotbe"
         DataSource      =   "Emtahan"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   79
         Top             =   7560
         Width           =   135
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ À» "
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
         Left            =   2040
         TabIndex        =   78
         Top             =   6240
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ ç«Å"
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
         Index           =   5
         Left            =   2040
         TabIndex        =   77
         Top             =   7200
         Width           =   675
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "Ê÷€Ì  ç«Å"
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
         TabIndex        =   76
         Top             =   6720
         Width           =   780
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "D"
         DataSource      =   "Emtahan"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   75
         Top             =   6240
         Width           =   120
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "DateOfChap"
         DataSource      =   "Emtahan"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   74
         Top             =   7200
         Width           =   135
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Chap"
         DataSource      =   "Emtahan"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   73
         Top             =   6720
         Width           =   135
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "KodE"
         DataSource      =   "Emtahan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   60
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Rotbe"
         DataSource      =   "Emtahan"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   3720
         Width           =   135
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Vazeyat"
         DataSource      =   "Emtahan"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   3360
         Width           =   135
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "NimPayan"
         DataSource      =   "Emtahan"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   3000
         Width           =   135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "joze"
         DataSource      =   "Emtahan"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   2640
         Width           =   135
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "ENahaee"
         DataSource      =   "Emtahan"
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
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "TEmtahan"
         DataSource      =   "Emtahan"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Tarh"
         DataSource      =   "Emtahan"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "kodclass"
         DataSource      =   "Emtahan"
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
         Height          =   375
         Left            =   240
         TabIndex        =   26
         ToolTipText     =   "»—«Ì ‰„«Ì‘ „‘Œ’«  ò·«” ò·Ìò ò‰Ìœ"
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Parvande"
         DataSource      =   "Emtahan"
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
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Tozih"
         DataSource      =   "Emtahan"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   5760
         Width           =   135
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Momtahen"
         DataSource      =   "Emtahan"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   5400
         Width           =   135
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "TQeybat"
         DataSource      =   "Emtahan"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   5040
         Width           =   135
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Shafahi"
         DataSource      =   "Emtahan"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   4680
         Width           =   135
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Katbi"
         DataSource      =   "Emtahan"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   4200
         Width           =   135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   ": òœ"
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
         TabIndex        =   19
         Top             =   360
         Width           =   270
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "— »Â ò·«”Ì"
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
         TabIndex        =   18
         Top             =   3840
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ê÷⁄Ì "
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   17
         Top             =   3360
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "„«œÂ «„ Õ«‰"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   16
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "‘„«—Â Ã“¡"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   15
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "«„ Ì«“ ‰Â«ÌÌ"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   14
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ «„ Õ«‰"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   13
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "ÿ—Õ"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   12
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "òœ ò·«”"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   11
         Top             =   1080
         Width           =   615
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
         Height          =   375
         Index           =   1
         Left            =   2040
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label14 
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
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   5880
         Width           =   615
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "„„ Õ‰"
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
         TabIndex        =   8
         Top             =   5400
         Width           =   420
      End
      Begin VB.Label Label19 
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
         Left            =   2040
         TabIndex        =   7
         Top             =   5040
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "‘›«ÂÌ"
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
         Left            =   2040
         TabIndex        =   6
         Top             =   4680
         Width           =   435
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "ò »Ì"
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
         TabIndex        =   5
         Top             =   4320
         Width           =   300
      End
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7080
      TabIndex        =   2
      Top             =   600
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "’œÊ— ò«—‰«„Â"
      Height          =   375
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Motor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   8400
      Visible         =   0   'False
      Width           =   3495
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
         Connect         =   $"Karname.frx":08CA
         OLEDBString     =   $"Karname.frx":0953
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from qeybat"
         Caption         =   "Qeybat"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
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
         Connect         =   $"Karname.frx":09DC
         OLEDBString     =   $"Karname.frx":0A65
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from emtahan"
         Caption         =   "Emtahan"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9.75
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
         Connect         =   $"Karname.frx":0AEE
         OLEDBString     =   $"Karname.frx":0B77
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *  from stu2class"
         Caption         =   "STU2CLASS"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9.75
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
         Connect         =   $"Karname.frx":0C00
         OLEDBString     =   $"Karname.frx":0C89
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from mclass"
         Caption         =   "mclass"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9.75
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
         Connect         =   $"Karname.frx":0D12
         OLEDBString     =   $"Karname.frx":0D9B
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from teacher"
         Caption         =   "teacher"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9.75
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
         Connect         =   $"Karname.frx":0E24
         OLEDBString     =   $"Karname.frx":0EAD
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from student"
         Caption         =   "Student"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9.75
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
         Connect         =   $"Karname.frx":0F36
         OLEDBString     =   $"Karname.frx":0FBF
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from tarhha"
         Caption         =   "Tarhha"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
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
         Connect         =   $"Karname.frx":1048
         OLEDBString     =   $"Karname.frx":10D1
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
   Begin ComctlLib.StatusBar stb1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   35
      Top             =   8385
      Width           =   13635
      _ExtentX        =   24051
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "ò«—»— Ã«—Ì"
            TextSave        =   "ò«—»— Ã«—Ì"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4304
            MinWidth        =   4304
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   " «—ÌŒ «„—Ê“"
            TextSave        =   " «—ÌŒ «„—Ê“"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Karname.frx":115A
      Height          =   5055
      Left            =   120
      TabIndex        =   37
      Top             =   3240
      Visible         =   0   'False
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   8916
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12640511
      DefColWidth     =   80
      HeadLines       =   1
      RowHeight       =   27
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Zar"
         Size            =   9.75
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
      Caption         =   "‰„—«  «„ Õ«‰"
      ColumnCount     =   24
      BeginProperty Column00 
         DataField       =   "Parvande"
         Caption         =   "Å—Ê‰œÂ"
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
         DataField       =   "kodclass"
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
      BeginProperty Column02 
         DataField       =   "Tarh"
         Caption         =   "ÿ—Õ "
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
         DataField       =   "KolMahfozat"
         Caption         =   "ò· „Õ›ÊŸ« "
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
         DataField       =   "MahdodeEmtahan"
         Caption         =   "„ÕœÊœÂ «„ Õ«‰"
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
         DataField       =   "TEmtahan"
         Caption         =   " «—ÌŒ «„ Õ«‰"
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
         DataField       =   "Hefz"
         Caption         =   "Õ›Ÿ"
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
         DataField       =   "Mafahim"
         Caption         =   "„›«ÂÌ„"
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
         DataField       =   "Mostamar"
         Caption         =   "„” „—"
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
         DataField       =   "EE"
         Caption         =   "«» œ« Ê «‰ Â«"
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
         DataField       =   "ENahaee"
         Caption         =   "«„ Ì«“ ‰Â«ÌÌ"
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
         DataField       =   "OP"
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
      BeginProperty Column12 
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
      BeginProperty Column13 
         DataField       =   "Momtahen"
         Caption         =   "„„ Õ‰"
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
         DataField       =   "TQeybat"
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
      BeginProperty Column15 
         DataField       =   "joze"
         Caption         =   "Ã“¡"
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
         DataField       =   "NimPayan"
         Caption         =   "‰Ì„Â Ê Å«Ì«‰ Ã“¡"
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
         DataField       =   "KasrEmtiaz"
         Caption         =   "ò”— «„ Ì«“"
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
         DataField       =   "Katbi"
         Caption         =   "ò »Ì"
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
         DataField       =   "Shafahi"
         Caption         =   "‘›«ÂÌ"
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
         DataField       =   "Vazeyat"
         Caption         =   "Ê÷⁄Ì "
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
         DataField       =   "Rotbe"
         Caption         =   "— »Â ò·«”Ì"
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
         DataField       =   "tozih"
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
      BeginProperty Column23 
         DataField       =   "KodE"
         Caption         =   "òœ «„ Õ«‰"
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
         BeginProperty Column12 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   1214.929
         EndProperty
      EndProperty
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   495
      Left            =   7080
      TabIndex        =   38
      Top             =   2640
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   873
      Style           =   1
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "ﬁ—¬‰ ¬„Ê“«‰"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "‰„—«  À»  ‘œÂ"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "‰„—«  ﬁ—«∆ "
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Karname.frx":1170
      Height          =   5055
      Left            =   120
      TabIndex        =   36
      Top             =   3240
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   8916
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12648384
      DefColWidth     =   80
      HeadLines       =   1
      RowHeight       =   27
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Zar"
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
      Caption         =   "ﬁ—¬‰ ¬„Ê“«‰"
      ColumnCount     =   27
      BeginProperty Column00 
         DataField       =   "Parvande"
         Caption         =   "Å—Ê‰œÂ"
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
         Caption         =   "‘„«—Â ê–— ‰«„Â"
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
         Caption         =   "‘„«—Â Â„—«Â"
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
         Caption         =   "«”ò‰"
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
         Caption         =   "Clas1"
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
      BeginProperty Column21 
         DataField       =   "Tarikh"
         Caption         =   "Tarikh"
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
         Caption         =   "Clas2"
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
         Caption         =   "Clas3"
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
         Caption         =   "Clas4"
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
         Caption         =   "Clas5"
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
         Caption         =   "NF"
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
         BeginProperty Column12 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column18 
            Object.Visible         =   0   'False
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column22 
            Object.Visible         =   0   'False
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column23 
            Object.Visible         =   0   'False
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column24 
            Object.Visible         =   0   'False
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column25 
            Object.Visible         =   0   'False
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column26 
            Object.Visible         =   0   'False
            ColumnWidth     =   1214.929
         EndProperty
      EndProperty
   End
   Begin VB.Label RokhaniKarname 
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "ò«—‰«„Â œÊ—Â —Ê ŒÊ«‰Ì"
      Height          =   330
      Left            =   12000
      TabIndex        =   89
      Top             =   0
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Ã” ÃÊ œ— ‰«„ Ê ‰«„ Œ«‰Ê«œêÌ Ê ‘„«—Â Å—Ê‰œÂ"
      Height          =   330
      Left            =   7560
      TabIndex        =   3
      Top             =   240
      Width           =   2790
   End
   Begin VB.Menu mnuhome 
      Caption         =   "#"
   End
   Begin VB.Menu MNUKARNAMNR 
      Caption         =   "Ê÷⁄Ì  ﬁ—¬‰ ¬„Ê“"
      Begin VB.Menu MNUcHARTj30 
         Caption         =   "ç«Å Ê÷⁄Ì  ﬁ—¬‰ ¬„Ê“"
         Begin VB.Menu mnucolor 
            Caption         =   "ç«Å —‰êÌ"
         End
         Begin VB.Menu mnubw 
            Caption         =   "ç«Å ”Ì«Â Ê ”›Ìœ"
         End
         Begin VB.Menu erer 
            Caption         =   "-"
            Index           =   10
         End
         Begin VB.Menu mnuchartjadid 
            Caption         =   "ç«Å Ê÷⁄Ì  ÃœÌœ"
         End
      End
   End
   Begin VB.Menu mnusettingemtahan 
      Caption         =   " ‰ŸÌ„« "
   End
End
Attribute VB_Name = "Karname"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()

Emtahan.Refresh
Emtahan.RecordSource = "select * from Emtahan where chap like ('%" & "" & "%')"
Emtahan.Refresh
Label58.Caption = Emtahan.Recordset.RecordCount

Emtahan.Refresh
Emtahan.RecordSource = "select * from Emtahan where chap like ('%" & "ç«Å ‰‘œÂ" & "%')"
Emtahan.Refresh
Label57.Caption = Emtahan.Recordset.RecordCount


'À»  ç«Å ‘œÂ Â«
Emtahan.Refresh
Emtahan.RecordSource = "select * from Emtahan where chap like ('%" & "ç«Å ‘œÂ" & "%')"
Emtahan.Refresh
Label56.Caption = Emtahan.Recordset.RecordCount



End Sub

Private Sub Command1_Click()
'\\yafatemeh2-pc\F\Markaz Quran & Hadis\FORMXLS\KarnameJadid.xlsx
If Entekhab.SB.Panels(1).Text = "»—‰«„Â ‰ÊÌ”" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "karname-print" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
PB1.Value = 1



'ÃÂ  «Ì‰òÂ ›—«„Ê‘ ‰ò‰œ òœ«„ ò«—‰«„Â —« »«Ìœ ç«Å ò‰œ »«Ìœ òœ «‰Õ’«—Ì ¬‰ —« »Â Œ«ÿ— »”Å«—œ
Dim KodEnhesariPrint As String


KodEnhesariPrint = "EMPTY"



KodEnhesariPrint = Emtahan.Recordset.Fields("kode")
'œ— ç«Å ò«—‰«„Â œ— «‰ Â«Ì òœ »Â œ—œ „Ì ŒÊ—œ

'On Error GoTo 9898
GoTo 9999
9898:
MsgBox "«‘ò«· œ— ç«Å ò«—‰«„Â" & Chr$(10) & "„„ò‰ «”  «Ì‰ «‘ò«· »Â ÌòÌ «“ ⁄·· “Ì— »«‘œ" & Chr$(10) & " ‰ÿÌ„«  ›«Ì· Œ—ÊÃÌ ’ÕÌÕ ‰„Ì »«‘œ" & Chr$(10) & "‰„—Â «Ì »—«Ì ﬁ—¬‰ ¬„Ê“ À»  ‰‘œÂ »«‘œ" & Chr$(10) & "«ÿ·«⁄«  Õ”«” „À· ‘„«—Â Å—Ê‰œÂ ﬁ—¬‰ ¬„Ê“ œ” ò«—Ì ‘œÂ »«‘œ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub

9999:

If Me.Emtahan.Recordset.Fields("Chap") = "ç«Å ‘œÂ" Then
    If MsgBox(" «Ì‰ ò«—‰«„Â ﬁ»·« œ—  «—ÌŒ " & Me.Emtahan.Recordset.Fields("dateofchap") & " ç«Å ‘œÂ «”  ¬Ì« „Ì ŒÊ«ÂÌœ œÊ»«—Â ¬‰ —« ç«Å ò‰Ìœ  ", vbExclamation + vbYesNo, "ç«Å ò«—‰«„Â") = vbYes Then
    GoTo 500
    Else
    Exit Sub
    End If
    
End If
500
Dim ASD As String
Dim oExcel As Object
'Dim oBook As Object
'Dim oSheet As Object
Dim AD, ParvandeQuranAmooZ As String
PB1.Value = 2
'On Error GoTo 1
GoTo 2
1: MsgBox "›«Ì· ÅÌœ« ‰‘œ∫ »—«Ì «Ã—«Ì «Ì‰ ›—„«‰ »«Ìœ «“ ”Ì” „ ”—Ê— «” ›«œÂ ò‰Ìœ", vbCritical, "Œÿ«"
Exit Sub

2:

If Entekhab.Pc.Checked = True Then
'Set oExcel = GetObject("F:\Markaz Quran & Hadis\FORMXLS\CopyofKarnameJadid.xlsx")
Set oExcel = GetObject(Entekhab.AdressLabel.Caption & "CopyofKarnameJadid.xlsx")
PB1.Value = 3
End If

If Entekhab.net.Checked = True Then
Set oExcel = GetObject(Entekhab.NetAdresslabel.Caption & "CopyofKarnameJadid.xlsx")
PB1.Value = 4
End If


'Set oExcel = GetObject("d:\vadiexls.xlsx")
oExcel.ActiveSheet.Range("f3").Value = Taqvim.Tarikh.Caption

'ç” ÊÃÊ »—«Ì À»  ‰«„ Ê „‘Œ’«  ›—œÌ ﬁ—¬ ‰ ¬„Ê“
ParvandeQuranAmooZ = Emtahan.Recordset.Fields("parvande")
Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" & ParvandeQuranAmooZ & "%')"
Student.Refresh
PB1.Value = 5
''''''''''''''''''''''
'ﬁ»  „‘Œ’«  ›—œÌ œ— ò«—‰„«„Â

                oExcel.ActiveSheet.Range("c2").Value = Student.Recordset.Fields("name") & " " & Student.Recordset.Fields("famil")
                oExcel.ActiveSheet.Range("c3").Value = Student.Recordset.Fields("parvande")
'Ã” ÊÃÊ »—«Ì À»  „‘Œ’«  ò·«”·

mclass.Refresh
mclass.RecordSource = "select * from mclass where kodclass like ('%" + Emtahan.Recordset.Fields("kodclass") + "%')"
mclass.Refresh

PB1.Value = 6
'À»  „‘Œ’«   ò·«”
                oExcel.ActiveSheet.Range("f2").Value = mclass.Recordset.Fields("ostad")
                oExcel.ActiveSheet.Range("g3").Value = mclass.Recordset.Fields("tarh")




'À»  ‰Ì„Â Ê Å«Ì«‰ Ã“∆ »Êœ‰ œ— ò«—‰«„Â

                oExcel.ActiveSheet.Range("a5").Value = Emtahan.Recordset.Fields("nimpayan") & " " & Emtahan.Recordset.Fields("joze")

'À»  ‰„—«  œ— ò«—‰«„Â

                oExcel.ActiveSheet.Range("b5").Value = Emtahan.Recordset.Fields("hefz")
                oExcel.ActiveSheet.Range("c5").Value = Emtahan.Recordset.Fields("ee")
                oExcel.ActiveSheet.Range("e5").Value = Emtahan.Recordset.Fields("mostamar")
                oExcel.ActiveSheet.Range("d5").Value = Emtahan.Recordset.Fields("mafahim")
'oExcel.ActiveSheet.Range("f4").Value = Emtahan.Recordset.Fields("kasremtiaz")

'‰„—Â ‰Â«ÌÌ
                oExcel.ActiveSheet.Range("g5").Value = Emtahan.Recordset.Fields("enahaee")

PB1.Value = 7
'À»  Ê÷⁄Ì  ﬁ»Ê·Ì Ì«  ÃÌœÌœ

'À»   ⁄œ«œ €Ì»  Â«
                oExcel.ActiveSheet.Range("f5").Value = Emtahan.Recordset.Fields("tqeybat")

'À»  — »Â ò·«”Ì
                oExcel.ActiveSheet.Range("e6").Value = Emtahan.Recordset.Fields("rotbe")

Dim SearchKarname, OPR, NimPayanJozeSabeq, JozeSabeq As String
'ÃÂ  À»  ‰„—«   çœÌœÌ ﬁ»·Ì òÂ ‘«„·  ÃœÌœÌ 1 Ê  ÃœÌœÌ 2 „Ì »«‘œ

'
'If Option1.Value = True Then OPR = "NP5"
'If Option2.Value = True Then OPR = "NP1"
' «“ Â„«‰ ›Ì·œ Ì „Ì ŒÊ«‰œ òÂ «‰ Œ«»‘œÂ Ê ﬁ—«— «”  ò«—‰«„Â »—«Ì ¬‰ ’«œ— ‘Êœ
If Emtahan.Recordset.Fields("nimpayan") = "‰Ì„Â Ã“¡" Then OPR = "NP5"
If Emtahan.Recordset.Fields("nimpayan") = "Å«Ì«‰ Ã“¡" Then OPR = "NP1"

SearchKarname = "P" & Emtahan.Recordset.Fields("parvande") & "J" & Emtahan.Recordset.Fields("joze") & OPR & "T1"
'
Emtahan.Refresh
Emtahan.RecordSource = "select * from Emtahan where kode like ('%" & SearchKarname & "%')"
Emtahan.Refresh
If Emtahan.Recordset.BOF = True Or Emtahan.Recordset.EOF = True Then
GoTo 4
Else
'
oExcel.ActiveSheet.Range("g6").Value = Emtahan.Recordset.Fields("enahaee")
End If
4:
'
'çÊ‰ Ìò »«— «“ œÌ « »Ì” òÊ∆—Ì ê—› Â ‘œÂ «”  »«Ìœ œÊ»«—Â »— ”— Ã«Ì ŒÊœ‘ »—ê—œœ

Emtahan.Refresh
Emtahan.RecordSource = "select * from Emtahan where kode like ('%" & KodEnhesariPrint & "%')"
Emtahan.Refresh
'Å«Ì«‰ »—ê‘  »Â Ã«Ì ﬁ»·Ì

PB1.Value = 8

If Emtahan.Recordset.Fields("nimpayan") = "‰Ì„Â Ã“¡" Then OPR = "NP5"
If Emtahan.Recordset.Fields("nimpayan") = "Å«Ì«‰ Ã“¡" Then OPR = "NP1"


'Ã” ÊÃÊ »—«Ì «Ì‰òÂ  ÃœÌœÌ œÊ„ ÊÃÊœ ‰œ«‘ Â »«‘œ
SearchKarname = "P" & Emtahan.Recordset.Fields("parvande") & "J" & Emtahan.Recordset.Fields("joze") & OPR & "T1"
'
Emtahan.Refresh
Emtahan.RecordSource = "select * from Emtahan where kode like ('%" & SearchKarname & "%')"
Emtahan.Refresh
'
If Emtahan.Recordset.BOF = True Or Emtahan.Recordset.EOF = True Then
GoTo 3
Else
'
oExcel.ActiveSheet.Range("g6").Value = Emtahan.Recordset.Fields("enahaee")
End If
3:

' À»  ‰„—Â ‰Ì„ Ã“¡ ”«»ﬁ  Ê”ÿ ”Ì” „
Emtahan.Refresh
Emtahan.RecordSource = "select * from Emtahan where kode like ('%" & KodEnhesariPrint & "%')"
Emtahan.Refresh

If Emtahan.Recordset.Fields("nimpayan") = "‰Ì„Â Ã“¡" Then
NimPayanJozeSabeq = "Å«Ì«‰ Ã“¡"
JozeSabeq = Val(Emtahan.Recordset.Fields("joze")) - 1

End If

If Emtahan.Recordset.Fields("nimpayan") = "Å«Ì«‰ Ã“¡" Then
NimPayanJozeSabeq = "‰Ì„Â Ã“¡"
JozeSabeq = Val(Emtahan.Recordset.Fields("joze"))
End If


Emtahan.Refresh
Emtahan.RecordSource = "select * from Emtahan where nimpayan like ('%" & NimPayanJozeSabeq & "%') and joze like ('%" & JozeSabeq & "%')and parvande like ('%" & Emtahan.Recordset.Fields("parvande") & "%')"
Emtahan.Refresh

If Emtahan.Recordset.BOF = True Or Emtahan.Recordset.EOF = True Then
oExcel.ActiveSheet.Range("c6").Value = "-"
Else

oExcel.ActiveSheet.Range("c6").Value = Emtahan.Recordset.Fields("enahaee")
End If


PB1.Value = 9
'Å«Ì«‰ À»  ‰„—Â ‰Ì„ Ã“¡ ”«»ﬁ
'À»   Ê÷ÌÕ« 
Emtahan.Refresh
Emtahan.RecordSource = "select * from Emtahan where kode like ('%" & KodEnhesariPrint & "%')"
Emtahan.Refresh
'ç“¡ 30

If Val(Emtahan.Recordset.Fields("joze")) = 30 Then
    If Val(Emtahan.Recordset.Fields("enahaee")) >= 19 And Val(Emtahan.Recordset.Fields("enahaee")) <= 20 Then
    oExcel.ActiveSheet.Range("d7").Value = "»« ò„«·  ‘ò— «“ “Õ„«  ‘„«° ò”» ‰„—Â ⁄«·Ì  Ê”ÿ ‘„« „«ÌÂ œ·ê—„Ì „«” "
    oExcel.ActiveSheet.Range("a8").Value = "⁄«·Ì"

    GoTo 110
    
    End If
    
        If Val(Emtahan.Recordset.Fields("enahaee")) >= 18 And Val(Emtahan.Recordset.Fields("enahaee")) < 19 Then
    oExcel.ActiveSheet.Range("d7").Value = "»« ò„«·  ‘ò— «“ “Õ„«  ‘„«° «„ÌœÊ«—Ì„ »« „Ê›ﬁÌ  »Ì‘ — »Â „”Ì— ŒÊœ «œ«„Â œÂÌœ »Â «„Ìœ ò”» ‰„—Â ⁄«·Ì œ— «„ Õ«‰ »⁄œÌ"
        oExcel.ActiveSheet.Range("a8").Value = "ŒÊ»"

        GoTo 110

    End If
    
        If Val(Emtahan.Recordset.Fields("enahaee")) >= 17 And Val(Emtahan.Recordset.Fields("enahaee")) < 18 Then
    oExcel.ActiveSheet.Range("d7").Value = "»« ò„«·  ‘ò— «“ “Õ„«  ‘„«° «„ÌœÊ«—Ì„ »« „Ê›ﬁÌ  »Ì‘ — »Â „”Ì— ŒÊœ «œ«„Â œÂÌœ° Â‰Ê“ Ã«Ì ÅÌ‘—›  œ«—Ìœ"
           oExcel.ActiveSheet.Range("a8").Value = "„ Ê”ÿ"
 
        GoTo 110

    End If
    
        If Val(Emtahan.Recordset.Fields("enahaee")) <= 17 Then
            oExcel.ActiveSheet.Range("d7").Value = "·ÿ›«  ·«‘ ŒÊœ —« »—«Ì ò”» ‰„—«  »Â — œÊç‰œ«‰ ò‰Ìœ° Ê÷⁄Ì  «‰ —÷«Ì  »Œ‘ ‰Ì” °  «—ÌŒ «„ Õ«‰ „Ãœœ ‘„«           /   /1390 „Ì »«‘œ"
                   oExcel.ActiveSheet.Range("a8").Value = "÷⁄Ì›"
 
                GoTo 110

    End If
    
End If





'ç“¡ 30

'“Ì— 10 ç“¡

If Val(Emtahan.Recordset.Fields("joze")) >= 1 And Val(Emtahan.Recordset.Fields("joze")) <= 10 Then
    If Val(Emtahan.Recordset.Fields("enahaee")) >= 19 And Val(Emtahan.Recordset.Fields("enahaee")) <= 20 Then
    oExcel.ActiveSheet.Range("d7").Value = "»« ò„«·  ‘ò— «“ “Õ„«  ‘„«° ò”» ‰„—Â ⁄«·Ì  Ê”ÿ ‘„« „«ÌÂ œ·ê—„Ì „«” "
    oExcel.ActiveSheet.Range("a8").Value = "⁄«·Ì"

    GoTo 110
    
    End If
    
        If Val(Emtahan.Recordset.Fields("enahaee")) >= 18 And Val(Emtahan.Recordset.Fields("enahaee")) < 19 Then
    oExcel.ActiveSheet.Range("d7").Value = "»« ò„«·  ‘ò— «“ “Õ„«  ‘„«° «„ÌœÊ«—Ì„ »« „Ê›ﬁÌ  »Ì‘ — »Â „”Ì— ŒÊœ «œ«„Â œÂÌœ »Â «„Ìœ ò”» ‰„—Â ⁄«·Ì œ— «„ Õ«‰ »⁄œÌ"
        oExcel.ActiveSheet.Range("a8").Value = "ŒÊ»"

        GoTo 110

    End If
    
        If Val(Emtahan.Recordset.Fields("enahaee")) >= 17 And Val(Emtahan.Recordset.Fields("enahaee")) < 18 Then
    oExcel.ActiveSheet.Range("d7").Value = "»« ò„«·  ‘ò— «“ “Õ„«  ‘„«° «„ÌœÊ«—Ì„ »« „Ê›ﬁÌ  »Ì‘ — »Â „”Ì— ŒÊœ «œ«„Â œÂÌœ° Â‰Ê“ Ã«Ì ÅÌ‘—›  œ«—Ìœ"
           oExcel.ActiveSheet.Range("a8").Value = "„ Ê”ÿ"
 
        GoTo 110

    End If
    
        If Val(Emtahan.Recordset.Fields("enahaee")) <= 17 Then
            oExcel.ActiveSheet.Range("d7").Value = "·ÿ›«  ·«‘ ŒÊœ —« »—«Ì ò”» ‰„—«  »Â — œÊç‰œ«‰ ò‰Ìœ° Ê÷⁄Ì  «‰ —÷«Ì  »Œ‘ ‰Ì” °  «—ÌŒ «„ Õ«‰ „Ãœœ ‘„«           /   /1390 „Ì »«‘œ"
                   oExcel.ActiveSheet.Range("a8").Value = "÷⁄Ì›"
 
                GoTo 110

    End If
    
End If
'«“ 10  « 20 Ã“¡√
If Val(Emtahan.Recordset.Fields("joze")) >= 11 And Val(Emtahan.Recordset.Fields("joze")) <= 20 Then
    If Val(Emtahan.Recordset.Fields("enahaee")) >= 19 And Val(Emtahan.Recordset.Fields("enahaee")) <= 20 Then
    oExcel.ActiveSheet.Range("d7").Value = "»« ò„«·  ‘ò— «“ “Õ„«  ‘„«° ò”» ‰„—Â ⁄«·Ì  Ê”ÿ ‘„« „«ÌÂ œ·ê—„Ì „«” "
        oExcel.ActiveSheet.Range("a8").Value = "⁄«·Ì"

        GoTo 110

    End If
    
        If Val(Emtahan.Recordset.Fields("enahaee")) >= 17 And Val(Emtahan.Recordset.Fields("enahaee")) < 19 Then
    oExcel.ActiveSheet.Range("d7").Value = "»« ò„«·  ‘ò— «“ “Õ„«  ‘„«° «„ÌœÊ«—Ì„ »« „Ê›ﬁÌ  »Ì‘ — »Â „”Ì— ŒÊœ «œ«„Â œÂÌœ »Â «„Ìœ ò”» ‰„—Â ⁄«·Ì œ— «„ Õ«‰ »⁄œÌ"
           oExcel.ActiveSheet.Range("a8").Value = "ŒÊ»"
 
        GoTo 110

    End If
    
        If Val(Emtahan.Recordset.Fields("enahaee")) >= 16 And Val(Emtahan.Recordset.Fields("enahaee")) < 17 Then
    oExcel.ActiveSheet.Range("d7").Value = "»« ò„«·  ‘ò— «“ “Õ„«  ‘„«° «„ÌœÊ«—Ì„ »« „Ê›ﬁÌ  »Ì‘ — »Â „”Ì— ŒÊœ «œ«„Â œÂÌœ° Â‰Ê“ Ã«Ì ÅÌ‘—›  œ«—Ìœ"
        oExcel.ActiveSheet.Range("a8").Value = "„ Ê”ÿ"

        GoTo 110

    End If
    
        If Val(Emtahan.Recordset.Fields("enahaee")) <= 16 Then
            oExcel.ActiveSheet.Range("d7").Value = "·ÿ›«  ·«‘ ŒÊœ —« »—«Ì ò”» ‰„—«  »Â — œÊç‰œ«‰ ò‰Ìœ° Ê÷⁄Ì  «‰ —÷«Ì  »Œ‘ ‰Ì” °  «—ÌŒ «„ Õ«‰ „Ãœœ ‘„«           /   /1390 „Ì »«‘œ"
                   oExcel.ActiveSheet.Range("a8").Value = "÷⁄Ì›"
 
                GoTo 110

    End If
    
End If
'»«·«Ì 20 Ã“¡
If Val(Emtahan.Recordset.Fields("joze")) >= 21 And Val(Emtahan.Recordset.Fields("joze")) < 30 Then
    If Val(Emtahan.Recordset.Fields("enahaee")) >= 18.5 And Val(Emtahan.Recordset.Fields("enahaee")) <= 20 Then
    oExcel.ActiveSheet.Range("d7").Value = "»« ò„«·  ‘ò— «“ “Õ„«  ‘„«° ò”» ‰„—Â ⁄«·Ì  Ê”ÿ ‘„« „«ÌÂ œ·ê—„Ì „«” "
        oExcel.ActiveSheet.Range("a8").Value = "⁄«·Ì"

        GoTo 110

    End If
    
        If Val(Emtahan.Recordset.Fields("enahaee")) >= 17 And Val(Emtahan.Recordset.Fields("enahaee")) < 18.5 Then
    oExcel.ActiveSheet.Range("d7").Value = "»« ò„«·  ‘ò— «“ “Õ„«  ‘„«° «„ÌœÊ«—Ì„ »« „Ê›ﬁÌ  »Ì‘ — »Â „”Ì— ŒÊœ «œ«„Â œÂÌœ »Â «„Ìœ ò”» ‰„—Â ⁄«·Ì œ— «„ Õ«‰ »⁄œÌ"
        oExcel.ActiveSheet.Range("a8").Value = "ŒÊ»"

        GoTo 110

    End If
    
        If Val(Emtahan.Recordset.Fields("enahaee")) >= 15 And Val(Emtahan.Recordset.Fields("enahaee")) < 17 Then
    oExcel.ActiveSheet.Range("d7").Value = "»« ò„«·  ‘ò— «“ “Õ„«  ‘„«° «„ÌœÊ«—Ì„ »« „Ê›ﬁÌ  »Ì‘ — »Â „”Ì— ŒÊœ «œ«„Â œÂÌœ° Â‰Ê“ Ã«Ì ÅÌ‘—›  œ«—Ìœ"
        oExcel.ActiveSheet.Range("a8").Value = "„ Ê”ÿ"

        GoTo 110

    End If
    
        If Val(Emtahan.Recordset.Fields("enahaee")) <= 15 Then
            oExcel.ActiveSheet.Range("d7").Value = "·ÿ›«  ·«‘ ŒÊœ —« »—«Ì ò”» ‰„—«  »Â — œÊç‰œ«‰ ò‰Ìœ° Ê÷⁄Ì  «‰ —÷«Ì  »Œ‘ ‰Ì” °  «—ÌŒ «„ Õ«‰ „Ãœœ ‘„«           /   /1390 „Ì »«‘œ"
                oExcel.ActiveSheet.Range("a8").Value = "÷⁄Ì›"

                GoTo 110

    End If
    
End If
'Å«Ì«‰ À»   Ê÷ÌÕ« 
'KodEnhesariPrint = Emtahan.Recordset.Fields("kode")

110:

Emtahan.Refresh
Emtahan.RecordSource = "select * from Emtahan where kode like ('%" & KodEnhesariPrint & "%')"
Emtahan.Refresh

'À»  ç«Å ò«—‰«„Â
Emtahan.Recordset.Fields("Chap") = "ç«Å ‘œÂ"
Emtahan.Recordset.Fields("dateofChap") = Me.stb1.Panels(3).Text
Emtahan.Recordset.Update
Emtahan.Refresh

'œ— ç«Å ò«—‰«„Â œ— «‰ Â«Ì òœ »Â œ—œ „Ì ŒÊ—œ
PB1.Value = 10





MsgBox "ò«—‰«„Â ¬„«œÂ ç«Å „Ì »«‘œ", vbInformation + vbOKOnly, "ç«Å ò«—‰«„Â"



oExcel.Application.Visible = True
On Error GoTo 722


oExcel.Parent.Windows(2).Visible = True
GoTo 910
722:

oExcel.Parent.Windows(1).Visible = True
910:
''''''

oExcel.SaveAs KodEnhesariPrint
'oExcel.Close
'
'
'Set oExcel = Nothing ' Remove object variable.
''''''''
'Shell "Explorer.exe " & "c:\" & KodEnhesariPrint & ".xlsx"

'oExcel.FileRef.Visible = True
'Catch
'»Â —Ê“ —”«‰Ì ¬„«— „ÊÃÊœ œ— ›—„
Call Command2_Click

PB1.Value = 0
End Sub

Private Sub Command3_Click()
If Entekhab.SB.Panels(1).Text = "„ÌÂ„«‰" Then Exit Sub
'If Option1.Value = False And Option2.Value = False Then Exit Sub

Dim ASD As String
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim AD As String
'On Error GoTo 1
GoTo 2
1: MsgBox "›«Ì· ÅÌœ« ‰‘œ∫ »—«Ì «Ã—«Ì «Ì‰ ›—„«‰ »«Ìœ «“ ”Ì” „ ”—Ê— «” ›«œÂ ò‰Ìœ", vbCritical, "Œÿ«"
Exit Sub

2:
'ÃÂ  «Ì‰òÂ ›—«„Ê‘ ‰ò‰œ òœ«„ ò«—‰«„Â —« »«Ìœ ç«Å ò‰œ »«Ìœ òœ «‰Õ’«—Ì ¬‰ —« »Â Œ«ÿ— »”Å«—œ
Dim KodEnhesariPrint As String
KodEnhesariPrint = Emtahan.Recordset.Fields("kode")
'œ— ç«Å ò«—‰«„Â œ— «‰ Â«Ì òœ »Â œ—œ „Ì ŒÊ—œ




Set oExcel = GetObject(Entekhab.AdressLabel.Caption & "KArname-Hefz.xlsx")
'Set oExcel = GetObject("d:\vadiexls.xlsx")
oExcel.ActiveSheet.Range("g1").Value = Taqvim.Tarikh.Caption

'ç” ÊÃÊ »—«Ì À»  ‰«„ Ê „‘Œ’«  ›—œÌ ﬁ—¬ ‰ ¬„Ê“

Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + Emtahan.Recordset.Fields("parvande") + "%')"
Student.Refresh

''''''''''''''''''''''
'ﬁ»  „‘Œ’«  ›—œÌ œ— ò«—‰„«„Â

oExcel.ActiveSheet.Range("b2").Value = Student.Recordset.Fields("name") & " " & Student.Recordset.Fields("famil")

'Ã” ÊÃÊ »—«Ì À»  „‘Œ’«  ò·«”·

mclass.Refresh
mclass.RecordSource = "select * from mclass where kodclass like ('%" + Emtahan.Recordset.Fields("kodclass") + "%')"
mclass.Refresh


'À»  „‘Œ’«   ò·«”
oExcel.ActiveSheet.Range("d2").Value = "«” «œ:" & " " & mclass.Recordset.Fields("ostad")
oExcel.ActiveSheet.Range("g2").Value = mclass.Recordset.Fields("tarh")




'À»  ‰Ì„Â Ê Å«Ì«‰ Ã“∆ »Êœ‰ œ— ò«—‰«„Â

oExcel.ActiveSheet.Range("a4").Value = Emtahan.Recordset.Fields("nimpayan") & " " & Emtahan.Recordset.Fields("joze")

'À»  ‰„—«  œ— ò«—‰«„Â

oExcel.ActiveSheet.Range("b4").Value = Emtahan.Recordset.Fields("hefz")
oExcel.ActiveSheet.Range("c4").Value = Emtahan.Recordset.Fields("ee")
oExcel.ActiveSheet.Range("d4").Value = Emtahan.Recordset.Fields("mostamar")
oExcel.ActiveSheet.Range("e4").Value = Emtahan.Recordset.Fields("mafahim")
oExcel.ActiveSheet.Range("f4").Value = Emtahan.Recordset.Fields("kasremtiaz")

'‰„—Â ‰Â«ÌÌ
oExcel.ActiveSheet.Range("g4").Value = Emtahan.Recordset.Fields("enahaee")


'À»  Ê÷⁄Ì  ﬁ»Ê·Ì Ì«  ÃÌœÌœ
oExcel.ActiveSheet.Range("b7").Value = Emtahan.Recordset.Fields("vazeyat")

'À»   ⁄œ«œ €Ì»  Â«
oExcel.ActiveSheet.Range("a7").Value = Emtahan.Recordset.Fields("tqeybat")

'À»  — »Â ò·«”Ì
oExcel.ActiveSheet.Range("c7").Value = Emtahan.Recordset.Fields("rotbe")

Dim SearchKarname, OPR As String
'ÃÂ  À»  ‰„—«   çœÌœÌ ﬁ»·Ì òÂ ‘«„·  ÃœÌœÌ 1 Ê  ÃœÌœÌ 2 „Ì »«‘œ

'
'If Option1.Value = True Then OPR = "NP5"
'If Option2.Value = True Then OPR = "NP1"
' «“ Â„«‰ ›Ì·œ Ì „Ì ŒÊ«‰œ òÂ «‰ Œ«»‘œÂ Ê ﬁ—«— «”  ò«—‰«„Â »—«Ì ¬‰ ’«œ— ‘Êœ
If Emtahan.Recordset.Fields("nimpayan") = "‰Ì„Â Ã“¡" Then OPR = "NP5"
If Emtahan.Recordset.Fields("nimpayan") = "Å«Ì«‰ Ã“¡" Then OPR = "NP1"

SearchKarname = "P" & Label22.Caption & "J" & Label5.Caption & OPR & "T1"
'
Emtahan.Refresh
Emtahan.RecordSource = "select * from Emtahan where kode like ('%" & SearchKarname & "%')"
Emtahan.Refresh
If Emtahan.Recordset.BOF = True Or Emtahan.Recordset.EOF = True Then
GoTo 4
Else
'
oExcel.ActiveSheet.Range("b5").Value = Emtahan.Recordset.Fields("enahaee")
End If
4:
'
'If Option1.Value = True Then OPR = "NP5"
'If Option2.Value = True Then OPR = "NP1"
'Ã” ÊÃÊ »—«Ì «Ì‰òÂ  ÃœÌœÌ œÊ„ ÊÃÊœ ‰œ«‘ Â »«‘œ
SearchKarname = "P" & Label22.Caption & "J" & Label5.Caption & OPR & "T2"
'
Emtahan.Refresh
Emtahan.RecordSource = "select * from Emtahan where kode like ('%" & SearchKarname & "%')"
Emtahan.Refresh
'
If Emtahan.Recordset.BOF = True Or Emtahan.Recordset.EOF = True Then
GoTo 3
Else
'
oExcel.ActiveSheet.Range("c5").Value = Emtahan.Recordset.Fields("enahaee")
End If
3:
'
MsgBox "ò«—‰«„Â ¬„«œÂ ç«Å „Ì »«‘œ", vbInformation + vbOKOnly, "ç«Å ò«—‰«„Â"
'

'
oExcel.SaveAs KodEnhesariPrint

oExcel.Application.Visible = True
oExcel.Parent.Windows(1).Visible = True
'
'






End Sub

Private Sub DataGrid2_DblClick()
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

Private Sub Form_Load()




Me.stb1.Panels(1).Text = user.OP.Text
Me.stb1.Panels(3).Text = Taqvim.Tarikh.Caption

End Sub

Private Sub Form_Unload(Cancel As Integer)
Entekhab.Show

End Sub

Private Sub Text1_Change()




End Sub

Private Sub Label22_Change()
Exit Sub

Emtahan.Refresh
Emtahan.RecordSource = "select * from Emtahan where parvande like ('%" & Label22.Caption & "%')"  ' or kode like ('%" & Text2.Text & "%') or kodclass like ('%" & Text2.Text & "%')"
Emtahan.Refresh
Label51.Caption = Emtahan.Recordset.RecordCount

End Sub

Private Sub Label37_Click()
mclass.Refresh
mclass.RecordSource = "select * from mclass where kodclass like ('%" & Label37.Caption & "%')"
mclass.Refresh


End Sub

Private Sub Label38_Change()
'Student.Refresh
'Student.RecordSource = "select * from student where parvande like ('%" + Label38.Caption + "%') "
'Student.Refresh
End Sub

Private Sub Option2_Click()

End Sub

Private Sub Label53_Click()

End Sub

Private Sub Label38_Click()
Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + Label38.Caption + "%') "
Student.Refresh
End Sub

Private Sub mnubw_Click()

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Emtahan.Refresh
Emtahan.RecordSource = "select * from Emtahan where PARVANDE like ('%" & Label22.Caption & "%') "  'and vazeyat like ('%" & "ﬁ»Ê·Ì" & "%')"
Emtahan.Refresh


'ÃÂ  «Ì‰òÂ ›—«„Ê‘ ‰ò‰œ òœ«„ ò«—‰«„Â —« »«Ìœ ç«Å ò‰œ »«Ìœ òœ «‰Õ’«—Ì ¬‰ —« »Â Œ«ÿ— »”Å«—œ
'Dim KodEnhesariPrint As String


'KodEnhesariPrint = "EMPTY"



'KodEnhesariPrint = Emtahan.Recordset.Fields("kode")
'œ— ç«Å ò«—‰«„Â œ— «‰ Â«Ì òœ »Â œ—œ „Ì ŒÊ—œ

On Error GoTo 9898
GoTo 9999
9898:
MsgBox "«‘ò«· œ— ç«Å ò«—‰«„Â" & Chr$(10) & "„„ò‰ «”  «Ì‰ «‘ò«· »Â ÌòÌ «“ ⁄·· “Ì— »«‘œ" & Chr$(10) & " ‰ÿÌ„«  ›«Ì· Œ—ÊÃÌ ’ÕÌÕ ‰„Ì »«‘œ" & Chr$(10) & "‰„—Â «Ì »—«Ì ﬁ—¬‰ ¬„Ê“ À»  ‰‘œÂ »«‘œ" & Chr$(10) & "«ÿ·«⁄«  Õ”«” „À· ‘„«—Â Å—Ê‰œÂ ﬁ—¬‰ ¬„Ê“ œ” ò«—Ì ‘œÂ »«‘œ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub

9999:
'
Dim ASD As String
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim AD, ParvandeQuranAmooZ As String



If Entekhab.Pc.Checked = True Then
Set oExcel = GetObject(Entekhab.AdressLabel.Caption & "CHARTJ30bw.xlsx")
End If

If Entekhab.net.Checked = True Then
Set oExcel = GetObject(Entekhab.NetAdresslabel.Caption & "CHARTJ30bw.xlsx")
End If

'Set oExcel = GetObject("E:\CHARTJ30.xlsx")
'Set oExcel = GetObject("E:\CHARTJ30.xlsx")
oExcel.ActiveSheet.Range("I26").Value = Taqvim.Tarikh.Caption

'ç” ÊÃÊ »—«Ì À»  ‰«„ Ê „‘Œ’«  ›—œÌ ﬁ—¬ ‰ ¬„Ê“
ParvandeQuranAmooZ = Emtahan.Recordset.Fields("parvande")





'‘„—«Â Å—Ê‰œÂ ﬁ—¬‰ ¬„Ê“



Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" & ParvandeQuranAmooZ & "%')"
Student.Refresh

''''''''''''''''''''''
'ﬁ»  „‘Œ’«  ›—œÌ œ— ò«—‰„«„Â

                oExcel.ActiveSheet.Range("H25").Value = Student.Recordset.Fields("name") & " " & Student.Recordset.Fields("famil")
                oExcel.ActiveSheet.Range("D25").Value = Student.Recordset.Fields("parvande")


'«“ «·«‰ Ê«—œ »Œ‘ À » ‰„—«  „Ì ‘Êœ

Dim xjOZEasli, xjOZProgram As Integer

Emtahan.Refresh
Emtahan.RecordSource = "select * from Emtahan where PARVANDE like ('%" & ParvandeQuranAmooZ & "%') and vazeyat like ('%" & "ﬁ»Ê·Ì" & "%')"
Emtahan.Refresh

For J = 1 To Emtahan.Recordset.RecordCount
xjoze = Emtahan.Recordset.Fields("JOZE")



'À»  ‰„—Â Å«Ì«‰ Ã“ ¡
If Emtahan.Recordset.Fields("NIMPAYAN") = "Å«Ì«‰ Ã“¡" Then

 oExcel.ActiveSheet.Range("O" & xjoze + 3).Value = Val(Emtahan.Recordset.Fields("eNAHAEE"))
 oExcel.ActiveSheet.Range("G" & xjoze + 27).Value = Val(Emtahan.Recordset.Fields("eNAHAEE"))
 ' —Ã„Â Ê „›«ÂÌ„ ‰Ì„Â Ã“¡
 
  oExcel.ActiveSheet.Range("v" & xjoze + 3).Value = Val(Emtahan.Recordset.Fields("mafahim"))
 
'Å«Ì«‰ À» 

End If



'À»  ‰„—Â Å«Ì«‰ Ã“ ¡
If Emtahan.Recordset.Fields("NIMPAYAN") = "‰Ì„Â Ã“¡" Then

 oExcel.ActiveSheet.Range("p" & xjoze + 3).Value = Val(Emtahan.Recordset.Fields("eNAHAEE"))
 oExcel.ActiveSheet.Range("f" & xjoze + 27).Value = Val(Emtahan.Recordset.Fields("eNAHAEE"))
  ' —Ã„Â Ê „›«ÂÌ„ ‰Ì„Â Ã“¡
 
  oExcel.ActiveSheet.Range("u" & xjoze + 3).Value = Val(Emtahan.Recordset.Fields("mafahim"))
 
'Å«Ì«‰ À» 

End If





'À»  „” „—
 oExcel.ActiveSheet.Range("q" & xjoze + 3).Value = Val(Emtahan.Recordset.Fields("Mostamar"))
 oExcel.ActiveSheet.Range("e" & xjoze + 27).Value = Val(Emtahan.Recordset.Fields("Mostamar"))



'À»  €Ì»  „ÊÃÂ
 oExcel.ActiveSheet.Range("r" & xjoze + 3).Value = Val(Emtahan.Recordset.Fields("movajah"))
 oExcel.ActiveSheet.Range("h" & xjoze + 27).Value = Val(Emtahan.Recordset.Fields("movajah"))



'À»  €Ì»  €Ì— „ÊÃÂ

 oExcel.ActiveSheet.Range("s" & xjoze + 3).Value = Val(Emtahan.Recordset.Fields("qeyremovajah"))
 oExcel.ActiveSheet.Range("i" & xjoze + 27).Value = Val(Emtahan.Recordset.Fields("qeyremovajah"))










Emtahan.Recordset.MoveNext

'Å«Ì«‰ Ê«—œ òœ—‰
Next J

'Å«Ì«‰ »Œ‘ À„»   „—«







Emtahan.Refresh
Emtahan.RecordSource = "select * from Emtahan where PARVANDE like ('%" & ParvandeQuranAmooZ & "%') and vazeyat like ('%" & " ÃœÌœ" & "%')"
Emtahan.Refresh









For I = 1 To Emtahan.Recordset.RecordCount
xjoze = Emtahan.Recordset.Fields("JOZE")



'À»  ‰„—Â Å«Ì«‰ Ã“ ¡
If Emtahan.Recordset.Fields("NIMPAYAN") = "Å«Ì«‰ Ã“¡" Then

 'oExcel.ActiveSheet.Range("O" & xjOZE + 3).Value = Val(Emtahan.Recordset.Fields("eNAHAEE"))
 oExcel.ActiveSheet.Range("G" & xjoze + 27).Value = oExcel.ActiveSheet.Range("G" & xjoze + 27).Value & "   "
 ' —Ã„Â Ê „›«ÂÌ„ ‰Ì„Â Ã“¡
 
  'oExcel.ActiveSheet.Range("v" & xjOZE + 3).Value = Val(Emtahan.Recordset.Fields("mafahim"))
 
'Å«Ì«‰ À» 

End If



'À»  ‰„—Â Å«Ì«‰ Ã“ ¡
If Emtahan.Recordset.Fields("NIMPAYAN") = "‰Ì„Â Ã“¡" Then

 'oExcel.ActiveSheet.Range("p" & xjOZE + 3).Value = Val(Emtahan.Recordset.Fields("eNAHAEE"))
 oExcel.ActiveSheet.Range("f" & xjoze + 27).Value = oExcel.ActiveSheet.Range("f" & xjoze + 27).Value & "   "
  ' —Ã„Â Ê „›«ÂÌ„ ‰Ì„Â Ã“¡
 
  'oExcel.ActiveSheet.Range("u" & xjOZE + 3).Value = Val(Emtahan.Recordset.Fields("mafahim"))
 
'Å«Ì«‰ À» 

End If


Emtahan.Recordset.MoveNext

'Å«Ì«‰ Ê«—œ òœ—‰
Next I

'MsgBox "ò«—‰«„Â ¬„«œÂ ç«Å „Ì »«‘œ", vbInformation + vbOKOnly, "ç«Å ò«—‰«„Â"
'

'
oExcel.SaveAs ParvandeQuranAmooZ & "Chart30J"



oExcel.Application.Visible = True
oExcel.Parent.Windows(1).Visible = True
'
'




'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%



End Sub


Private Sub mnuchartjadid_Click()

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
PB1.Value = 1
Emtahan.Refresh
Emtahan.RecordSource = "select * from Emtahan where PARVANDE like ('%" & Label22.Caption & "%') "  'and vazeyat like ('%" & "ﬁ»Ê·Ì" & "%')"
Emtahan.Refresh


'ÃÂ  «Ì‰òÂ ›—«„Ê‘ ‰ò‰œ òœ«„ ò«—‰«„Â —« »«Ìœ ç«Å ò‰œ »«Ìœ òœ «‰Õ’«—Ì ¬‰ —« »Â Œ«ÿ— »”Å«—œ
'Dim KodEnhesariPrint As String


'KodEnhesariPrint = "EMPTY"



'KodEnhesariPrint = Emtahan.Recordset.Fields("kode")
'œ— ç«Å ò«—‰«„Â œ— «‰ Â«Ì òœ »Â œ—œ „Ì ŒÊ—œ

'On Error GoTo 9898
GoTo 9999
9898:
MsgBox "«‘ò«· œ— ç«Å ò«—‰«„Â" & Chr$(10) & "„„ò‰ «”  «Ì‰ «‘ò«· »Â ÌòÌ «“ ⁄·· “Ì— »«‘œ" & Chr$(10) & " ‰ÿÌ„«  ›«Ì· Œ—ÊÃÌ ’ÕÌÕ ‰„Ì »«‘œ" & Chr$(10) & "‰„—Â «Ì »—«Ì ﬁ—¬‰ ¬„Ê“ À»  ‰‘œÂ »«‘œ" & Chr$(10) & "«ÿ·«⁄«  Õ”«” „À· ‘„«—Â Å—Ê‰œÂ ﬁ—¬‰ ¬„Ê“ œ” ò«—Ì ‘œÂ »«‘œ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub

9999:
'
Dim ASD As String
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim AD, ParvandeQuranAmooZ As String



If Entekhab.Pc.Checked = True Then
Set oExcel = GetObject(Entekhab.AdressLabel.Caption & "CHARTJd.xlsx")
PB1.Value = 2

End If

If Entekhab.net.Checked = True Then
Set oExcel = GetObject(Entekhab.NetAdresslabel.Caption & "CHARTJd.xlsx")
PB1.Value = 3
End If

'Set oExcel = GetObject("E:\CHARTJ30.xlsx")
'Set oExcel = GetObject("E:\CHARTJ30.xlsx")
oExcel.ActiveSheet.Range("ak1").Value = Taqvim.Tarikh.Caption
'oExcel.ActiveSheet.Range("aj1").Value = lkodclass.Caption

'ç” ÊÃÊ »—«Ì À»  ‰«„ Ê „‘Œ’«  ›—œÌ ﬁ—¬ ‰ ¬„Ê“
ParvandeQuranAmooZ = Emtahan.Recordset.Fields("parvande")





'‘„—«Â Å—Ê‰œÂ ﬁ—¬‰ ¬„Ê“


PB1.Value = 4
Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" & ParvandeQuranAmooZ & "%')"
Student.Refresh

''''''''''''''''''''''
'ﬁ»  „‘Œ’«  ›—œÌ œ— ò«—‰„«„Â

                oExcel.ActiveSheet.Range("u1").Value = Student.Recordset.Fields("name") & " " & Student.Recordset.Fields("famil")
                oExcel.ActiveSheet.Range("b1").Value = Student.Recordset.Fields("parvande")


'«“ «·«‰ Ê«—œ »Œ‘ À » ‰„—«  „Ì ‘Êœ

Dim xjOZEasli, xjOZProgram As Integer

Emtahan.Refresh
Emtahan.RecordSource = "select * from Emtahan where PARVANDE like ('%" & ParvandeQuranAmooZ & "%') and vazeyat like ('%" & "ﬁ»Ê·Ì" & "%')"
Emtahan.Refresh

For J = 1 To Emtahan.Recordset.RecordCount
xjoze = Emtahan.Recordset.Fields("JOZE")


If Val(Emtahan.Recordset.Fields("JOZE")) = 30 Then

 oExcel.ActiveSheet.Range("an2").Value = Val(Emtahan.Recordset.Fields("eNAHAEE"))
' oExcel.ActiveSheet.Range("G" & xjoze + 27).Value = Val(Emtahan.Recordset.Fields("eNAHAEE"))
 ' —Ã„Â Ê „›«ÂÌ„ ‰Ì„Â Ã“¡
 
  oExcel.ActiveSheet.Range("ar2").Value = Val(Emtahan.Recordset.Fields("mafahim"))
  PB1.Value = 5
  
 oExcel.ActiveSheet.Range("ao2").Value = Val(Emtahan.Recordset.Fields("Mostamar"))
 oExcel.ActiveSheet.Range("ap2").Value = Val(Emtahan.Recordset.Fields("movajah"))
 oExcel.ActiveSheet.Range("aq2").Value = Val(Emtahan.Recordset.Fields("qeyremovajah"))
' oExcel.ActiveSheet.Range("ar" & xjoze + (xjoze - 1) + 3).Value = Val(Emtahan.Recordset.Fields("mafahim"))
 
GoTo 222
End If

'À»  ‰„—Â Å«Ì«‰ Ã“ ¡
If Emtahan.Recordset.Fields("NIMPAYAN") = "Å«Ì«‰ Ã“¡" Then

 oExcel.ActiveSheet.Range("an" & xjoze + (xjoze - 1) + 3).Value = Val(Emtahan.Recordset.Fields("eNAHAEE"))
' oExcel.ActiveSheet.Range("G" & xjoze + 27).Value = Val(Emtahan.Recordset.Fields("eNAHAEE"))
 ' —Ã„Â Ê „›«ÂÌ„ ‰Ì„Â Ã“¡
 
  oExcel.ActiveSheet.Range("ar" & xjoze + (xjoze - 1) + 3).Value = Val(Emtahan.Recordset.Fields("mafahim"))
  
  
 oExcel.ActiveSheet.Range("ao" & xjoze + (xjoze - 1) + 3).Value = Val(Emtahan.Recordset.Fields("Mostamar"))
 oExcel.ActiveSheet.Range("ap" & xjoze + (xjoze - 1) + 3).Value = Val(Emtahan.Recordset.Fields("movajah"))
oExcel.ActiveSheet.Range("aq" & xjoze + (xjoze - 1) + 3).Value = Val(Emtahan.Recordset.Fields("qeyremovajah"))
 oExcel.ActiveSheet.Range("as" & xjoze + (xjoze - 1) + 3).Value = Val(Emtahan.Recordset.Fields("qeyremovajah")) + Val(Emtahan.Recordset.Fields("movajah"))
 
'Å«Ì«‰ À» 

End If



'À»  ‰„—Â Å«Ì«‰ Ã“ ¡
If Emtahan.Recordset.Fields("NIMPAYAN") = "‰Ì„Â Ã“¡" Then
Dim NimeJozeKasrshode As Single
NimeJozeKasrshode = (Val(xjoze) - Val(0.5) + 1) / Val(0.5)
 
 oExcel.ActiveSheet.Range("an" & NimeJozeKasrshode).Value = Val(Emtahan.Recordset.Fields("eNAHAEE"))
 'oExcel.ActiveSheet.Range("f" & xjoze + 27).Value = Val(Emtahan.Recordset.Fields("eNAHAEE"))
  ' —Ã„Â Ê „›«ÂÌ„ ‰Ì„Â Ã“¡
 
  oExcel.ActiveSheet.Range("ar" & NimeJozeKasrshode).Value = Val(Emtahan.Recordset.Fields("mafahim"))
  PB1.Value = 6
  
   oExcel.ActiveSheet.Range("ao" & NimeJozeKasrshode).Value = Val(Emtahan.Recordset.Fields("Mostamar"))
 oExcel.ActiveSheet.Range("ap" & NimeJozeKasrshode).Value = Val(Emtahan.Recordset.Fields("movajah"))
 oExcel.ActiveSheet.Range("aq" & NimeJozeKasrshode).Value = Val(Emtahan.Recordset.Fields("qeyremovajah"))

   oExcel.ActiveSheet.Range("as" & NimeJozeKasrshode).Value = Val(Emtahan.Recordset.Fields("qeyremovajah")) + Val(Emtahan.Recordset.Fields("movajah"))
 
 
'Å«Ì«‰ À» 

End If

' Ì   —«ÌŒ «„ Õ«‰ 


 'oExcel.ActiveSheet.Range("T" & xjoze + 3).Value = Val(Emtahan.Recordset.Fields("TEMTAHAN"))
 'oExcel.ActiveSheet.Range("D" & xjoze + 27).Value = Val(Emtahan.Recordset.Fields("TEMTAHAN"))





'À»  „” „—
' oExcel.ActiveSheet.Range("ao" & xjoze + (xjoze - 1) + 3).Value = Val(Emtahan.Recordset.Fields("Mostamar"))
 'oExcel.ActiveSheet.Range("e" & xjoze + 27).Value = Val(Emtahan.Recordset.Fields("Mostamar"))



'À»  €Ì»  „ÊÃÂ
 'oExcel.ActiveSheet.Range("r" & xjoze + 3).Value = Val(Emtahan.Recordset.Fields("movajah"))
 'oExcel.ActiveSheet.Range("h" & xjoze + 27).Value = Val(Emtahan.Recordset.Fields("movajah"))



'À»  €Ì»  €Ì— „ÊÃÂ

 'oExcel.ActiveSheet.Range("s" & xjoze + 3).Value = Val(Emtahan.Recordset.Fields("qeyremovajah"))
 'oExcel.ActiveSheet.Range("i" & xjoze + 27).Value = Val(Emtahan.Recordset.Fields("qeyremovajah"))








222:

Emtahan.Recordset.MoveNext

'Å«Ì«‰ Ê«—œ òœ—‰
Next J

'Å«Ì«‰ »Œ‘ À„»   „—«
PB1.Value = 7

GoTo 500200





Emtahan.Refresh
Emtahan.RecordSource = "select * from Emtahan where PARVANDE like ('%" & ParvandeQuranAmooZ & "%') and vazeyat like ('%" & " ÃœÌœ" & "%')"
Emtahan.Refresh






PB1.Value = 8


For I = 1 To Emtahan.Recordset.RecordCount
xjoze = Emtahan.Recordset.Fields("JOZE")



'À»  ‰„—Â Å«Ì«‰ Ã“ ¡
If Emtahan.Recordset.Fields("NIMPAYAN") = "Å«Ì«‰ Ã“¡" Then

 'oExcel.ActiveSheet.Range("O" & xjOZE + 3).Value = Val(Emtahan.Recordset.Fields("eNAHAEE"))
 oExcel.ActiveSheet.Range("G" & xjoze + 27).Value = oExcel.ActiveSheet.Range("G" & xjoze + 27).Value & "   "
 ' —Ã„Â Ê „›«ÂÌ„ ‰Ì„Â Ã“¡
 
  'oExcel.ActiveSheet.Range("v" & xjOZE + 3).Value = Val(Emtahan.Recordset.Fields("mafahim"))
 
'Å«Ì«‰ À» 

End If


PB1.Value = 9

'À»  ‰„—Â Å«Ì«‰ Ã“ ¡
If Emtahan.Recordset.Fields("NIMPAYAN") = "‰Ì„Â Ã“¡" Then

 'oExcel.ActiveSheet.Range("p" & xjOZE + 3).Value = Val(Emtahan.Recordset.Fields("eNAHAEE"))
 oExcel.ActiveSheet.Range("f" & xjoze + 27).Value = oExcel.ActiveSheet.Range("f" & xjoze + 27).Value & "   "
  ' —Ã„Â Ê „›«ÂÌ„ ‰Ì„Â Ã“¡
 
  'oExcel.ActiveSheet.Range("u" & xjOZE + 3).Value = Val(Emtahan.Recordset.Fields("mafahim"))
 
'Å«Ì«‰ À» 

End If


Emtahan.Recordset.MoveNext

'Å«Ì«‰ Ê«—œ òœ—‰
Next I

'MsgBox "ò«—‰«„Â ¬„«œÂ ç«Å „Ì »«‘œ", vbInformation + vbOKOnly, "ç«Å ò«—‰«„Â"
'
500200:
'
oExcel.SaveAs ParvandeQuranAmooZ & "Chart30J"

PB1.Value = 10

oExcel.Application.Visible = True
oExcel.Parent.Windows(1).Visible = True
'
'
PB1.Value = 0



'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


End Sub

Private Sub mnucolor_Click()

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Emtahan.Refresh
Emtahan.RecordSource = "select * from Emtahan where PARVANDE like ('%" & Label22.Caption & "%') "  'and vazeyat like ('%" & "ﬁ»Ê·Ì" & "%')"
Emtahan.Refresh


'ÃÂ  «Ì‰òÂ ›—«„Ê‘ ‰ò‰œ òœ«„ ò«—‰«„Â —« »«Ìœ ç«Å ò‰œ »«Ìœ òœ «‰Õ’«—Ì ¬‰ —« »Â Œ«ÿ— »”Å«—œ
'Dim KodEnhesariPrint As String


'KodEnhesariPrint = "EMPTY"



'KodEnhesariPrint = Emtahan.Recordset.Fields("kode")
'œ— ç«Å ò«—‰«„Â œ— «‰ Â«Ì òœ »Â œ—œ „Ì ŒÊ—œ

On Error GoTo 9898
GoTo 9999
9898:
MsgBox "«‘ò«· œ— ç«Å ò«—‰«„Â" & Chr$(10) & "„„ò‰ «”  «Ì‰ «‘ò«· »Â ÌòÌ «“ ⁄·· “Ì— »«‘œ" & Chr$(10) & " ‰ÿÌ„«  ›«Ì· Œ—ÊÃÌ ’ÕÌÕ ‰„Ì »«‘œ" & Chr$(10) & "‰„—Â «Ì »—«Ì ﬁ—¬‰ ¬„Ê“ À»  ‰‘œÂ »«‘œ" & Chr$(10) & "«ÿ·«⁄«  Õ”«” „À· ‘„«—Â Å—Ê‰œÂ ﬁ—¬‰ ¬„Ê“ œ” ò«—Ì ‘œÂ »«‘œ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub

9999:
'
Dim ASD As String
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim AD, ParvandeQuranAmooZ As String



If Entekhab.Pc.Checked = True Then
Set oExcel = GetObject(Entekhab.AdressLabel.Caption & "CHARTJ30.xlsx")
End If

If Entekhab.net.Checked = True Then
Set oExcel = GetObject(Entekhab.NetAdresslabel.Caption & "CHARTJ30.xlsx")
End If

'Set oExcel = GetObject("E:\CHARTJ30.xlsx")
'Set oExcel = GetObject("E:\CHARTJ30.xlsx")
oExcel.ActiveSheet.Range("I26").Value = Taqvim.Tarikh.Caption

'ç” ÊÃÊ »—«Ì À»  ‰«„ Ê „‘Œ’«  ›—œÌ ﬁ—¬ ‰ ¬„Ê“
ParvandeQuranAmooZ = Emtahan.Recordset.Fields("parvande")





'‘„—«Â Å—Ê‰œÂ ﬁ—¬‰ ¬„Ê“



Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" & ParvandeQuranAmooZ & "%')"
Student.Refresh

''''''''''''''''''''''
'ﬁ»  „‘Œ’«  ›—œÌ œ— ò«—‰„«„Â

                oExcel.ActiveSheet.Range("H25").Value = Student.Recordset.Fields("name") & " " & Student.Recordset.Fields("famil")
                oExcel.ActiveSheet.Range("D25").Value = Student.Recordset.Fields("parvande")


'«“ «·«‰ Ê«—œ »Œ‘ À » ‰„—«  „Ì ‘Êœ

Dim xjOZEasli, xjOZProgram As Integer

Emtahan.Refresh
Emtahan.RecordSource = "select * from Emtahan where PARVANDE like ('%" & ParvandeQuranAmooZ & "%') and vazeyat like ('%" & "ﬁ»Ê·Ì" & "%')"
Emtahan.Refresh

For J = 1 To Emtahan.Recordset.RecordCount
xjoze = Emtahan.Recordset.Fields("JOZE")



'À»  ‰„—Â Å«Ì«‰ Ã“ ¡
If Emtahan.Recordset.Fields("NIMPAYAN") = "Å«Ì«‰ Ã“¡" Then

 oExcel.ActiveSheet.Range("O" & xjoze + 3).Value = Val(Emtahan.Recordset.Fields("eNAHAEE"))
 oExcel.ActiveSheet.Range("G" & xjoze + 27).Value = Val(Emtahan.Recordset.Fields("eNAHAEE"))
 ' —Ã„Â Ê „›«ÂÌ„ ‰Ì„Â Ã“¡
 
  oExcel.ActiveSheet.Range("v" & xjoze + 3).Value = Val(Emtahan.Recordset.Fields("mafahim"))
 
'Å«Ì«‰ À» 

End If



'À»  ‰„—Â Å«Ì«‰ Ã“ ¡
If Emtahan.Recordset.Fields("NIMPAYAN") = "‰Ì„Â Ã“¡" Then

 oExcel.ActiveSheet.Range("p" & xjoze + 3).Value = Val(Emtahan.Recordset.Fields("eNAHAEE"))
 oExcel.ActiveSheet.Range("f" & xjoze + 27).Value = Val(Emtahan.Recordset.Fields("eNAHAEE"))
  ' —Ã„Â Ê „›«ÂÌ„ ‰Ì„Â Ã“¡
 
  oExcel.ActiveSheet.Range("u" & xjoze + 3).Value = Val(Emtahan.Recordset.Fields("mafahim"))
 
'Å«Ì«‰ À» 

End If

' Ì   —«ÌŒ «„ Õ«‰ 


 oExcel.ActiveSheet.Range("T" & xjoze + 3).Value = Val(Emtahan.Recordset.Fields("TEMTAHAN"))
 oExcel.ActiveSheet.Range("D" & xjoze + 27).Value = Val(Emtahan.Recordset.Fields("TEMTAHAN"))





'À»  „” „—
 oExcel.ActiveSheet.Range("q" & xjoze + 3).Value = Val(Emtahan.Recordset.Fields("Mostamar"))
 oExcel.ActiveSheet.Range("e" & xjoze + 27).Value = Val(Emtahan.Recordset.Fields("Mostamar"))



'À»  €Ì»  „ÊÃÂ
 oExcel.ActiveSheet.Range("r" & xjoze + 3).Value = Val(Emtahan.Recordset.Fields("movajah"))
 oExcel.ActiveSheet.Range("h" & xjoze + 27).Value = Val(Emtahan.Recordset.Fields("movajah"))



'À»  €Ì»  €Ì— „ÊÃÂ

 oExcel.ActiveSheet.Range("s" & xjoze + 3).Value = Val(Emtahan.Recordset.Fields("qeyremovajah"))
 oExcel.ActiveSheet.Range("i" & xjoze + 27).Value = Val(Emtahan.Recordset.Fields("qeyremovajah"))










Emtahan.Recordset.MoveNext

'Å«Ì«‰ Ê«—œ òœ—‰
Next J

'Å«Ì«‰ »Œ‘ À„»   „—«







Emtahan.Refresh
Emtahan.RecordSource = "select * from Emtahan where PARVANDE like ('%" & ParvandeQuranAmooZ & "%') and vazeyat like ('%" & " ÃœÌœ" & "%')"
Emtahan.Refresh









For I = 1 To Emtahan.Recordset.RecordCount
xjoze = Emtahan.Recordset.Fields("JOZE")



'À»  ‰„—Â Å«Ì«‰ Ã“ ¡
If Emtahan.Recordset.Fields("NIMPAYAN") = "Å«Ì«‰ Ã“¡" Then

 'oExcel.ActiveSheet.Range("O" & xjOZE + 3).Value = Val(Emtahan.Recordset.Fields("eNAHAEE"))
 oExcel.ActiveSheet.Range("G" & xjoze + 27).Value = oExcel.ActiveSheet.Range("G" & xjoze + 27).Value & "   "
 ' —Ã„Â Ê „›«ÂÌ„ ‰Ì„Â Ã“¡
 
  'oExcel.ActiveSheet.Range("v" & xjOZE + 3).Value = Val(Emtahan.Recordset.Fields("mafahim"))
 
'Å«Ì«‰ À» 

End If



'À»  ‰„—Â Å«Ì«‰ Ã“ ¡
If Emtahan.Recordset.Fields("NIMPAYAN") = "‰Ì„Â Ã“¡" Then

 'oExcel.ActiveSheet.Range("p" & xjOZE + 3).Value = Val(Emtahan.Recordset.Fields("eNAHAEE"))
 oExcel.ActiveSheet.Range("f" & xjoze + 27).Value = oExcel.ActiveSheet.Range("f" & xjoze + 27).Value & "   "
  ' —Ã„Â Ê „›«ÂÌ„ ‰Ì„Â Ã“¡
 
  'oExcel.ActiveSheet.Range("u" & xjOZE + 3).Value = Val(Emtahan.Recordset.Fields("mafahim"))
 
'Å«Ì«‰ À» 

End If


Emtahan.Recordset.MoveNext

'Å«Ì«‰ Ê«—œ òœ—‰
Next I

'MsgBox "ò«—‰«„Â ¬„«œÂ ç«Å „Ì »«‘œ", vbInformation + vbOKOnly, "ç«Å ò«—‰«„Â"
'

'
oExcel.SaveAs ParvandeQuranAmooZ & "Chart30J"



oExcel.Application.Visible = True
oExcel.Parent.Windows(1).Visible = True
'
'




'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%



End Sub


Private Sub net_Click()
Pc.Checked = False
net.Checked = True
End Sub

Private Sub pc_Click()
Pc.Checked = True
net.Checked = False
End Sub

Private Sub mnuhome_Click()
Entekhab.Show

End Sub

Private Sub mnusettingemtahan_Click()
SettingEmtahan.Show

End Sub

Private Sub RokhaniKarname_Click()



'ÃÂ  «Ì‰òÂ ›—«„Ê‘ ‰ò‰œ òœ«„ ò«—‰«„Â —« »«Ìœ ç«Å ò‰œ »«Ìœ òœ «‰Õ’«—Ì ¬‰ —« »Â Œ«ÿ— »”Å«—œ
Dim KodEnhesariPrint As String


KodEnhesariPrint = "EMPTY"



KodEnhesariPrint = Emtahan.Recordset.Fields("kode")
'œ— ç«Å ò«—‰«„Â œ— «‰ Â«Ì òœ »Â œ—œ „Ì ŒÊ—œ

'On Error GoTo 9898
GoTo 9999
9898:
MsgBox "«‘ò«· œ— ç«Å ò«—‰«„Â" & Chr$(10) & "„„ò‰ «”  «Ì‰ «‘ò«· »Â ÌòÌ «“ ⁄·· “Ì— »«‘œ" & Chr$(10) & " ‰ÿÌ„«  ›«Ì· Œ—ÊÃÌ ’ÕÌÕ ‰„Ì »«‘œ" & Chr$(10) & "‰„—Â «Ì »—«Ì ﬁ—¬‰ ¬„Ê“ À»  ‰‘œÂ »«‘œ" & Chr$(10) & "«ÿ·«⁄«  Õ”«” „À· ‘„«—Â Å—Ê‰œÂ ﬁ—¬‰ ¬„Ê“ œ” ò«—Ì ‘œÂ »«‘œ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub

9999:

If Me.Emtahan.Recordset.Fields("Chap") = "ç«Å ‘œÂ" Then
    If MsgBox(" «Ì‰ ò«—‰«„Â ﬁ»·« œ—  «—ÌŒ " & Me.Emtahan.Recordset.Fields("dateofchap") & " ç«Å ‘œÂ «”  ¬Ì« „Ì ŒÊ«ÂÌœ œÊ»«—Â ¬‰ —« ç«Å ò‰Ìœ  ", vbExclamation + vbYesNo, "ç«Å ò«—‰«„Â") = vbYes Then
    GoTo 500
    Else
    Exit Sub
    End If
    
End If
500
Dim ASD As String
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim AD, ParvandeQuranAmooZ As String


If Entekhab.Pc.Checked = True Then
'Set oExcel = GetObject("F:\Markaz Quran & Hadis\FORMXLS\CopyofKarnameJadid.xlsx")
Set oExcel = GetObject(Entekhab.AdressLabel.Caption & "KarnameDOmomi.xlsx")

End If

If Entekhab.net.Checked = True Then
Set oExcel = GetObject(Entekhab.NetAdresslabel.Caption & "KarnameDOmomi.xlsx")
End If


'Set oExcel = GetObject("d:\vadiexls.xlsx")
oExcel.ActiveSheet.Range("f3").Value = Taqvim.Tarikh.Caption

'ç” ÊÃÊ »—«Ì À»  ‰«„ Ê „‘Œ’«  ›—œÌ ﬁ—¬ ‰ ¬„Ê“
ParvandeQuranAmooZ = Emtahan.Recordset.Fields("parvande")
Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" & ParvandeQuranAmooZ & "%')"
Student.Refresh

''''''''''''''''''''''
'ﬁ»  „‘Œ’«  ›—œÌ œ— ò«—‰„«„Â

                oExcel.ActiveSheet.Range("c2").Value = Student.Recordset.Fields("name") & " " & Student.Recordset.Fields("famil")
                oExcel.ActiveSheet.Range("c3").Value = Student.Recordset.Fields("parvande")
'Ã” ÊÃÊ »—«Ì À»  „‘Œ’«  ò·«”·

mclass.Refresh
mclass.RecordSource = "select * from mclass where kodclass like ('%" + Emtahan.Recordset.Fields("kodclass") + "%')"
mclass.Refresh


'À»  „‘Œ’«   ò·«”
                oExcel.ActiveSheet.Range("f2").Value = mclass.Recordset.Fields("ostad")
                
                
                'À»  ⁄‰Ê«‰ ò«—‰«„Â œÊ—Â ⁄„Ê„Ì
                oExcel.ActiveSheet.Range("a1").Value = " ò«—‰«„Â œÊ—Â " & Emtahan.Recordset.Fields("tarh")


'mahdodeemtahan
                oExcel.ActiveSheet.Range("a5").Value = Emtahan.Recordset.Fields("mahdodeemtahan")

'À»  ‰Ì„Â Ê Å«Ì«‰ Ã“∆ »Êœ‰ œ— ò«—‰«„Â

'                oExcel.ActiveSheet.Range("a5").Value = Emtahan.Recordset.Fields("nimpayan") & " " & Emtahan.Recordset.Fields("joze")

'À»  ‰„—«  œ— ò«—‰«„Â

'                oExcel.ActiveSheet.Range("b5").Value = Emtahan.Recordset.Fields("hefz")
'                oExcel.ActiveSheet.Range("c5").Value = Emtahan.Recordset.Fields("ee")
               'À»  „” „—
              
             oExcel.ActiveSheet.Range("b5").Value = Emtahan.Recordset.Fields("mostamar")
                
                
                
                'oExcel.ActiveSheet.Range("d5").Value = Emtahan.Recordset.Fields("mafahim")
'oExcel.ActiveSheet.Range("f4").Value = Emtahan.Recordset.Fields("kasremtiaz")

'‰„—Â ‰Â«ÌÌ
                oExcel.ActiveSheet.Range("g5").Value = Emtahan.Recordset.Fields("enahaee")


'À»  Ê÷⁄Ì  ﬁ»Ê·Ì Ì«  ÃÌœÌœ

'À»   ⁄œ«œ €Ì»  Â«
                oExcel.ActiveSheet.Range("f5").Value = Emtahan.Recordset.Fields("tqeybat")

'À»  — »Â ò·«”Ì
                oExcel.ActiveSheet.Range("e6").Value = Emtahan.Recordset.Fields("rotbe")
                
               'tarikh Azmoon
            oExcel.ActiveSheet.Range("b6").Value = Emtahan.Recordset.Fields("temtahan")
            
            'sabt Vazeyat
            oExcel.ActiveSheet.Range("g6").Value = Emtahan.Recordset.Fields("vazeyat")
            
            'q Movajah
            oExcel.ActiveSheet.Range("e5").Value = Emtahan.Recordset.Fields("Movajah")
            
            'katbi shafahi
            
            oExcel.ActiveSheet.Range("c5").Value = Emtahan.Recordset.Fields("katbi")
            oExcel.ActiveSheet.Range("D5").Value = Emtahan.Recordset.Fields("shafahi")
                
                
                
                
                

Dim SearchKarname, OPR, NimPayanJozeSabeq, JozeSabeq As String
'ÃÂ  À»  ‰„—«   çœÌœÌ ﬁ»·Ì òÂ ‘«„·  ÃœÌœÌ 1 Ê  ÃœÌœÌ 2 „Ì »«‘œ

'
'Å«Ì«‰ À»  ‰„—Â ‰Ì„ Ã“¡ ”«»ﬁ
'À»   Ê÷ÌÕ« 

Emtahan.Refresh
Emtahan.RecordSource = "select * from Emtahan where kode like ('%" & KodEnhesariPrint & "%')"
Emtahan.Refresh

'À»  ç«Å ò«—‰«„Â
Emtahan.Recordset.Fields("Chap") = "ç«Å ‘œÂ"
Emtahan.Recordset.Fields("dateofChap") = Me.stb1.Panels(3).Text
Emtahan.Recordset.Update
Emtahan.Refresh

'œ— ç«Å ò«—‰«„Â œ— «‰ Â«Ì òœ »Â œ—œ „Ì ŒÊ—œ






MsgBox "ò«—‰«„Â ¬„«œÂ ç«Å „Ì »«‘œ", vbInformation + vbOKOnly, "ç«Å ò«—‰«„Â"
'

''''''
oExcel.SaveAs KodEnhesariPrint

oExcel.Application.Visible = True
oExcel.Parent.Windows(1).Visible = True
'
'
'Set oExcel = Nothing ' Remove object variable.
''''''''

'oExcel.FileRef.Visible = True
'Catch
'»Â —Ê“ —”«‰Ì ¬„«— „ÊÃÊœ œ— ›—„
Call Command2_Click


End Sub

Private Sub TabStrip1_Click()
Select Case TabStrip1.SelectedItem.Caption

            Case "ﬁ—¬‰ ¬„Ê“«‰"
            DataGrid2.Visible = True
            DataGrid1.Visible = False
            Command1.Enabled = False
            
            
            
            Case "‰„—«  À»  ‘œÂ"
            On Error Resume Next
            DataGrid2.Visible = False
            DataGrid1.Visible = True
            Command1.Enabled = True
            Emtahan.Refresh
            Emtahan.RecordSource = "select * from Emtahan where parvande like ('%" & Student.Recordset.Fields("parvande") & "%') "
            Emtahan.Refresh
          
                
End Select
End Sub

Private Sub Text2_Change()
Emtahan.Refresh
Emtahan.RecordSource = "select * from Emtahan where parvande like ('%" & Text2.Text & "%') or kode like ('%" & Text2.Text & "%') or kodclass like ('%" & Text2.Text & "%')"
Emtahan.Refresh

Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" & Text2.Text & "%') or nf like ('%" & Text2.Text & "%')"
Student.Refresh
End Sub
