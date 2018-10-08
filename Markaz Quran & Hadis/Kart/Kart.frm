VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Kart 
   Caption         =   "ò«—  ⁄÷ÊÌ "
   ClientHeight    =   8700
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11700
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Kart.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8700
   ScaleWidth      =   11700
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command12 
      BackColor       =   &H80000002&
      Caption         =   "„Ãœœ"
      Height          =   300
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5640
      Picture         =   "Kart.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "«›“Êœ‰ Ê–ŒÌ—Â ò—œ‰"
      Top             =   240
      Width           =   855
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000003&
      Height          =   2580
      ItemData        =   "Kart.frx":461A
      Left            =   3480
      List            =   "Kart.frx":461C
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   1080
      Width           =   8055
   End
   Begin VB.CommandButton Command10 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      Picture         =   "Kart.frx":461E
      Style           =   1  'Graphical
      TabIndex        =   42
      TabStop         =   0   'False
      ToolTipText     =   "—òÊ—œÌ òÂ „ÌŒÊ«ÂÌœ Õ–› ‰„«ÌÌœ —« «‰ Œ«» Ê «Ì‰ œò„Â —« »“‰Ìœ"
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      Picture         =   "Kart.frx":8BE3
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "ÃÂ  «÷«›Â ò—œ‰ —òÊ—œ «Ì‰ œò„Â —« »“‰Ìœ"
      Top             =   240
      Width           =   855
   End
   Begin VB.Frame Frame5 
      Caption         =   "„‘Œ’«  ﬁ—¬‰ ¬„Ê“"
      Height          =   3975
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   3255
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "ò«— "
         Height          =   315
         Left            =   2280
         TabIndex        =   40
         Top             =   3240
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Xselect"
         DataSource      =   "Student"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   480
         TabIndex        =   39
         Top             =   3240
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "‘„«—Â Å—Ê‰œÂ"
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   36
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "‰«„"
         Height          =   315
         Left            =   2280
         TabIndex        =   35
         Top             =   720
         Width           =   195
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "‰«„ Œ«‰Ê«œêÌ"
         Height          =   315
         Left            =   2280
         TabIndex        =   34
         Top             =   1080
         Width           =   870
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "‰«„ Åœ—"
         Height          =   315
         Left            =   2280
         TabIndex        =   33
         Top             =   1440
         Width           =   450
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ  Ê·œ"
         Height          =   315
         Left            =   2280
         TabIndex        =   32
         Top             =   1800
         Width           =   660
      End
      Begin VB.Label Label9 
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
         TabIndex        =   31
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label12 
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
         Left            =   480
         TabIndex        =   30
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label13 
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
         Left            =   480
         TabIndex        =   29
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Tavalod"
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
         TabIndex        =   28
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label15 
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
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   480
         TabIndex        =   27
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "òœ „·Ì"
         Height          =   315
         Left            =   2280
         TabIndex        =   26
         Top             =   2160
         Width           =   465
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Kodmeli"
         DataSource      =   "Student"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   480
         TabIndex        =   25
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Shsh"
         DataSource      =   "Student"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   480
         TabIndex        =   24
         Top             =   2520
         Width           =   135
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "‘.‘"
         Height          =   315
         Left            =   2280
         TabIndex        =   23
         Top             =   2520
         Width           =   420
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   " Ê÷ÌÕ« "
         Height          =   315
         Left            =   2280
         TabIndex        =   22
         Top             =   2880
         Width           =   585
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "Tozih"
         DataSource      =   "Student"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   480
         TabIndex        =   21
         Top             =   2880
         Width           =   135
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "„‘Œ’«  êÊ«ÂÌ ‰«„Â"
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "ç«Å"
         Height          =   315
         Left            =   2280
         TabIndex        =   19
         Top             =   2040
         Width           =   300
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "CHAP"
         DataSource      =   "Govahi"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   480
         TabIndex        =   18
         Top             =   2040
         Width           =   135
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "SATH"
         DataSource      =   "Govahi"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   480
         TabIndex        =   17
         Top             =   1320
         Width           =   135
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "”ÿÕ"
         Height          =   315
         Left            =   2280
         TabIndex        =   16
         Top             =   1320
         Width           =   330
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "JOZE"
         DataSource      =   "Govahi"
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
         TabIndex        =   15
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "kodG"
         DataSource      =   "Govahi"
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
         TabIndex        =   14
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Ã“¡"
         Height          =   315
         Left            =   2280
         TabIndex        =   13
         Top             =   960
         Width           =   270
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "‘„«—Â"
         Height          =   315
         Index           =   2
         Left            =   2280
         TabIndex        =   12
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "Ê÷⁄Ì  ’œÊ—"
         Height          =   315
         Left            =   2280
         TabIndex        =   11
         Top             =   2760
         Width           =   840
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "sader"
         DataSource      =   "Govahi"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   480
         TabIndex        =   10
         Top             =   2760
         Width           =   135
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "TAHVIL"
         DataSource      =   "Govahi"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   480
         TabIndex        =   9
         Top             =   2400
         Width           =   135
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   " ÕÊÌ·"
         Height          =   315
         Left            =   2280
         TabIndex        =   8
         Top             =   2400
         Width           =   405
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "noe"
         DataSource      =   "Govahi"
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
         TabIndex        =   7
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "‰Ê⁄ êÊ«ÂÌ"
         Height          =   315
         Left            =   2280
         TabIndex        =   6
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "„⁄œ·"
         Height          =   315
         Left            =   2280
         TabIndex        =   5
         Top             =   1680
         Width           =   360
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "-"
         DataField       =   "MOADEL"
         DataSource      =   "Govahi"
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   480
         TabIndex        =   4
         Top             =   1680
         Width           =   135
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Motor"
      Height          =   3375
      Left            =   360
      TabIndex        =   1
      Top             =   8280
      Visible         =   0   'False
      Width           =   3255
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
         Connect         =   $"Kart.frx":C640
         OLEDBString     =   $"Kart.frx":C6C9
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from govahi"
         Caption         =   "Govahi"
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
         Connect         =   $"Kart.frx":C752
         OLEDBString     =   $"Kart.frx":C7DB
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from qeybat"
         Caption         =   "Qeybat"
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
         Connect         =   $"Kart.frx":C864
         OLEDBString     =   $"Kart.frx":C8ED
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from emtahan"
         Caption         =   "Emtahan"
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
         Connect         =   $"Kart.frx":C976
         OLEDBString     =   $"Kart.frx":C9FF
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
         Connect         =   $"Kart.frx":CA88
         OLEDBString     =   $"Kart.frx":CB11
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
         Connect         =   $"Kart.frx":CB9A
         OLEDBString     =   $"Kart.frx":CC23
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
         Connect         =   $"Kart.frx":CCAC
         OLEDBString     =   $"Kart.frx":CD35
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
         Connect         =   $"Kart.frx":CDBE
         OLEDBString     =   $"Kart.frx":CE47
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
      Begin MSAdodcLib.Adodc vadie 
         Height          =   330
         Left            =   0
         Top             =   0
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
         Connect         =   $"Kart.frx":CED0
         OLEDBString     =   $"Kart.frx":CF59
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
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   6600
      TabIndex        =   0
      Top             =   480
      Width           =   4935
   End
   Begin ComctlLib.StatusBar stb1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   8325
      Width           =   11700
      _ExtentX        =   20638
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
   Begin MSDataGridLib.DataGrid DataGridSTUDENT 
      Bindings        =   "Kart.frx":CFE2
      Height          =   4215
      Left            =   120
      TabIndex        =   37
      Top             =   4080
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   7435
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777152
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ã” ÃÊ"
      Height          =   330
      Left            =   10920
      TabIndex        =   38
      Top             =   120
      Width           =   420
   End
End
Attribute VB_Name = "Kart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command10_Click()
On Error Resume Next


Student.Recordset.Fields("xkard") = "0"
Student.Recordset.Update
Beep
Call Command6_Click

End Sub

Private Sub Command12_Click()
Call Text1_Change


End Sub

Private Sub Command6_Click()
On Error Resume Next

Student.Refresh
Student.RecordSource = "select * from student where xkard like ('%" & "Kard-New" & "%') "
Student.Refresh

List1.Clear


For i = 1 To Student.Recordset.RecordCount

List1.AddItem (Student.Recordset.Fields("parvande") & " - " & Student.Recordset.Fields("name") & " - " & Student.Recordset.Fields("famil") & " - " & Student.Recordset.Fields("d"))

Student.Recordset.MoveNext
Next i

End Sub

Private Sub Command7_Click()
'On Error Resume Next
vadie.Refresh
vadie.RecordSource = "select * from vadie where parvande like ('%" & Label9.Caption & "%')"
vadie.Refresh
If vadie.Recordset.BOF = True Or vadie.Recordset.EOF = True Or vadie.Recordset.RecordCount = 0 Then
If MsgBox("»—«Ì «Ì‰ ﬁ—¬‰ ¬„Ê“ ÊœÌ⁄Â À»  ‰‘œÂ «”  ¬Ì« «œ«„Â „Ì œÂÌœ", vbCritical + vbYesNo, "ÊœÌ⁄Â") = vbNo Then Exit Sub
End If

If Student.Recordset.Fields("Clas1") = "‰œ«—œ" And Student.Recordset.Fields("Clas2") = "‰œ«—œ" And Student.Recordset.Fields("Clas3") = "‰œ«—œ" And Student.Recordset.Fields("Clas4") = "‰œ«—œ" And Student.Recordset.Fields("Clas5") = "‰œ«—œ" Then
If MsgBox("«Ì‰ ﬁ—¬‰ ¬„Ê“ œ— ÂÌç ò·«”Ì ‘—ò  ‰„Ì ò‰œ ¬Ì« «œ«„Â „Ì œÂÌœ", vbCritical + vbYesNo, "ò·«”") = vbYes Then

GoTo 1
Else
Exit Sub

End If
End If


1

Student.Recordset.Fields("xkard") = "Kard-New"
Student.Recordset.Update
Beep
Call Command6_Click

End Sub

Private Sub List1_Click()
On Error Resume Next
Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" + Mid(List1.Text, 1, 7) + "%')"
Student.Refresh

End Sub

Private Sub Text1_Change()
On Error Resume Next

Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" & Text1.Text & "%') or nf like ('%" & Text1.Text & "%')"
Student.Refresh
End Sub
