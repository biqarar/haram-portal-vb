VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form EmtahanF 
   Caption         =   "��� ����� �������� ����� "
   ClientHeight    =   10650
   ClientLeft      =   3375
   ClientTop       =   2310
   ClientWidth     =   12480
   BeginProperty Font 
      Name            =   "B Homa"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Emtahan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10650
   ScaleWidth      =   12480
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
      Caption         =   "������ ���� ����"
      Height          =   5535
      Left            =   12480
      TabIndex        =   112
      Top             =   0
      Width           =   3495
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "����� ���"
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
         TabIndex        =   138
         Top             =   5160
         Width           =   615
      End
      Begin VB.Label Label78 
         AutoSize        =   -1  'True
         Caption         =   "����"
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
         TabIndex        =   137
         Top             =   4320
         Width           =   300
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�����"
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
         Left            =   2040
         TabIndex        =   136
         Top             =   4680
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "����� ������"
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
         Index           =   3
         Left            =   2040
         TabIndex        =   135
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "�� ����"
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
         TabIndex        =   134
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label74 
         AutoSize        =   -1  'True
         Caption         =   "���"
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
         TabIndex        =   133
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "����� ������"
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
         TabIndex        =   132
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         Caption         =   "������ �����"
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
         TabIndex        =   131
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "����� ���"
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
         TabIndex        =   130
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label72 
         AutoSize        =   -1  'True
         Caption         =   "���� ������"
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
         TabIndex        =   129
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         Caption         =   "�����"
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
         TabIndex        =   128
         Top             =   3360
         Width           =   495
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         Caption         =   "���� �����"
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
         TabIndex        =   127
         Top             =   3840
         Width           =   705
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         Caption         =   ": ��"
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
         TabIndex        =   126
         Top             =   360
         Width           =   270
      End
      Begin VB.Label Label68 
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
         TabIndex        =   125
         Top             =   4200
         Width           =   135
      End
      Begin VB.Label Label67 
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
         TabIndex        =   124
         Top             =   4680
         Width           =   135
      End
      Begin VB.Label Label63 
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
         TabIndex        =   123
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label62 
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
         TabIndex        =   122
         ToolTipText     =   "���� ����� ������ ���� ��� ����"
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label61 
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
         TabIndex        =   121
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label60 
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
         TabIndex        =   120
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label59 
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
         TabIndex        =   119
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label58 
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
         TabIndex        =   118
         Top             =   2640
         Width           =   135
      End
      Begin VB.Label Label57 
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
         TabIndex        =   117
         Top             =   3000
         Width           =   135
      End
      Begin VB.Label Label56 
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
         TabIndex        =   116
         Top             =   3360
         Width           =   135
      End
      Begin VB.Label Label55 
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
         TabIndex        =   115
         Top             =   3720
         Width           =   135
      End
      Begin VB.Label Label54 
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
         TabIndex        =   114
         Top             =   360
         Width           =   60
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
         TabIndex        =   113
         Top             =   5160
         Width           =   120
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Motor"
      Height          =   135
      Left            =   1560
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
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
         Connect         =   $"Emtahan.frx":08CA
         OLEDBString     =   $"Emtahan.frx":0953
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
         Connect         =   $"Emtahan.frx":09DC
         OLEDBString     =   $"Emtahan.frx":0A65
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
         Connect         =   $"Emtahan.frx":0AEE
         OLEDBString     =   $"Emtahan.frx":0B77
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
         Connect         =   $"Emtahan.frx":0C00
         OLEDBString     =   $"Emtahan.frx":0C89
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
         Connect         =   $"Emtahan.frx":0D12
         OLEDBString     =   $"Emtahan.frx":0D9B
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
         Connect         =   $"Emtahan.frx":0E24
         OLEDBString     =   $"Emtahan.frx":0EAD
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
         Connect         =   $"Emtahan.frx":0F36
         OLEDBString     =   $"Emtahan.frx":0FBF
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
      Begin MSAdodcLib.Adodc Setting 
         Height          =   330
         Left            =   360
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
         Connect         =   $"Emtahan.frx":1048
         OLEDBString     =   $"Emtahan.frx":10D1
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
         Connect         =   $"Emtahan.frx":115A
         OLEDBString     =   $"Emtahan.frx":11E3
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
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   1095
      Left            =   12840
      TabIndex        =   101
      Top             =   5880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��� ����"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   3840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Check3 
      Alignment       =   1  'Right Justify
      Caption         =   "���� ������� �� �� ���"
      Height          =   420
      Left            =   5160
      TabIndex        =   99
      Top             =   5160
      Width           =   1815
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   495
      Left            =   120
      TabIndex        =   97
      Top             =   5160
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   873
      Style           =   1
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "��� �����"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "����� ��� ���"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "����� ����� ��� ��� ���� ���� ����"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0E0FF&
      Caption         =   "����� �����"
      Height          =   495
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   8040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "����� ����� ���� ����"
      Height          =   375
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   7320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFF00&
      Caption         =   "�ǘ����"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000FF00&
      Caption         =   "��� �����"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   7080
      TabIndex        =   80
      Text            =   "�������"
      Top             =   5040
      Width           =   5295
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "Emtahan.frx":126C
      Height          =   4695
      Left            =   120
      TabIndex        =   46
      Top             =   5640
      Visible         =   0   'False
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   8281
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777152
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
      Caption         =   "���� ���� ��"
      ColumnCount     =   17
      BeginProperty Column00 
         DataField       =   "Parvande"
         Caption         =   "����� ������"
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
         Caption         =   "���"
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
         Caption         =   "��� �����ϐ�"
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
         DataField       =   "Ostad"
         Caption         =   "��� �����"
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
         DataField       =   "Tarh"
         Caption         =   "���"
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
         DataField       =   "Sal"
         Caption         =   "���"
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
         DataField       =   "Mah"
         Caption         =   "���"
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
         DataField       =   "Rooz"
         Caption         =   "���"
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
         DataField       =   "Noe"
         Caption         =   "���"
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
         DataField       =   "Elat"
         Caption         =   "���"
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
         DataField       =   "Tozih"
         Caption         =   "�������"
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
         DataField       =   "Clas"
         Caption         =   "����"
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
         DataField       =   "Op"
         Caption         =   "ǁ�����"
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
         DataField       =   "D"
         Caption         =   "����� ���"
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
         DataField       =   "Vazeyat"
         Caption         =   "����� �����"
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
         DataField       =   "Natije"
         Caption         =   "����� �����"
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
         DataField       =   "EMTAHANAT"
         Caption         =   "��������"
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
      EndProperty
   End
   Begin VB.Frame Frame11 
      Height          =   615
      Left            =   1560
      TabIndex        =   72
      Top             =   4440
      Width           =   5415
      Begin VB.ComboBox Combo5 
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   6
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "Emtahan.frx":1281
         Left            =   120
         List            =   "Emtahan.frx":1291
         TabIndex        =   85
         Text            =   "������ ����"
         Top             =   220
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "�����"
         Height          =   300
         Left            =   1440
         TabIndex        =   84
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "��� ������"
         Height          =   300
         Left            =   2880
         TabIndex        =   76
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   2160
         TabIndex        =   75
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "_"
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
         Left            =   3960
         TabIndex        =   74
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "������ �����"
         Height          =   300
         Left            =   4560
         TabIndex        =   73
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Frame Frame7 
      Height          =   855
      Left            =   7680
      TabIndex        =   71
      Top             =   11040
      Visible         =   0   'False
      Width           =   2535
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "����� ������"
         Height          =   300
         Left            =   1200
         TabIndex        =   79
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton OptionHEFZ 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "��� �����"
         Height          =   300
         Left            =   120
         TabIndex        =   78
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "������ ����"
      Height          =   3015
      Left            =   120
      TabIndex        =   53
      Top             =   120
      Width           =   3135
      Begin VB.CommandButton Command10 
         Caption         =   "<"
         Height          =   420
         Left            =   120
         TabIndex        =   109
         ToolTipText     =   "���� ����"
         Top             =   2520
         Width           =   375
      End
      Begin VB.CommandButton Command9 
         Caption         =   ">"
         Height          =   420
         Left            =   2160
         TabIndex        =   108
         ToolTipText     =   "���� ����"
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox Text5 
         Height          =   420
         Left            =   600
         TabIndex        =   107
         Text            =   "����� �� ����"
         ToolTipText     =   "����� �� �� ���ӡ ��� ����� � ��͡ ����"
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   2640
         X2              =   2640
         Y1              =   120
         Y2              =   3000
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   300
         Left            =   2040
         TabIndex        =   103
         Top             =   2160
         Width           =   525
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "-  "
         DataField       =   "Tozih"
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
         TabIndex        =   102
         Top             =   2160
         Width           =   225
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   330
         Left            =   2760
         TabIndex        =   67
         Top             =   1560
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
         Left            =   2760
         TabIndex        =   66
         Top             =   1200
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
         TabIndex        =   65
         Top             =   1800
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
         TabIndex        =   64
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
         Left            =   2760
         TabIndex        =   63
         Top             =   2040
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
         TabIndex        =   62
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
         TabIndex        =   61
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
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   240
         TabIndex        =   60
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   330
         Left            =   2040
         TabIndex        =   59
         Top             =   1800
         Width           =   405
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   330
         Left            =   2700
         TabIndex        =   58
         Top             =   840
         Width           =   405
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "��� �����"
         Height          =   330
         Left            =   2040
         TabIndex        =   57
         Top             =   1440
         Width           =   585
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   330
         Left            =   2040
         TabIndex        =   56
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "���"
         Height          =   330
         Left            =   2040
         TabIndex        =   55
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�� ����"
         Height          =   330
         Index           =   2
         Left            =   2040
         TabIndex        =   54
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "������ ���� ����"
      Height          =   1695
      Left            =   7080
      TabIndex        =   41
      Top             =   960
      Width           =   5295
      Begin VB.ComboBox Combo11 
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
         TabIndex        =   139
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "���� ���� ��"
         Height          =   300
         Left            =   240
         TabIndex        =   86
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "����� ���� ��"
         Height          =   420
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1455
      End
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
         Left            =   2280
         TabIndex        =   4
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Text3 
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
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text18 
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
         Left            =   2280
         TabIndex        =   6
         Text            =   "��� ����"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ComboBox Combo3 
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
         Left            =   2280
         TabIndex        =   2
         Text            =   "������ ����"
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   1800
         TabIndex        =   70
         Top             =   1200
         Width           =   120
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "�����"
         Height          =   300
         Left            =   1680
         TabIndex        =   69
         Top             =   840
         Width           =   315
      End
      Begin VB.Label Label11 
         Caption         =   "����"
         Height          =   300
         Left            =   4800
         TabIndex        =   50
         Top             =   840
         Width           =   315
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   300
         Left            =   1320
         TabIndex        =   44
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "���"
         Height          =   300
         Left            =   4920
         TabIndex        =   43
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "����� ���� ��� ���� ���"
         Height          =   300
         Left            =   3600
         TabIndex        =   42
         Top             =   1320
         Width           =   1530
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "������ ���� ����"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   3360
      TabIndex        =   30
      Top             =   120
      Width           =   3615
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
         TabIndex        =   96
         Top             =   2400
         Width           =   135
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "����� ����"
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
         TabIndex        =   95
         Top             =   2520
         Width           =   765
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
         TabIndex        =   94
         Top             =   2040
         Width           =   135
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "�������"
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
         TabIndex        =   93
         Top             =   2160
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "����� ������"
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
         TabIndex        =   40
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "���"
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
         TabIndex        =   39
         Top             =   600
         Width           =   195
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "��� �����ϐ�"
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
         TabIndex        =   38
         Top             =   960
         Width           =   870
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "��� ���"
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
         Index           =   0
         Left            =   2280
         TabIndex        =   37
         Top             =   1320
         Width           =   450
      End
      Begin VB.Label Label38 
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
         TabIndex        =   36
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label37 
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
         TabIndex        =   35
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label36 
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
         TabIndex        =   34
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label34 
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
         TabIndex        =   33
         Top             =   1320
         Width           =   135
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "��� �����"
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
         TabIndex        =   32
         Top             =   1800
         Width           =   555
      End
      Begin VB.Label Label32 
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
         TabIndex        =   31
         Top             =   1680
         Width           =   135
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "����"
      Enabled         =   0   'False
      Height          =   495
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   -120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      Caption         =   "������ ������"
      Height          =   2295
      Left            =   7080
      TabIndex        =   25
      Top             =   2640
      Width           =   5295
      Begin VB.ComboBox Combo10 
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
         Left            =   2400
         TabIndex        =   110
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox Combo9 
         BackColor       =   &H00FFFFFF&
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
         Left            =   3840
         TabIndex        =   106
         Top             =   1800
         Width           =   615
      End
      Begin VB.ComboBox Combo8 
         BackColor       =   &H00FFFFFF&
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
         Left            =   3240
         TabIndex        =   105
         Top             =   1800
         Width           =   615
      End
      Begin VB.ComboBox Combo7 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2400
         TabIndex        =   104
         Top             =   1800
         Width           =   855
      End
      Begin VB.ComboBox Combo6 
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
         Left            =   2400
         TabIndex        =   98
         Top             =   1320
         Width           =   1815
      End
      Begin VB.ComboBox Combo4 
         Height          =   420
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.Frame Frame10 
         Height          =   855
         Left            =   120
         TabIndex        =   48
         Top             =   1320
         Width           =   2175
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Caption         =   "����� ����"
            Height          =   300
            Left            =   120
            TabIndex        =   82
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option6 
            Alignment       =   1  'Right Justify
            Caption         =   "�����"
            Height          =   300
            Left            =   1320
            TabIndex        =   81
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option4 
            Alignment       =   1  'Right Justify
            Caption         =   "������ 1"
            Height          =   300
            Left            =   1200
            TabIndex        =   11
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton Option5 
            Alignment       =   1  'Right Justify
            Caption         =   "������ 2"
            Height          =   300
            Left            =   240
            TabIndex        =   12
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.Frame Frame9 
         Height          =   615
         Left            =   120
         TabIndex        =   47
         Top             =   720
         Width           =   2175
         Begin VB.OptionButton Option2 
            Alignment       =   1  'Right Justify
            Caption         =   "����� ���"
            Height          =   225
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            Caption         =   "���� ���"
            Height          =   300
            Left            =   1080
            TabIndex        =   9
            Top             =   240
            Width           =   855
         End
      End
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
         Left            =   2400
         TabIndex        =   7
         Text            =   "������ ����"
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "����� ������"
         Height          =   300
         Left            =   4440
         TabIndex        =   49
         Top             =   1920
         Width           =   720
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "��� ������"
         Height          =   300
         Left            =   4560
         TabIndex        =   29
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "��� �����"
         Height          =   300
         Left            =   4560
         TabIndex        =   28
         Top             =   1440
         Width           =   660
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "����� ������"
         Height          =   300
         Left            =   4440
         TabIndex        =   27
         Top             =   960
         Width           =   810
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "����� ���"
         Height          =   300
         Left            =   1560
         TabIndex        =   26
         Top             =   360
         Width           =   600
      End
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   420
      Left            =   7080
      TabIndex        =   1
      Top             =   480
      Width           =   5295
   End
   Begin VB.Frame Frame1 
      Caption         =   "�����"
      Height          =   1335
      Left            =   1560
      TabIndex        =   20
      Top             =   3120
      Width           =   5415
      Begin VB.TextBox Text11 
         BackColor       =   &H00FFFFFF&
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
         Left            =   3840
         TabIndex        =   14
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text9 
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
         Left            =   3840
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1920
         TabIndex        =   17
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1920
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00FFFFFF&
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
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H00FFFFFF&
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
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   975
      End
      Begin VB.Line Line5 
         X1              =   3960
         X2              =   3960
         Y1              =   240
         Y2              =   360
      End
      Begin VB.Line Line4 
         X1              =   960
         X2              =   960
         Y1              =   240
         Y2              =   360
      End
      Begin VB.Line Line3 
         X1              =   960
         X2              =   3960
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line1 
         X1              =   3720
         X2              =   3720
         Y1              =   360
         Y2              =   1200
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "�����"
         Height          =   300
         Left            =   4920
         TabIndex        =   52
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   300
         Left            =   5040
         TabIndex        =   51
         Top             =   840
         Width           =   285
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "����� � ����� "
         Height          =   300
         Left            =   1080
         TabIndex        =   45
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "���"
         Height          =   300
         Left            =   3120
         TabIndex        =   24
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   300
         Left            =   3120
         TabIndex        =   23
         Top             =   960
         Width           =   450
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "�����"
         Height          =   300
         Left            =   1200
         TabIndex        =   22
         Top             =   480
         Width           =   435
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Emtahan.frx":12AD
      Height          =   4695
      Left            =   120
      TabIndex        =   21
      Top             =   5640
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   8281
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
      Caption         =   "���� ������"
      ColumnCount     =   27
      BeginProperty Column00 
         DataField       =   "Parvande"
         Caption         =   "������"
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
         Caption         =   "���"
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
         Caption         =   "��� �����ϐ�"
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
         Caption         =   "��� ���"
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
         Caption         =   "����� ����"
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
         Caption         =   "����� ��������"
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
         Caption         =   "�����"
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
         Caption         =   "����"
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
         Caption         =   "����"
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
         Caption         =   "�� ���"
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
         Caption         =   "����� ��� ����"
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
         Caption         =   "����� ����"
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
         Caption         =   "����� �������"
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
         Caption         =   "�������"
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
         Caption         =   "�����"
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
         Caption         =   "�������"
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
         Caption         =   "���� ����"
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
         Caption         =   "����� �����"
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
         Caption         =   "�Ә�"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Emtahan.frx":12C3
      Height          =   4695
      Left            =   120
      TabIndex        =   77
      Top             =   5640
      Visible         =   0   'False
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   8281
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
      Caption         =   "����� ������"
      ColumnCount     =   24
      BeginProperty Column00 
         DataField       =   "Parvande"
         Caption         =   "������"
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
         Caption         =   "�� ����"
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
         Caption         =   "��� "
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
         Caption         =   "�� �������"
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
         Caption         =   "������ ������"
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
         Caption         =   "����� ������"
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
         Caption         =   "���"
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
         Caption         =   "������"
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
         Caption         =   "�����"
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
         Caption         =   "����� � �����"
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
         Caption         =   "������ �����"
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
         Caption         =   "ǁ�����"
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
         Caption         =   "����� ���"
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
         Caption         =   "�����"
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
         Caption         =   "����� ����"
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
         Caption         =   "���"
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
         Caption         =   "���� � ����� ���"
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
         Caption         =   "��� ������"
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
         Caption         =   "����"
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
         Caption         =   "�����"
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
         Caption         =   "�����"
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
         Caption         =   "���� �����"
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
         Caption         =   "�������"
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
         Caption         =   "�� ������"
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
   Begin ComctlLib.StatusBar Sb1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   90
      Top             =   10275
      Width           =   12480
      _ExtentX        =   22013
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "����� ����"
            TextSave        =   "����� ����"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "����� �����"
            TextSave        =   "����� �����"
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
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "Ϙ�� ���� ��� �����"
      Height          =   300
      Left            =   8520
      TabIndex        =   111
      Top             =   0
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "���� ���� ��"
      Height          =   300
      Left            =   7200
      TabIndex        =   89
      Top             =   120
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   300
      Left            =   8400
      TabIndex        =   88
      Top             =   120
      Width           =   75
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "����� �� ��� � ��� �����ϐ� � ����� ������"
      Height          =   300
      Left            =   9720
      TabIndex        =   68
      Top             =   120
      Width           =   2580
   End
   Begin VB.Menu mnuhome 
      Caption         =   "#"
   End
   Begin VB.Menu MNUPARVAND 
      Caption         =   "������"
      Begin VB.Menu MNUSABTNOMARAT 
         Caption         =   "��� �����"
      End
      Begin VB.Menu mnuedite 
         Caption         =   "����� �����"
      End
      Begin VB.Menu fddg 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu MNUDELL 
         Caption         =   "��� ����"
      End
   End
   Begin VB.Menu mnuklarnamekelass 
      Caption         =   "������� �����"
   End
End
Attribute VB_Name = "EmtahanF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Check2_Click()
If Check2.Value = 0 Then

Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label38.Caption + "%') and clas like ('%" + Combo1.Text + "%') and noe like ('%" + "���" + "%') and emtahanat like ('%" + "����� ����" + "%') "

Qeybat.Refresh
DataGrid3.Visible = True

Label20.Caption = Qeybat.Recordset.RecordCount

End If

If Check2.Value = 1 Then

Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label38.Caption + "%') and noe like ('%" + "���" + "%') "

Qeybat.Refresh
DataGrid3.Visible = True

Label20.Caption = Qeybat.Recordset.RecordCount

End If

End Sub

Private Sub Combo1_Change()
mclass.Refresh
mclass.RecordSource = "seleCt * from mclass where kodclass like ('%" + Combo1.Text + "%')"
mclass.Refresh
End Sub

Private Sub Combo1_Click()
mclass.Refresh
mclass.RecordSource = "seleCt * from mclass where kodclass like ('%" + Combo1.Text + "%')"
mclass.Refresh
End Sub


Private Sub Combo11_Click()
If Combo11.Text = "������" Then

Text6.Enabled = False
Text6.BackColor = &HC0C0C0


End If
If Combo11.Text = "��ѐ���" Then



Text6.Enabled = True
Text6.BackColor = &H80000005




End If


End Sub

Private Sub Combo3_Click()
If Combo3.Text = "����� ��� 1" Or Combo3.Text = "����� ��� 2" Then
Label17.Caption = "����"
Else
Label17.Caption = "�����"
End If



If Combo3.Text = "�� �����" Then
Combo2.Clear
Combo2.Enabled = True
'Combo2.AddItem ("������ �����")
'Combo2.Text = Combo2.List(0)
Check1.Enabled = False
Check1.Value = 0


Combo10.Enabled = True
Combo10.Clear
Combo10.AddItem ("���� ����")
Combo10.AddItem ("����� ����")
Combo10.Text = Combo10.List(0)


'ʘ�� ��� ���
'Text12.Enabled = False
'Text12.BackColor = &HC0C0C0
'
'Text10.Enabled = False
'Text10.BackColor = &HC0C0C0
'
'Text8.Enabled = False
'Text8.BackColor = &HC0C0C0
'
'Text6.Enabled = False
'Text6.BackColor = &HC0C0C0
'

'ʘ�� ��� �� ����� �  ���� �����


'Text9.Enabled = False
'Text9.BackColor = &HC0C0C0

'Text11.Enabled = False
'Text11.BackColor = &HC0C0C0



'Option5.Enabled = False







Exit Sub

End If
If Combo3.Text = "����� � ������" Then
    Combo2.Clear
Combo2.Enabled = True
Combo2.AddItem ("������ ����")
Combo2.Text = Combo2.List(0)

   'ʘ�� ��� ���
'Text12.Enabled = False
'Text12.BackColor = &HC0C0C0
'
'Text10.Enabled = False
'Text10.BackColor = &HC0C0C0

'Text8.Enabled = False
'Text8.BackColor = &HC0C0C0

'Text6.Enabled = False
'Text6.BackColor = &HC0C0C0
'ʘ�� ��� �� ����� �  ���� �����
'Text9.Enabled = False
'Text9.BackColor = &HC0C0C0
''
'Text11.Enabled = True
'Text11.BackColor = &H80000005

Combo10.AddItem ("���� ��� ��� 1")
Combo10.AddItem ("����� ��� 1")
Combo10.AddItem ("���� ��� ��� 2")
Combo10.AddItem ("����� ��� 2")
Combo10.AddItem ("���� ��� ��� 3")
Combo10.AddItem ("����� ��� 3")
Combo10.Text = "������ ����"




Exit Sub
End If




If Combo3.Text = "���� �����" Or Combo3.Text = "����� ��� 1" Or Combo3.Text = "����� ��� 2" Then

'Text11.Enabled = True
'Text11.BackColor = &H80000005

Option5.Enabled = False


Combo10.Enabled = True
Combo10.Clear
Combo10.AddItem ("���� ����")
Combo10.AddItem ("����� ����")
Combo10.Text = Combo10.List(0)


'ʘ�� ��� ���
'Text12.Enabled = False
'Text12.BackColor = &HC0C0C0

'Text10.Enabled = False
'Text10.BackColor = &HC0C0C0

'Text8.Enabled = False
'Text8.BackColor = &HC0C0C0

'Text6.Enabled = False
'Text6.BackColor = &HC0C0C0


'ʘ�� ��� �� ����� �  ���� �����


'Text9.Enabled = True
'Text9.BackColor = &H80000005

'Text11.Enabled = True
'Text11.BackColor = &H80000005



'Check1.Enabled = True








Combo2.Clear
Combo2.Enabled = True
Combo2.Text = "������ ����"
Combo2.AddItem ("���� ����")
Combo2.AddItem ("����� ����")
Else

'Text11.Enabled = False
'Text11.BackColor = &H80000005


Option5.Enabled = True
'ʘ�� ��� ���
'Text12.Enabled = True
'Text12.BackColor = &H80000005

'Text10.Enabled = True
'Text10.BackColor = &H80000005
'
'Text8.Enabled = True
'Text8.BackColor = &H80000005
'
'Text6.Enabled = True
'Text6.BackColor = &H80000005
'
'
''ʘ�� ��� �� ����� �  ���� �����


'Text9.Enabled = False
'Text9.BackColor = &HC0C0C0

'Text11.Enabled = False
'Text11.BackColor = &HC0C0C0





'Check1.Enabled = True








Combo2.Clear
Combo2.Enabled = False
End If


End Sub


Private Sub Command1_Click()
If Command1.Caption = "����� ���� ��" Then


Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label38.Caption + "%') and clas like ('%" + Combo1.Text + "%') and noe like ('%" + "" + "%') and emtahanat like ('%" + "����� ����" + "%') "

Qeybat.Refresh
DataGrid3.Visible = True

Label20.Caption = Qeybat.Recordset.RecordCount


If Check2.Value = 1 Then

Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label38.Caption + "%') and noe like ('%" + "" + "%') "

Qeybat.Refresh
DataGrid3.Visible = True

Label20.Caption = Qeybat.Recordset.RecordCount

End If




Command1.Caption = "������ ���� ����"
Else
Command1.Caption = "����� ���� ��"
DataGrid3.Visible = False
DataGrid2.Visible = True
End If

End Sub


Private Sub Command10_Click()
On Error GoTo 9898
GoTo 9999
9898:
MsgBox "���� ���� ���� �����", vbCritical + vbOKOnly, "���"
Exit Sub
9999:


mclass.Recordset.MoveNext



End Sub

Private Sub Command2_Click()
'If Entekhab.SB.Panels(1).Text = "������" Then Exit Sub
'If MsgBox("����� ��� ������ �� ��� ����� �����", vbQuestion + vbYesNo, "�����") = vbYes Then
'GoTo 8
'Else
'Exit Sub
'End If
'Exit Sub
'8

If Option4.Value = False And Option5.Value = False And Option6.Value = False Then
MsgBox "����� ����� � ������ ���� ���� �� ������ ����", vbCritical + vbOKOnly, "���"
Exit Sub
End If



Emtahan.Refresh
Emtahan.Recordset.AddNew




Dim TaSe As String
If Option6.Value = True Then TaSe = "Q"
If Option4.Value = True Then TaSe = "T1"
If Option5.Value = True Then TaSe = "T2"


If Option1.Value = True Then Emtahan.Recordset.Fields("kode") = "P" & Label38.Caption & "J" & Combo4.Text & "NP5" & TaSe & Combo1.Text
If Option2.Value = True Then Emtahan.Recordset.Fields("kode") = "P" & Label38.Caption & "J" & Combo4.Text & "NP1" & TaSe & Combo1.Text

Dim Kode_yadam_nare As String
Kode_yadam_nare = Emtahan.Recordset.Fields("kode")





If Combo3.Text = "�� �����" Then Emtahan.Recordset.Fields("kode") = "P" & Label38.Caption & "RO" & TaSe & Combo1.Text
If Combo3.Text = "���� �����" Then Emtahan.Recordset.Fields("kode") = "P" & Label38.Caption & "RA" & TaSe & Combo1.Text
If Combo3.Text = "����� ��� 1" Then Emtahan.Recordset.Fields("kode") = "P" & Label38.Caption & "TJ1" & TaSe & Combo1.Text
If Combo3.Text = "����� ��� 2" Then Emtahan.Recordset.Fields("kode") = "P" & Label38.Caption & "TJ2" & TaSe & Combo1.Text
If Combo3.Text = "����� � ������" Then Emtahan.Recordset.Fields("kode") = "P" & Label38.Caption & "TR" & TaSe & Combo1.Text




'��� ���� ��

 Dim XQeybatCont As Integer
 '
 '���� ����� ���� ��� ���� � ��� ���� �� ����� ���� � �� ���� ���� ���
 '
 '
 XQeybatCont = 0
 
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label38.Caption + "%') and emtahanat like ('%" + "����� ����" + "%') and noe like ('%" + "��� ����" + "%')"
Qeybat.Refresh
 '
 Emtahan.Recordset.Fields("qeyremovajah") = Qeybat.Recordset.RecordCount
 
 
 
 '��� ��� ��� � ����
 
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label38.Caption + "%') and emtahanat like ('%" + "����� ����" + "%') and noe like ('%" + "���� ����" + "%')"
Qeybat.Refresh
 '
 
 

 XQeybatCont = Qeybat.Recordset.RecordCount
 
  
Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label38.Caption + "%') and emtahanat like ('%" + "����� ����" + "%') and noe like ('%" + "�����" + "%')"
Qeybat.Refresh
 
 
  XQeybatCont = XQeybatCont + Qeybat.Recordset.RecordCount
 
 Emtahan.Recordset.Fields("movajah") = XQeybatCont
  
 
 
 
 
 
 Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label38.Caption + "%') and emtahanat like ('%" + "����� ����" + "%')"
Qeybat.Refresh
 
 '
 
For I = 1 To Qeybat.Recordset.RecordCount
Qeybat.Recordset.Fields("Emtahanat") = "P" & Label38.Caption & "J" & Combo4.Text & "NP5" & TaSe
Qeybat.Recordset.Update
Qeybat.Recordset.MoveNext
Next I


'����� ��� ��



Emtahan.Recordset.Fields("parvande") = Label38.Caption
Emtahan.Recordset.Fields("kodclass") = lkodclass.Caption
Emtahan.Recordset.Fields("tarh") = Combo3.Text
Emtahan.Recordset.Fields("kolmahfozat") = Combo11.Text
Emtahan.Recordset.Fields("mahdodeemtahan") = Combo10.Text
Emtahan.Recordset.Fields("temtahan") = Combo7.Text & "/" & Combo8.Text & "/" & Combo9.Text
Emtahan.Recordset.Fields("hefz") = Text8.Text
Emtahan.Recordset.Fields("mafahim") = Text6.Text
Emtahan.Recordset.Fields("mostamar") = Text10.Text
Emtahan.Recordset.Fields("ee") = Text12.Text
Emtahan.Recordset.Fields("enahaee") = ""
Emtahan.Recordset.Fields("op") = Entekhab.SB.Panels(1).Text
Emtahan.Recordset.Fields("d") = Taqvim.Label1.Caption
Emtahan.Recordset.Fields("momtahen") = Combo6.Text
Emtahan.Recordset.Fields("tqeybat") = Text18.Text
Emtahan.Recordset.Fields("kasremtiaz") = ""
Emtahan.Recordset.Fields("joze") = Combo4.Text
Emtahan.Recordset.Fields("enahaee") = Label29.Caption
Emtahan.Recordset.Fields("rotbe") = "���� ����"
If Option6.Value = True Then Emtahan.Recordset.Fields("vazeyat") = Option6.Caption
If Option4.Value = True Then Emtahan.Recordset.Fields("vazeyat") = Option4.Caption
If Option5.Value = True Then Emtahan.Recordset.Fields("vazeyat") = Option5.Caption


 'Label1.Caption
'Emtahan.Recordset.Fields("rotbe") = "??"
Emtahan.Recordset.Fields("tozih") = Text1.Text

If Option1.Value = True Then Emtahan.Recordset.Fields("nimpayan") = "���� ���"
If Option2.Value = True Then Emtahan.Recordset.Fields("nimpayan") = "����� ���"

Emtahan.Recordset.Fields("kasremtiaz") = Label22.Caption

Emtahan.Recordset.Fields("katbi") = Text11.Text
Emtahan.Recordset.Fields("shafahi") = Text9.Text

Emtahan.Recordset.Fields("Chap") = "�ǁ ����"
Emtahan.Recordset.Fields("dateofChap") = "0000-00-00"

Emtahan.Recordset.Update
Emtahan.Refresh





Command2.Enabled = False

MsgBox "����� �� ������ �� ����� ��� ����", vbInformation + vbOKOnly, "��������"



Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Combo11.Text = ""
Text1.Text = "�������"
Text18.Text = "��� ����"
Text6.Text = ""
Combo4.Text = ""
Combo5.Text = "������ ����"


Option1.Value = False
Option2.Value = False
Option4.Value = False
Option5.Value = False
Option6.Value = False


Label20.Caption = "-"
Label29.Caption = "-"
Label22.Caption = "-"


If Check3.Value = 1 Then


Karname.Show
Karname.Text2.Text = Kode_yadam_nare
Karname.DataGrid1.Visible = True
Karname.Command1.Enabled = True

End If


End Sub

Private Sub Command3_Click()
If Entekhab.SB.Panels(1).Text = "������ ����" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "emtahan-delete" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub

14082513
If MsgBox("��� �� ������ ����� ��� ��� �� ��� ����", vbQuestion + vbYesNo, "��� ����") = vbYes Then
On Error Resume Next

Emtahan.Recordset.Delete




End If


End Sub


Private Sub Command4_Click()







Dim TaSe As String
If Option6.Value = True Then TaSe = "Q"
If Option4.Value = True Then TaSe = "T1"
If Option5.Value = True Then TaSe = "T2"


If Option1.Value = True Then Emtahan.Recordset.Fields("kode") = "P" & Label38.Caption & "J" & Combo4.Text & "NP5" & TaSe & Combo1.Text
If Option2.Value = True Then Emtahan.Recordset.Fields("kode") = "P" & Label38.Caption & "J" & Combo4.Text & "NP1" & TaSe & Combo1.Text






If Combo3.Text = "�� �����" Then Emtahan.Recordset.Fields("kode") = "P" & Label38.Caption & "RO" & Combo1.Text
If Combo3.Text = "���� �����" Then Emtahan.Recordset.Fields("kode") = "P" & Label38.Caption & "RA" & Combo1.Text
If Combo3.Text = "����� ��� 1" Then Emtahan.Recordset.Fields("kode") = "P" & Label38.Caption & "TJ1" & Combo1.Text
If Combo3.Text = "����� ��� 2" Then Emtahan.Recordset.Fields("kode") = "P" & Label38.Caption & "TJ2" & Combo1.Text
If Combo3.Text = "����� � ������" Then Emtahan.Recordset.Fields("kode") = "P" & Label38.Caption & "TR" & Combo1.Text




'��� ���� ��

Qeybat.Refresh
Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label38.Caption + "%') and emtahanat like ('%" + "����� ����" + "%')"
Qeybat.Refresh
 
For I = 1 To Qeybat.Recordset.RecordCount
Qeybat.Recordset.Fields("Emtahanat") = "P" & Label38.Caption & "J" & Combo4.Text & "NP5" & TaSe
Qeybat.Recordset.Update
Qeybat.Recordset.MoveNext
Next I


'����� ��� ��



Label38.Caption = Emtahan.Recordset.Fields("parvande")
                                                        
                                                        
                                                        
                                                        lkodclass.Caption = Emtahan.Recordset.Fields("kodclass")
Combo3.Text = Emtahan.Recordset.Fields("tarh")
                                                        

Combo11.Text = Emtahan.Recordset.Fields("kolmahfozat")
Combo10.Text = Emtahan.Recordset.Fields("mahdodeemtahan")
 Text7.Text = Emtahan.Recordset.Fields("temtahan")
Text8.Text = Emtahan.Recordset.Fields("hefz")
  Text6.Text = Emtahan.Recordset.Fields("mafahim")
 Text10.Text = Emtahan.Recordset.Fields("mostamar")
 Text12.Text = Emtahan.Recordset.Fields("ee")
                                                    Emtahan.Recordset.Fields("enahaee") = ""
                                                    Emtahan.Recordset.Fields("op") = Entekhab.SB.Panels(1).Text
                                                    Emtahan.Recordset.Fields("d") = Taqvim.Label1.Caption
  Combo6.Text = Emtahan.Recordset.Fields("momtahen")
  Text18.Text = Emtahan.Recordset.Fields("tqeybat")
                                                    Emtahan.Recordset.Fields("kasremtiaz") = ""
 Combo4.Text = Emtahan.Recordset.Fields("joze")
 Label29.Caption = Emtahan.Recordset.Fields("enahaee")



If Option6.Value = True Then Emtahan.Recordset.Fields("vazeyat") = Option6.Caption
If Option4.Value = True Then Emtahan.Recordset.Fields("vazeyat") = Option4.Caption
If Option5.Value = True Then Emtahan.Recordset.Fields("vazeyat") = Option5.Caption


 'Label1.Caption
'Emtahan.Recordset.Fields("rotbe") = "??"
Emtahan.Recordset.Fields("tozih") = Text1.Text

If Option1.Value = True Then Emtahan.Recordset.Fields("nimpayan") = "���� ���"
If Option2.Value = True Then Emtahan.Recordset.Fields("nimpayan") = "����� ���"

  Label22.Caption = Emtahan.Recordset.Fields("kasremtiaz")

  Text11.Text = Emtahan.Recordset.Fields("katbi")
 Text9.Text = Emtahan.Recordset.Fields("shafahi")












End Sub


Private Sub Command5_Click()
If Entekhab.SB.Panels(1).Text = "������ ����" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "emtahan-newsabt" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
Call Label39_Click
Exit Sub



Command2.Enabled = False

If Me.lkodclass.Caption = "�����" Then
MsgBox "���� ���� ���� �� ����� ����", vbCritical + vbOKOnly, "���"
Exit Sub
End If




If Combo7.Text = "" Or Combo8.Text = "" Or Combo9.Text = "" Then
MsgBox "����� ������ �� �� ���� ���� ���� ����", vbCritical + vbOKOnly, "���"
Exit Sub
End If

 
If Text18.Text = "��� ����" Then
MsgBox "����� ���� �� �� ���� ����", vbCritical + vbOKOnly, "���"
Exit Sub
End If



'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%5 ��� �� �����
If Combo3.Text = "�� �����" Then
If Text9.Text = "" Or Val(Text9.Text) > 20 Then
MsgBox "���� ������ ����� �� ����� ����", vbCritical + vbOKOnly, "���"
Exit Sub
End If
Label29.Caption = Text9.Text '��� ������ ����� � �� ����� ��� ��� � ���� ����  � ���� �� ����� ���� ����� �� ����
Label22.Caption = "�����" '��� ������ �� �����

If Val(Label29.Caption) >= 15 Then Option6.Value = True '���� ����� ����� ���� ��� ��� �� ���� ���

If Val(Label29.Caption) < 15 Then Option4.Value = True '���� ������ ���� ��� ��� ���� ������



GoTo 1 '��� �� ����� �� ��� ����� �� ��� ���� �� ��� ��� ȁ��


End If

'��� ����� � ������

         If Combo3.Text = "����� � ������" Then
           If Text11.Text = "" Or Val(Text11.Text) > 20 Then
            MsgBox "����� �� ���� ���� ����", vbCritical + vbOKOnly, "���"
            Exit Sub
            Else
            '���� ������ �� �� ��� ��� ���� ���� ����
            GoTo 25
            ' ������ �� ��� ��� ����� �� ���
            
            
            End If
        
         End If
         
         
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%5 ��� �� �� �����

If Combo3.Text = "���� �����" Or Combo3.Text = "����� ��� 1" Or Combo3.Text = "����� ��� 2" Or Combo3.Text = "�� �����" Then
If Text9.Text = "" Or Val(Text9.Text) > 20 Or Text11.Text = "" Or Val(Text11.Text) > 20 Then
MsgBox "����� �� ���� ���� ����", vbCritical + vbOKOnly, "���"
Exit Sub
End If


Label29.Caption = Val(Text9.Text) + Val(Text11.Text) '������ ���� � ����� ���� ���� ��� ����

Label22.Caption = "�����" '����� ��� ������ �����

If Val(Label29.Caption) >= 16 Then Option6.Value = True '�� ����� ���� �� ���
If Val(Label29.Caption) < 16 Then Option4.Value = True '�� ������ �� �� �� ���

GoTo 1 '���� �� ��� ��� ȁ��

'���� �� ������ �� ��� �� ����� � ���� ����� � ����� � ������ ����� �� ��� �����
'�� ���� ���� �� ���� ��� ���� � � ��� ���� ���� ������
'��� ���� ������� ���� �� ��� ������ ��� ����� ��� ��� ������ ��
'��� ���� �� �� else���� �� ���
Else '������� ���� ���� �� ����� � ���� ����� � ����� � ����� � ������ ���
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%5 ��� ���




'��� ���



    'If Label20.Caption = "-" Then 'ǐ� ����� ���� �� ���� ���� ��� ���� ����� �� ��� � ������ ���
'Qeybat.Refresh
'Qeybat.RecordSource = "select * from qeybat where parvande like ('%" + Label38.Caption + "%') and clas like ('%" + Combo1.Text + "%') and noe like ('%" + "���" + "%') and emtahanat like ('%" + "����� ����" + "%') "
'Qeybat.Refresh
'DataGrid3.Visible = True
'Label20.Caption = Qeybat.Recordset.RecordCount
'Text18.Text = Qeybat.Recordset.RecordCount
'End If




If Combo4.Text = "" Then '���� ����� ���
MsgBox "����� ��� �� ��� ����", vbCritical + vbOKOnly, "���"

Exit Sub

End If

If Option1.Value = False And Option2.Value = False Then '���� ���� � ����� ��� ����
MsgBox "���� ��� � �� ����� ��� ���� ������ �� ���� ����", vbCritical + vbOKOnly, "���"
Exit Sub
End If



'���� �����
If Text8.Text = "" Or Val(Text8.Text) > 15 Or Text6.Text = "" Or Val(Text6.Text) > 2 Or Text10.Text = "" Or Val(Text10.Text) > 2 Or Text12.Text = "" Or Val(Text12.Text) > 1 Then
MsgBox "����� �� ���� ���� ����", vbCritical + vbOKOnly, "���"
Exit Sub
End If


End If ' ���� ���� �� ����� � ���� ����� � ����� � ������ ���� �� ����



'sssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssss   ��� ����

'������� ����� �� ��� ������ � ����
Label29.Caption = (Val(Text8.Text) + Val(Text6.Text) + Val(Text10.Text) + Val(Text12.Text)) - Val(Label22.Caption) ' ��� ����




If Val(Combo4.Text) >= 1 And Val(Combo4.Text) <= 9 Or Val(Combo4.Text) = 30 Then 'ǐ� ���� ���� ��� ����� 1 � 10 ��� ����� �� ���� ���� ����� �����

        If Val(Label29.Caption) >= 17 Then '���� ����� �� 17 ���� �������� ���
        Option6.Value = True



Else '������
'���� ��� ����� 1 �10 ���� ������ ����� ���



        Dim TaS As String
        If Option1.Value = True Then TaS = Option1.Caption '���� ���
        If Option2.Value = True Then TaS = Option2.Caption '����� ���
            '�� ����� ����� �� ��� ����� ��� ������ �ј� ���� ���
        Emtahan.Refresh
        Emtahan.RecordSource = "select * from emtahan where parvande like ('%" + Label38.Caption + "%') and nimpayan like ('%" + TaS + "%') and joze like ('%" + Combo4.Text + "%') and vazeyat like ('%" + "�����" + "%')"
        Emtahan.Refresh
        '
            If Emtahan.Recordset.BOF = True Or Student.Recordset.EOF = True Then
            '���� ��� �� ��� ������� �ј� ���� ��� � ����� ��� ��� �� ���� ������1 ���� ��� ������
            Option4.Value = True 'sabt tajdidi 1
            GoTo 4
            '��� ��� ������ ���� �� ���� �� �� ��� ����� �� ����� 2 ���� ��� ���
            End If

    If Emtahan.Recordset.RecordCount <= 2 Then '���� ���� ���� �� ��� ������ �� ��� 1 ��� �ј� ���� ���
        Option5.Value = True '����� 2 ���� �� ���
        GoTo 2

    Else '���� ��� �� 2 ��� �� ��� ������ �ј� ���� ���
        MsgBox "��� ���� ���� ��� �� 2 ���  �� ������ �ј� ���� ��� ���� ��� ���� ���� �����", vbCritical + vbOKOnly, "���"
        Exit Sub

    End If
End If


End If '     ���� ��� 1 �� 10 �� ����


2: '����� 2
4: '����� 1



'ǐ� ���� ���� ��� ����� 11 �� 20 ��� ������ ������ ���� ����� ����
If Val(Combo4.Text) >= 10 And Val(Combo4.Text) <= 19 Then

If Val(Label29.Caption) >= 16 Then '����� �� 16 ���� �� ��� �����
Option6.Value = True



Else '������





If Option1.Value = True Then TaS = Option1.Caption
If Option2.Value = True Then TaS = Option2.Caption
'����� �� ��� �� ����� ��� ��� �� ��� ������ �ј� ���� ���
Emtahan.Refresh
Emtahan.RecordSource = "select * from emtahan where parvande like ('%" + Label38.Caption + "%') and nimpayan like ('%" + TaS + "%') and joze like ('%" + Combo4.Text + "%') and vazeyat like ('%" + "�����" + "%')"
Emtahan.Refresh
If Emtahan.Recordset.BOF = True Or Student.Recordset.EOF = True Then '�� �� ��� �� ��� ������ �ј� ��� ����
Option4.Value = True '����� 1 ���� �� ���
GoTo 5
'
End If

If Emtahan.Recordset.RecordCount < 2 Then 'ǐ� ����� �� 1 ��� �� ��� ������ �ј� ���� ����
Option5.Value = True '����� 2 ���� �� ���
GoTo 6
'
Else '��� �� 2 ��� �� ������ �ј� ���� ���
MsgBox "��� ���� ���� ��� �� 2 ���  �� ������ �ј� ���� ��� ���� ��� ���� ���� �����", vbCritical + vbOKOnly, "���"
Exit Sub

End If
End If


End If '���� ����� 10 �� 20


5
6

'���� ����� ��� 21 �� 30
If Val(Combo4.Text) >= 20 And Val(Combo4.Text) <= 29 Then

If Val(Label29.Caption) >= 10 Then '���� ����� �� 10 ������ �� ���
Option6.Value = True

Else '������




'���� ����� �� ��� ���� �� ����
If Option1.Value = True Then TaS = Option1.Caption
If Option2.Value = True Then TaS = Option2.Caption

Emtahan.Refresh
Emtahan.RecordSource = "select * from emtahan where parvande like ('%" + Label38.Caption + "%') and nimpayan like ('%" + TaS + "%') and joze like ('%" + Combo4.Text + "%') and vazeyat like ('%" + "�����" + "%')"
Emtahan.Refresh
If Emtahan.Recordset.BOF = True Or Student.Recordset.EOF = True Then
Option4.Value = True '����� � 1 ���� �����
GoTo 7

End If

If Emtahan.Recordset.RecordCount < 2 Then
Option5.Value = True
'������ 2 ���� �����
GoTo 8

Else
MsgBox "��� ���� ���� ��� �� 2 ���  �� ������ �ј� ���� ��� ���� ��� ���� ���� �����", vbCritical + vbOKOnly, "���"
Exit Sub

End If
End If


End If


7
8

'��� ����� ����� �� ����� ����� ���
'�� ������ �� �э �� ��� �� ����� �� ��� ��� ����� �� ��� ���� ��� �� ���


Dim TaSe, SERCH As String
If Option6.Value = True Then TaSe = "Q" '�����"
If Option4.Value = True Then TaSe = "T1" '������1
If Option5.Value = True Then TaSe = "T2" '������2


If Option1.Value = True Then SERCH = "P" & Label38.Caption & "J" & Combo4.Text & "NP5" & TaSe & Combo1.Text
If Option2.Value = True Then SERCH = "P" & Label38.Caption & "J" & Combo4.Text & "NP1" & TaSe & Combo1.Text
'�� ������ �� �� ����
Emtahan.Refresh
Emtahan.RecordSource = " select * from emtahan where kode like ('%" & SERCH & "%')"
Emtahan.Refresh
 If Emtahan.Recordset.BOF = True Or Emtahan.Recordset.EOF = True Then
 GoTo 17
 Else
 MsgBox " ��� ���� ��� ����� �� ��� ���� ���", vbCritical + vbOKOnly, "���"
 Exit Sub
 End If
 Exit Sub
17:
'���� ���� ��� �����

If MsgBox("���� ����� ����� ��� � ����� ����� ��� ����� �� ���� ��� ���� ����� �� ǘ��� ����� �� ��� ����", vbInformation + vbYesNo, "��� ����� ������") = vbYes Then
'�� ǘ��� ���� ����� ��� ���
Call Command2_Click
Else
Exit Sub
End If
'���� �Ԙ�� � �� ��� ��� ��� � ����� ��� ����� �� ����
'Command2.Enabled = True
Exit Sub

'��� ��� ��� ���� �� ����� � ���� ����� � ����� ��� �� 2 ���

1:
If Combo3.Text = "�� �����" Then SERCH = "P" & Label38.Caption & "RO" & Combo1.Text
If Combo3.Text = "���� �����" Then SERCH = "P" & Label38.Caption & "RA" & Combo1.Text
If Combo3.Text = "����� ��� 1" Then SERCH = "P" & Label38.Caption & "TJ1" & Combo1.Text
If Combo3.Text = "����� ��� 2" Then SERCH = "P" & Label38.Caption & "TJ2" & Combo1.Text
If Combo3.Text = "����� � ������" Then SERCH = "P" & Label38.Caption & "TR" & Combo1.Text

Emtahan.Refresh
Emtahan.RecordSource = " select * from emtahan where kode like ('%" & SERCH & "%')"
Emtahan.Refresh


If Val(Emtahan.Recordset.RecordCount) < 2 Then
 GoTo 25
 Else
 MsgBox "��� ���� ���� 2 ��� �� ��� ������ �ј� ���� � ����� �ј� �� ������ ���� �� �����", vbCritical + vbOKOnly, "���"
 Exit Sub
 End If
 Exit Sub
25:

If MsgBox("���� ����� ����� ��� � ����� ����� ��� ����� �� ���� ��� ���� ����� �� ǘ��� ����� �� ��� ����" & Chr$(10) & Label29.Caption & "���� �����", vbInformation + vbYesNo, "��� ����� ������") = vbYes Then
'�� ǘ��� ���� ����� ��� ���
Call Command2_Click
Else
Exit Sub
End If




End Sub

Private Sub Command6_Click()
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Combo11.Text = ""
Text1.Text = "�������"
Text18.Text = "��� ����"
Text6.Text = ""
Combo4.Text = ""
Combo5.Text = "������ ����"


Option1.Value = False
Option2.Value = False
Option4.Value = False
Option5.Value = False
Option6.Value = False


Label20.Caption = "-"
Label29.Caption = "-"
Label22.Caption = "-"


End Sub

Private Sub D_Click()

End Sub

Private Sub Command7_Click()
Option3.Value = True
Emtahan.Refresh
Emtahan.RecordSource = "select * from emtahan where parvande like ('%" + Student.Recordset.Fields("parvande") + "%')"
Emtahan.Refresh

End Sub

Private Sub Command8_Click()
If DataGrid1.AllowUpdate = False Then
Option3.Value = True

DataGrid1.AllowUpdate = True
MsgBox "���� ������� ���� ���� ����� ���� ����� �� �����", vbInformation, "����� �����"
Command8.Caption = "�����"
Else
DataGrid1.AllowUpdate = False
Command8.Caption = "����� �����"
End If

End Sub

Private Sub Command9_Click()

On Error GoTo 9898
GoTo 9999
9898:
MsgBox "���� ���� ���� �����", vbCritical + vbOKOnly, "���"
Exit Sub
9999:


mclass.Recordset.MovePrevious




End Sub

Private Sub DataGrid1_Click()
On Error Resume Next

Me.Emtahan.Recordset.Update

End Sub

Private Sub DataGrid2_DblClick()
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>


Dim A As String

On Error GoTo 7
GoTo 8
7:
MsgBox "���� ���� ���", vbCritical + vbOKOnly, "���"
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
 Me.Width = 12750
 
 Combo11.AddItem ("������")
Combo11.AddItem ("��ѐ���")
 
 
Combo3.AddItem ("�� �����")
Combo3.AddItem ("���� �����")
Combo3.AddItem ("����� ��� 1")
Combo3.AddItem ("����� ��� 2")
Combo3.AddItem ("��� 2 ����")
Combo3.AddItem ("��� 4 ����")
Combo3.AddItem ("��� 6 ����")
Combo3.AddItem ("����� �������")
Combo3.AddItem ("����")
Combo3.AddItem ("�������")
Combo3.AddItem ("����� � ������")

For I = 1 To 30
Combo4.AddItem (I)
Next I

Combo3.Text = Combo3.List(0)
Combo3.Text = Combo3.List(1)
Combo3.Text = Combo3.List(0)

Sb1.Panels(1).Text = user.OP.Text
Me.Sb1.Panels(3).Text = Taqvim.Tarikh.Caption
'Sb1.Panels(3).Text = Taqvim.Label1.Caption



'Combo6.AddItem ("���� ������")

'Combo6.AddItem ("���� ����� ����")
'Combo6.AddItem ("���� ���� ����� ����")
'
'Combo6.AddItem ("������ ����� ���")

Setting.Refresh
Setting.RecordSource = " select * from settingtable where xkodsetting like ('%" & "User-EmtahanF-Momtahen" & "%') "
Setting.Refresh
Combo6.Clear

 For I = 1 To Setting.Recordset.RecordCount
 Combo6.AddItem (Setting.Recordset.Fields("xtext"))
Setting.Recordset.MoveNext
Next I





For I = 1390 To 1408
Combo7.AddItem (I)
Next I

For I = 1 To 12
If I < 10 Then
Combo8.AddItem ("0" & I)
Else
Combo8.AddItem (I)
End If
Next I

For I = 1 To 31
If I < 10 Then
Combo9.AddItem ("0" & I)
Else
Combo9.AddItem (I)
End If

Next I

End Sub

Private Sub Form_Resize()
On Error Resume Next
'

DataGrid2.Width = EmtahanF.Width - 330
DataGrid3.Width = EmtahanF.Width - 330
DataGrid1.Width = EmtahanF.Width - 330


DataGrid2.Height = EmtahanF.Height - 6885

DataGrid3.Height = EmtahanF.Height - 6885
DataGrid1.Height = EmtahanF.Height - 6885

End Sub

Private Sub Form_Unload(Cancel As Integer)
Entekhab.Show

End Sub

Private Sub Label22_Change()
Label29.Caption = Val(Label29.Caption) - Val(Label22.Caption)

End Sub

Private Sub Label29_Change()
If Val(Label29.Caption) >= 19 And Val(Label29.Caption) <= 20 Then Combo5.Text = Combo5.List(0)
If Val(Label29.Caption) >= 18 And Val(Label29.Caption) < 19 Then Combo5.Text = Combo5.List(1)
If Val(Label29.Caption) >= 17 And Val(Label29.Caption) < 18 Then Combo5.Text = Combo5.List(2)
If Val(Label29.Caption) >= 15 And Val(Label29.Caption) < 17 Then Combo5.Text = Combo5.List(3)

If Val(Label29.Caption) >= 0 And Val(Label29.Caption) < 15 Then Combo5.Text = "������ ����"









End Sub

Private Sub Label38_Change()
''STU2CLASS.Refresh
'STU2CLASS.RecordSource = "select * from stu2class where parvande like ('%" + Label38.Caption + "%')"
''STU2CLASS.Refresh




'Combo1.Clear
'For I = 1 To STU2CLASS.Recordset.RecordCount


'Combo1.AddItem (STU2CLASS.Recordset.Fields("kodclass"))
'STU2CLASS.Recordset.MoveNext
'Next I
'Combo1.Text = Combo1.List(0)


On Error Resume Next


Combo1.Clear


Combo1.AddItem (Me.Student.Recordset.Fields("clas1"))
Combo1.AddItem (Me.Student.Recordset.Fields("clas2"))
Combo1.AddItem (Me.Student.Recordset.Fields("clas3"))
Combo1.AddItem (Me.Student.Recordset.Fields("clas4"))
Combo1.AddItem (Me.Student.Recordset.Fields("clas5"))

Combo1.Text = Combo1.List(0)





End Sub

Private Sub Label39_Click()
'Command2.Enabled = False
'
'�� ���� �� �� �� ��� �� ���� �� ����� �ј� ���� ����
If Me.lkodclass.Caption = "�����" Then
MsgBox "���� ���� ���� �� ����� ����", vbCritical + vbOKOnly, "���"
Exit Sub
End If
'��� ����� ������ ��� ���� ���� ����
If Combo7.Text = "" Or Combo8.Text = "" Or Combo9.Text = "" Then
MsgBox "����� ������ �� �� ���� ���� ���� ����", vbCritical + vbOKOnly, "���"
Exit Sub
End If
'��� ���� ��� ���� ���� ��� �� ���� ���� ����
If Text18.Text = "��� ����" Then
MsgBox "����� ���� �� �� ���� ����", vbCritical + vbOKOnly, "���"
Exit Sub
End If
'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
'����� ���ј �� ���� ��� ���
'��� �� �� �� ���� �� ���� �� ���� ��� ����� ���

'Hefz-Hefz-Hefz-Hefz-Hefz-Hefz-Hefzefz-Hefz-Hefz-Hefz-Hefz-Hefz-Hefz-Hefz-Hefz-Hefz-Hefz-Hefz-
'��� ��� ���� ��� ��� ��� � ����� �� ����� ���� �� ��� ��� ���� ���� ���

If Combo3.Text = "��� 2 ����" Or Combo3.Text = "�������" Or Combo3.Text = "��� 4 ����" Or Combo3.Text = "��� 6 ����" Or Combo3.Text = "����� �������" Or Combo3.Text = "����" Then

If Combo4.Text = "" Then '���� ����� ���
MsgBox "����� ��� �� ��� ����", vbCritical + vbOKOnly, "���"
Exit Sub
End If

'���� ���� ̎��� ������ ������
If Option1.Value = False And Option2.Value = False Then '���� ���� � ����� ��� ����
MsgBox "���� ��� � �� ����� ��� ���� ������ �� ���� ����", vbCritical + vbOKOnly, "���"
Exit Sub
End If
'���� ����� ���
'��������
'���� �����
If Combo11.Text = "������" Then
If Text8.Text = "" Or Val(Text8.Text) > 16 Or Text10.Text = "" Or Val(Text10.Text) > 2 Or Text12.Text = "" Or Val(Text12.Text) > 2 Then
MsgBox "����� �� ���� ���� ����", vbCritical + vbOKOnly, "���"
Exit Sub
End If
End If

'���� ����� ���
'���� ��ј���/
If Combo11.Text <> "������" Then
If Text8.Text = "" Or Val(Text8.Text) > 15 Or Text6.Text = "" Or Val(Text6.Text) > 2 Or Text10.Text = "" Or Val(Text10.Text) > 2 Or Text12.Text = "" Or Val(Text12.Text) > 1 Then
MsgBox "����� �� ���� ���� ����", vbCritical + vbOKOnly, "���"
Exit Sub
End If
End If


'Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-
'Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-
'Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-
'Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-
'Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-


'������� ����� �� ��� ������ � ����
Label29.Caption = (Val(Text8.Text) + Val(Text6.Text) + Val(Text10.Text) + Val(Text12.Text)) - Val(Label22.Caption) ' ��� ����




If Val(Combo4.Text) >= 1 And Val(Combo4.Text) <= 9 Or Val(Combo4.Text) = 30 Then 'ǐ� ���� ���� ��� ����� 1 � 10 ��� ����� �� ���� ���� ����� �����

        If Val(Label29.Caption) >= 17 Then '���� ����� �� 17 ���� �������� ���
        Option6.Value = True



Else '������
'���� ��� ����� 1 �10 ���� ������ ����� ���



        Dim TaS As String
        If Option1.Value = True Then TaS = Option1.Caption '���� ���
        If Option2.Value = True Then TaS = Option2.Caption '����� ���
            '�� ����� ����� �� ��� ����� ��� ������ �ј� ���� ���
        Emtahan.Refresh
        Emtahan.RecordSource = "select * from emtahan where parvande like ('%" + Label38.Caption + "%') and nimpayan like ('%" + TaS + "%') and joze like ('%" + Combo4.Text + "%') and vazeyat like ('%" + "�����" + "%')"
        Emtahan.Refresh
        '
            If Emtahan.Recordset.BOF = True Or Student.Recordset.EOF = True Then
            '���� ��� �� ��� ������� �ј� ���� ��� � ����� ��� ��� �� ���� ������1 ���� ��� ������
            Option4.Value = True 'sabt tajdidi 1
            GoTo 4
            '��� ��� ������ ���� �� ���� �� �� ��� ����� �� ����� 2 ���� ��� ���
            End If

    If Emtahan.Recordset.RecordCount <= 2 Then '���� ���� ���� �� ��� ������ �� ��� 1 ��� �ј� ���� ���
        Option5.Value = True '����� 2 ���� �� ���
        GoTo 2

    Else '���� ��� �� 2 ��� �� ��� ������ �ј� ���� ���
        MsgBox "��� ���� ���� ��� �� 2 ���  �� ������ �ј� ���� ��� ���� ��� ���� ���� �����", vbCritical + vbOKOnly, "���"
        Exit Sub

    End If
End If


End If '     ���� ��� 1 �� 10 �� ����


2: '����� 2
4: '����� 1



'ǐ� ���� ���� ��� ����� 11 �� 20 ��� ������ ������ ���� ����� ����
If Val(Combo4.Text) >= 10 And Val(Combo4.Text) <= 19 Then

If Val(Label29.Caption) >= 16 Then '����� �� 16 ���� �� ��� �����
Option6.Value = True



Else '������





If Option1.Value = True Then TaS = Option1.Caption
If Option2.Value = True Then TaS = Option2.Caption
'����� �� ��� �� ����� ��� ��� �� ��� ������ �ј� ���� ���
Emtahan.Refresh
Emtahan.RecordSource = "select * from emtahan where parvande like ('%" + Label38.Caption + "%') and nimpayan like ('%" + TaS + "%') and joze like ('%" + Combo4.Text + "%') and vazeyat like ('%" + "�����" + "%')"
Emtahan.Refresh
If Emtahan.Recordset.BOF = True Or Student.Recordset.EOF = True Then '�� �� ��� �� ��� ������ �ј� ��� ����
Option4.Value = True '����� 1 ���� �� ���
GoTo 5
'
End If

If Emtahan.Recordset.RecordCount < 2 Then 'ǐ� ����� �� 1 ��� �� ��� ������ �ј� ���� ����
Option5.Value = True '����� 2 ���� �� ���
GoTo 6
'
Else '��� �� 2 ��� �� ������ �ј� ���� ���
MsgBox "��� ���� ���� ��� �� 2 ���  �� ������ �ј� ���� ��� ���� ��� ���� ���� �����", vbCritical + vbOKOnly, "���"
Exit Sub

End If
End If


End If '���� ����� 10 �� 20


5
6

'���� ����� ��� 21 �� 30
If Val(Combo4.Text) >= 20 And Val(Combo4.Text) <= 29 Then

If Val(Label29.Caption) >= 10 Then '���� ����� �� 10 ������ �� ���
Option6.Value = True

Else '������




'���� ����� �� ��� ���� �� ����
If Option1.Value = True Then TaS = Option1.Caption
If Option2.Value = True Then TaS = Option2.Caption

Emtahan.Refresh
Emtahan.RecordSource = "select * from emtahan where parvande like ('%" + Label38.Caption + "%') and nimpayan like ('%" + TaS + "%') and joze like ('%" + Combo4.Text + "%') and vazeyat like ('%" + "�����" + "%')"
Emtahan.Refresh
If Emtahan.Recordset.BOF = True Or Student.Recordset.EOF = True Then
Option4.Value = True '����� � 1 ���� �����
GoTo 7

End If

If Emtahan.Recordset.RecordCount < 2 Then
Option5.Value = True
'������ 2 ���� �����
GoTo 8

Else
MsgBox "��� ���� ���� ��� �� 2 ���  �� ������ �ј� ���� ��� ���� ��� ���� ���� �����", vbCritical + vbOKOnly, "���"
Exit Sub

End If
End If


End If


7
8

'��� ����� ����� �� ����� ����� ���
'�� ������ �� �э �� ��� �� ����� �� ��� ��� ����� �� ��� ���� ��� �� ���


Dim TaSe, SERCH As String
If Option6.Value = True Then TaSe = "Q" '�����"
If Option4.Value = True Then TaSe = "T1" '������1
If Option5.Value = True Then TaSe = "T2" '������2


If Option1.Value = True Then SERCH = "P" & Label38.Caption & "J" & Combo4.Text & "NP5" & TaSe & Combo1.Text
If Option2.Value = True Then SERCH = "P" & Label38.Caption & "J" & Combo4.Text & "NP1" & TaSe & Combo1.Text
'�� ������ �� �� ����
Emtahan.Refresh
Emtahan.RecordSource = " select * from emtahan where kode like ('%" & SERCH & "%')"
Emtahan.Refresh
 If Emtahan.Recordset.BOF = True Or Emtahan.Recordset.EOF = True Then
 GoTo 17
 Else
 MsgBox " ��� ���� ��� ����� �� ��� ���� ���", vbCritical + vbOKOnly, "���"
 Exit Sub
 End If
 Exit Sub
17:
'���� ���� ��� �����

If MsgBox("���� ����� ����� ��� � ����� ����� ��� ����� �� ���� ��� ���� ����� �� ǘ��� ����� �� ��� ����", vbInformation + vbYesNo, "��� ����� ������") = vbYes Then
'�� ǘ��� ���� ����� ��� ���
Call Command2_Click
Else
Exit Sub
End If
'���� �Ԙ�� � �� ��� ��� ��� � ����� ��� ����� �� ����
'Command2.Enabled = True
Exit Sub

'��� ��� ��� ���� �� ����� � ���� ����� � ����� ��� �� 2 ���

1:
If Combo3.Text = "�� �����" Then SERCH = "P" & Label38.Caption & "RO" & Combo1.Text
If Combo3.Text = "���� �����" Then SERCH = "P" & Label38.Caption & "RA" & Combo1.Text
If Combo3.Text = "����� ��� 1" Then SERCH = "P" & Label38.Caption & "TJ1" & Combo1.Text
If Combo3.Text = "����� ��� 2" Then SERCH = "P" & Label38.Caption & "TJ2" & Combo1.Text
If Combo3.Text = "����� � ������" Then SERCH = "P" & Label38.Caption & "TR" & Combo1.Text

Emtahan.Refresh
Emtahan.RecordSource = " select * from emtahan where kode like ('%" & SERCH & "%')"
Emtahan.Refresh


If Val(Emtahan.Recordset.RecordCount) < 2 Then
 GoTo 25
 Else
 MsgBox "��� ���� ���� 2 ��� �� ��� ������ �ј� ���� � ����� �ј� �� ������ ���� �� �����", vbCritical + vbOKOnly, "���"
 Exit Sub
 End If
 Exit Sub
25:

If MsgBox("���� ����� ����� ��� � ����� ����� ��� ����� �� ���� ��� ���� ����� �� ǘ��� ����� �� ��� ����" & Chr$(10) & Label29.Caption & "���� �����", vbInformation + vbYesNo, "��� ����� ������") = vbYes Then
'�� ǘ��� ���� ����� ��� ���
Call Command2_Click
Else
Exit Sub
End If




'Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-
'Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-
'Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-
'Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-
'Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-
'Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-
'Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-Mohasebat Bakhsh Hefz-

'GoTo 555

Exit Sub

End If '���� ��� ���











 '��� �� ����� �� ��� ����� �� ��� ���� �� ��� ��� ȁ��


'End If

'��� ����� � ������


         
         
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%5 ��� �� �� �����
'Rokhani-Rokhani-Rokhani-Rokhani-Rokhani-Rokhani-Rokhani-Rokhani-Rokhani-Rokhani-Rokhani-Rokhani-Rokhani-Rokhani-
If Combo3.Text = "�� �����" Then


If Val(Text10.Text) > 3 Or Val(Text11.Text) > 20 Or Val(Text9.Text) > 17 Then
MsgBox "����� �� ���� ���� ����", vbCritical + vbOKOnly, "���"
Exit Sub
End If

'������ ����� ������ ������� ����� �� 17 + ����� �� 3 + ���� �� 20/����� �� 2
Label29.Caption = ((Val(Text10.Text) + Val(Text9.Text)) + Val(Text11.Text)) / 2
Label22.Caption = "�����" '����� ��� ������ �����

If Val(Text10.Text) + Val(Text9.Text) < 16 Or Val(Text11.Text) < 17 Then
Option4.Value = True
Else
Option6.Value = True
End If



If MsgBox("���� ����� ����� ��� � ����� ����� ��� ����� �� ���� ��� ���� ����� �� ǘ��� ����� �� ��� ����" & Chr$(10) & Label29.Caption & "���� �����", vbInformation + vbYesNo, "��� ����� ������") = vbYes Then
'�� ǘ��� ���� ����� ��� ���
Call Command2_Click
Else
Exit Sub
End If


'���� �� ��� ��� ����� 2 �ѐ�� ����� ����� ���
Exit Sub
End If




'Ravankhani-Ravankhani-Ravankhani-Ravankhani-Ravankhani-Ravankhani-Ravankhani-Ravankhani-Ravankhani-Ravankhani-Ravankhani-Ravankhani-
If Combo3.Text = "���� �����" Then

If Val(Text10.Text) > 3 Or Val(Text9.Text) > 17 Then
MsgBox "����� �� ���� ���� ����", vbCritical + vbOKOnly, "���"
Exit Sub
End If

'������ ����� ������ ������� ����� �� 17 + ����� �� 3 + ���� �� 20/����� �� 2
Label29.Caption = (Val(Text10.Text) + Val(Text9.Text))
Label22.Caption = "�����" '����� ��� ������ �����


'nomre qabooli
'mohasebe az 18 nomre ast
If Val(Label29.Caption) < 17 Then
Option4.Value = True

Else
Option6.Value = True


End If



If MsgBox("���� ����� ����� ��� � ����� ����� ��� ����� �� ���� ��� ���� ����� �� ǘ��� ����� �� ��� ����" & Chr$(10) & Label29.Caption & "���� �����", vbInformation + vbYesNo, "��� ����� ������") = vbYes Then
'�� ǘ��� ���� ����� ��� ���
Call Command2_Click
Else
Exit Sub
End If


'���� �� ��� ��� ����� 2 �ѐ�� ����� ����� ���
Exit Sub
End If






'tajvid-tajvid-tajvid-tajvid-tajvid-tajvid-tajvid-tajvid-tajvid-tajvid-tajvid-tajvid-tajvid-tajvid-tajvid-tajvid-tajvid-tajvid-

If Combo3.Text = "����� ��� 1" Or Combo3.Text = "����� ��� 2" Or Combo3.Text = "�����" Then

If Val(Text11.Text) > 20 Or Val(Text9.Text) > 20 Then
MsgBox "����� �� ���� ���� ����", vbCritical + vbOKOnly, "���"
Exit Sub
End If

'������ ����� ������ ������� ����� �� 17 + ����� �� 3 + ���� �� 20/����� �� 2
Label29.Caption = (Val(Text11.Text) + Val(Text9.Text)) / 2
Label22.Caption = "�����" '����� ��� ������ �����


'nomre qabooli
'mohasebe az 18 nomre ast
If Val(Label29.Caption) < 14 Then
Option4.Value = True

Else
Option6.Value = True


End If



If MsgBox("���� ����� ����� ��� � ����� ����� ��� ����� �� ���� ��� ���� ����� �� ǘ��� ����� �� ��� ����" & Chr$(10) & Label29.Caption & "���� �����", vbInformation + vbYesNo, "��� ����� ������") = vbYes Then
'�� ǘ��� ���� ����� ��� ���
Call Command2_Click
Else
Exit Sub
End If


'���� �� ��� ��� ����� 2 �ѐ�� ����� ����� ���
Exit Sub
End If





'Tarjome-Tarjome-Tarjome-Tarjome-Tarjome-Tarjome-Tarjome-Tarjome-Tarjome-Tarjome-Tarjome-Tarjome-Tarjome-Tarjome-Tarjome-

If Combo3.Text = "����� � ������" Then

If Val(Text11.Text) > 20 Then
MsgBox "����� �� ���� ���� ����", vbCritical + vbOKOnly, "���"
Exit Sub
End If

'������ ����� ������ ������� ����� �� 17 + ����� �� 3 + ���� �� 20/����� �� 2
Label29.Caption = Val(Text11.Text)
Label22.Caption = "�����" '����� ��� ������ �����


'nomre qabooli
'mohasebe az 18 nomre ast
If Val(Label29.Caption) < 15 Then
Option4.Value = True

Else
Option6.Value = True


End If



If MsgBox("���� ����� ����� ��� � ����� ����� ��� ����� �� ���� ��� ���� ����� �� ǘ��� ����� �� ��� ����" & Chr$(10) & Label29.Caption & "���� �����", vbInformation + vbYesNo, "��� ����� ������") = vbYes Then
'�� ǘ��� ���� ����� ��� ���
Call Command2_Click
Else
Exit Sub
End If


'���� �� ��� ��� ����� 2 �ѐ�� ����� ����� ���
Exit Sub
End If




End Sub

Private Sub Label62_Click()
mclass.Refresh
mclass.RecordSource = "seleCt * from mclass where kodclass like ('%" + Label62.Caption + "%')"
mclass.Refresh
End Sub

Private Sub Label63_Click()
Student.Refresh
Student.RecordSource = "select * from student where parvande like ('%" & Label63.Caption & "%')"
Student.Refresh
End Sub


Private Sub lkodclass_Click()
Student.Refresh
Student.RecordSource = "select * from student where clas1 like ('%" + lkodclass.Caption + "%') or clas2 like ('%" + lkodclass.Caption + "%') or clas3 like ('%" + lkodclass.Caption + "%') or clas4 like ('%" + lkodclass.Caption + "%') or clas5 like ('%" + lkodclass.Caption + "%')"
Student.Refresh

End Sub

Private Sub mnudell_Click()
Call Command3_Click

End Sub

Private Sub MNUEDITE_Click()
If Entekhab.SB.Panels(1).Text = "������ ����" Then GoTo 14082513

userprofiletable.Refresh
userprofiletable.RecordSource = "select * from userprofiletable where userid like ('" & user.useridtext.Text & "') and status like ('" & "on" & "') and commandname like ('" & "emtahan-edit" & "')"
userprofiletable.Refresh
If userprofiletable.Recordset.RecordCount <> 1 Then Exit Sub
14082513
If DataGrid1.AllowUpdate = False Then
MsgBox "����� ����� ������� ���� ���� ���� ����� �� �����" & Chr$(10) & "�� ����� �� ����� �� ����� ����� ������ ����� �� �� ������ ����", vbExclamation + vbOKOnly, "����� �����"
mnuedite.Checked = True

DataGrid1.AllowUpdate = True

Else

DataGrid1.AllowUpdate = False

mnuedite.Checked = False

End If


End Sub

Private Sub mnuhome_Click()
Entekhab.Show

End Sub

Private Sub mnuklarnamekelass_Click()
KarnameClas.Show

End Sub

Private Sub MNUSABTNOMARAT_Click()
Call Command5_Click

End Sub

Private Sub Option1_DblClick()
Option1.Value = False
End Sub

Private Sub Option2_DblClick()
Option2.Value = False
End Sub

Private Sub Option3_Click()
DataGrid1.Visible = True
DataGrid2.Visible = False
DataGrid3.Visible = False

End Sub

Private Sub Option4_DblClick()
Option4.Value = False

End Sub

Private Sub Option5_DblClick()
Option5.Value = False
End Sub


Private Sub Text13_Change()


Student.Refresh
Student.RecordSource = "select * from student where name like ('%" + Text13.Text + "%') or famil like ('%" + Text2.Text + "%')  or nf like ('%" + Text2.Text + "%') or parvande like ('%" + Text2.Text + "%')"
Student.Refresh






End Sub

Private Sub OptionHEFZ_Click()
DataGrid1.Visible = False
DataGrid2.Visible = True
DataGrid3.Visible = True
Call Command1_Click
Call Command1_Click

End Sub

Private Sub Sb1_PanelClick(ByVal Panel As ComctlLib.Panel)
Combo7.Text = Mid(Me.Sb1.Panels(3).Text, 1, 4)
Combo8.Text = Mid(Me.Sb1.Panels(3).Text, 6, 2)
Combo9.Text = Mid(Me.Sb1.Panels(3).Text, 9, 2)
End Sub

Private Sub TabStrip1_Click()
Select Case TabStrip1.SelectedItem.Caption

            Case "��� �����"
            OptionHEFZ.Value = True
            
            
            
            Case "����� ��� ���"
           Option3.Value = True
           Emtahan.Refresh
        Emtahan.RecordSource = "select * from emtahan where parvande like ('%" + "" + "%')"
        Emtahan.Refresh
           
            Case "����� ����� ��� ��� ���� ���� ����"
                Call Command7_Click
                
End Select


           
End Sub

Private Sub Text18_Change()
Label22.Caption = (Val(Text18.Text) * 0.1)

End Sub

Private Sub Text18_DblClick()
Text18.Text = "0"
End Sub

Private Sub Text2_Change()


Student.Refresh
Student.RecordSource = "select * from student where name like ('%" + Text2.Text + "%')or parvande like ('%" + Text2.Text + "%') or famil like ('%" + Text2.Text + "%') or nf like ('%" + Text2.Text + "%') "
Student.Refresh

Label1.Caption = Student.Recordset.RecordCount






End Sub

Private Sub Text2_DblClick()
Text2.Text = ""

End Sub

Private Sub Text5_Change()
mclass.Refresh
mclass.RecordSource = "select * from mclass where tarh like ('%" + Text5.Text + "%') or maqta like ('%" + Text5.Text + "%')or ostad like ('%" + Text5.Text + "%')or kodclass like ('%" + Text5.Text + "%')"
mclass.Refresh
End Sub
