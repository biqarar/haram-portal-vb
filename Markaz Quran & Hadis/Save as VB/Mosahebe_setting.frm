VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form mosahebe_settingf 
   Caption         =   " ‰ŸÌ„«  „’«Õ»Â"
   ClientHeight    =   7350
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14175
   BeginProperty Font 
      Name            =   "B Homa"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Mosahebe_setting.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   14175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1095
      Left            =   2280
      TabIndex        =   3
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   -120
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   2895
      Begin MSAdodcLib.Adodc list_mosahebe 
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Reza\Desktop\DB_Haram_new\db.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Reza\Desktop\DB_Haram_new\db.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from list_mosahebe"
         Caption         =   "list_mosa"
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
      Begin MSAdodcLib.Adodc mosahebe_table 
         Height          =   615
         Left            =   0
         Top             =   720
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1085
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Reza\Desktop\DB_Haram_new\db.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Reza\Desktop\DB_Haram_new\db.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from mosahebe_table"
         Caption         =   "mosahebe_table"
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
   End
   Begin MSDataGridLib.DataGrid DataGridSTUDENT 
      Bindings        =   "Mosahebe_setting.frx":08CA
      Height          =   3255
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   5741
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   12648384
      DefColWidth     =   120
      HeadLines       =   1
      RowHeight       =   26
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
      Caption         =   " «—ÌŒ Ê “„«‰ „’«Õ»Â "
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
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
         DataField       =   "tarikh"
         Caption         =   "tarikh"
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
         DataField       =   "rooz"
         Caption         =   "rooz"
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
         DataField       =   "start"
         Caption         =   "start"
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
         DataField       =   "end"
         Caption         =   "end"
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
         DataField       =   "tedad"
         Caption         =   "tedad"
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
         DataField       =   "vazeyat"
         Caption         =   "vazeyat"
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
         DataField       =   "vorodi"
         Caption         =   "vorodi"
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
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Mosahebe_setting.frx":08E7
      Height          =   3255
      Left            =   3840
      TabIndex        =   2
      Top             =   3600
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   5741
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   12648384
      DefColWidth     =   120
      HeadLines       =   1
      RowHeight       =   26
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
      Caption         =   " ›òÌò ‘œÂ “„«‰ „’«Õ»Â"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
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
         DataField       =   "tarikh"
         Caption         =   "tarikh"
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
         DataField       =   "saat"
         Caption         =   "saat"
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
         DataField       =   "id_paziresh"
         Caption         =   "id_paziresh"
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
         DataField       =   "date_sabt"
         Caption         =   "date_sabt"
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
         DataField       =   "time"
         Caption         =   "time"
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
         DataField       =   "vazeyat"
         Caption         =   "vazeyat"
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
      EndProperty
   End
End
Attribute VB_Name = "mosahebe_settingf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim x_tarikh, x_rooz, x_start, x_end, x_zamane_har_fard, X_kol_zamane_mosahebe, x_tedad As String

x_tarikh = mosahebe_table.Recordset.Fields("tarikh")
x_rooz = mosahebe_table.Recordset.Fields("rooz")
x_start = mosahebe_table.Recordset.Fields("start")
x_end = mosahebe_table.Recordset.Fields("end")
x_tedad = mosahebe_table.Recordset.Fields("tedad")


X_kol_zamane_mosahebe = Val(x_end) - Val(x_start)
x_zamane_har_fard = (Val(X_kol_zamane_mosahebe) * 60) / Val(x_tedad)
' ta alan zamane har fard moshakhas shode ta badan bebinim chi mishe
list_mosahebe.Refresh
list_mosahebe.RecordSource = "select * from list_mosahebe where tarikh like ('%" + x_tarikh + "%') and saat like ('%" + x_start + "%')"
list_mosahebe.Refresh
If list_mosahebe.Recordset.BOF = False Or list_mosahebe.Recordset.EOF = False Then
MsgBox "een mored vojood darad"
Exit Sub
End If
list_mosahebe.Refresh
list_mosahebe.RecordSource = "select * from list_mosahebe" ' where tarikh like ('%" + x_tarikh + "%') and saat like ('%" + x_start + "%')"
list_mosahebe.Refresh
list_mosahebe.Recordset.Sort = "id"
list_mosahebe.Recordset.MovePrevious
list_mosahebe.Recordset.MoveLast

x_id = Val(list_mosahebe.Recordset.Fields("xid"))
'x_id = "910001"
list_mosahebe.Refresh
list_mosahebe.Recordset.AddNew
list_mosahebe.Recordset.Fields("xid") = x_id + 1
list_mosahebe.Recordset.Fields("vazeyat") = "Free" 'x_tarikh
list_mosahebe.Recordset.Fields("tarikh") = x_tarikh
list_mosahebe.Recordset.Fields("rooz") = x_rooz
list_mosahebe.Recordset.Fields("saat") = x_start
list_mosahebe.Recordset.Update
list_mosahebe.Refresh
T = 0
Y = x_id + 2

For I = 1 To x_tedad
Y = Y + 1

T = T + 1
If T * Val(x_zamane_har_fard) >= 60 Then
x_start = Val(x_start) + 1
T = 1
End If



zaman = Val(x_start) & ":" & (T * Val(x_zamane_har_fard))
If Val(x_start) >= Val(x_end) Then Exit Sub

'MsgBox zaman
list_mosahebe.Refresh
list_mosahebe.Recordset.AddNew
list_mosahebe.Recordset.Fields("xid") = Y

list_mosahebe.Recordset.Fields("vazeyat") = "Free" 'x_tarikh
list_mosahebe.Recordset.Fields("tarikh") = x_tarikh
list_mosahebe.Recordset.Fields("rooz") = x_rooz
list_mosahebe.Recordset.Fields("saat") = zaman
list_mosahebe.Recordset.Update
list_mosahebe.Refresh

Next I

list_mosahebe.Refresh
list_mosahebe.Recordset.AddNew
list_mosahebe.Recordset.Fields("xid") = Y + 1

list_mosahebe.Recordset.Fields("vazeyat") = "Free" 'x_tarikh
list_mosahebe.Recordset.Fields("tarikh") = x_tarikh
list_mosahebe.Recordset.Fields("rooz") = x_rooz
list_mosahebe.Recordset.Fields("saat") = x_end
list_mosahebe.Recordset.Update
list_mosahebe.Refresh
MsgBox "»Â « „«„ —”Ìœ"



End Sub

Private Sub Form_Unload(Cancel As Integer)
Entekhab.Show

End Sub
