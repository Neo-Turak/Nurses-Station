VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form 出院结算 
   BackColor       =   &H00FFFFC0&
   Caption         =   "出院结算"
   ClientHeight    =   10380
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16605
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10380
   ScaleWidth      =   16605
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      Caption         =   "删除项"
      Height          =   495
      Left            =   3840
      TabIndex        =   10
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Timer Timer2 
      Left            =   5640
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Left            =   5640
      Top             =   1440
   End
   Begin VB.CommandButton Command5 
      Caption         =   "出院"
      Height          =   495
      Left            =   2640
      TabIndex        =   9
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "结算单打印"
      Height          =   495
      Left            =   600
      TabIndex        =   8
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "导入长期医嘱"
      Height          =   495
      Left            =   3360
      TabIndex        =   7
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "导入临时医嘱"
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   1320
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   495
      Left            =   1680
      Top             =   9000
      Visible         =   0   'False
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
      CommandType     =   2
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
      Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "结算清单"
      Caption         =   "Adodc4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid4 
      Bindings        =   "出院结算.frx":0000
      Height          =   7215
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   12726
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   20
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "使用清单"
         Caption         =   "使用清单"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "属性"
         Caption         =   "属性"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "规格"
         Caption         =   "规格"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "数量"
         Caption         =   "数量"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "单价"
         Caption         =   "单价"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "总价"
         Caption         =   "总价"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "部门结算人"
         Caption         =   "部门结算人"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "结算日期"
         Caption         =   "结算日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1379.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1305.071
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   10440
      Top             =   3240
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "临时医嘱"
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   450
      Left            =   10560
      Top             =   8880
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   794
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "长期医嘱"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "出院结算.frx":0015
      Height          =   2655
      Left            =   6480
      TabIndex        =   4
      Top             =   1200
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   4683
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   20
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "类别"
         Caption         =   "类别"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "所配科室"
         Caption         =   "所配科室"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "名称"
         Caption         =   "名称"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "规格"
         Caption         =   "规格"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "数量"
         Caption         =   "数量"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "单价"
         Caption         =   "单价"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "金额"
         Caption         =   "金额"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "科室"
         Caption         =   "科室"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "医生"
         Caption         =   "医生"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "医嘱日期"
         Caption         =   "医嘱日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "医嘱时间"
         Caption         =   "医嘱时间"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "执行时间"
         Caption         =   "执行时间"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "状态"
         Caption         =   "状态"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   824.882
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1425.26
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   750.047
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "出院结算.frx":002A
      Height          =   5775
      Left            =   6480
      TabIndex        =   3
      Top             =   3960
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   10186
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   20
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   16
      BeginProperty Column00 
         DataField       =   "医嘱名称"
         Caption         =   "医嘱名称"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "规格"
         Caption         =   "规格"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "执行频率"
         Caption         =   "执行频率"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "单价"
         Caption         =   "单价"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "一次数量"
         Caption         =   "一次数量"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "给药方式"
         Caption         =   "给药方式"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "总价"
         Caption         =   "总价"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "医嘱日期"
         Caption         =   "医嘱日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "医嘱时间"
         Caption         =   "医嘱时间"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "执行天数"
         Caption         =   "执行天数"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "停止日期"
         Caption         =   "停止日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "停止时间"
         Caption         =   "停止时间"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "科室"
         Caption         =   "科室"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "医生"
         Caption         =   "医生"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "护士"
         Caption         =   "护士"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "状态"
         Caption         =   "状态"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1409.953
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   3465.071
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   629.858
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   734.74
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查找"
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7680
      Top             =   720
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "床位动态"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   975
      Left            =   1560
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   1720
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   0   'False
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   20
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   14
      BeginProperty Column00 
         DataField       =   "患者编号"
         Caption         =   "患者编号"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "病床分区"
         Caption         =   "病床分区"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "床位号"
         Caption         =   "床位号"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "患者姓名"
         Caption         =   "患者姓名"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "性别"
         Caption         =   "性别"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "年龄"
         Caption         =   "年龄"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "住院号"
         Caption         =   "住院号"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "诊疗医生"
         Caption         =   "诊疗医生"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "诊断"
         Caption         =   "诊断"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "入院日期"
         Caption         =   "入院日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "天数"
         Caption         =   "天数"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "所属医生"
         Caption         =   "所属医生"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "出院日期"
         Caption         =   "出院日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "状态"
         Caption         =   "状态"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1769.953
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   884.976
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   629.858
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   840.189
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   240
      TabIndex        =   0
      Text            =   "1"
      Top             =   300
      Width           =   615
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Left            =   1560
      Top             =   120
      Width           =   14655
   End
End
Attribute VB_Name = "出院结算"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim SQL As String
SQL = "Provider=sqloledb.1;Data Source=NURA\SQLexpress;Persist Security Info=true;user id=sa;password=sa;initial catalog=ghgl"
Con.Open SQL
Con.CursorLocation = adUseClient
Mrc.Open "select * from 床位动态 where 病床分区='" & 护士工作站MDI.StatusBar1.Panels(4).Text & " ' and 床位号='" & Text1.Text & "'", Con, adOpenKeyset, adLockOptimistic
Set DataGrid1.DataSource = Mrc
Set Adodc1.Recordset = Mrc
Timer1.Interval = 100
End Sub


Private Sub Command2_Click()
Adodc3.Recordset.MoveFirst
For XH = 0 To Adodc3.Recordset.RecordCount
Adodc4.Recordset.AddNew
Adodc4.Recordset.Fields("住院号") = DataGrid1.Columns("住院号").CellValue(DataGrid1.Bookmark)
Adodc4.Recordset.Fields("床号") = DataGrid1.Columns("床位号").CellValue(DataGrid1.Bookmark)
Adodc4.Recordset.Fields("使用清单") = DataGrid3.Columns("名称").CellValue(DataGrid3.Bookmark)
Adodc4.Recordset.Fields("属性") = "临时医嘱"
Adodc4.Recordset.Fields("规格") = DataGrid3.Columns("规格").CellValue(DataGrid3.Bookmark)
Adodc4.Recordset.Fields("数量") = DataGrid3.Columns("数量").CellValue(DataGrid3.Bookmark)
Adodc4.Recordset.Fields("单价") = DataGrid3.Columns("单价").CellValue(DataGrid3.Bookmark)
Adodc4.Recordset.Fields("总价") = DataGrid3.Columns("金额").CellValue(DataGrid3.Bookmark)
Adodc4.Recordset.Fields("部门结算人") = 护士站.护士工作站MDI.StatusBar1.Panels(3).Text
Adodc4.Recordset.Fields("结算日期") = Date
Adodc3.Recordset.MoveNext
If Adodc3.Recordset.EOF = True Then
Adodc4.Recordset.Update
MsgBox "已完成导入临时医嘱清单！"
Exit For
End If
Next XH

End Sub

Private Sub Command3_Click()
On Error Resume Next
Adodc2.Recordset.MoveFirst

If Trim(Adodc2.Recordset.Fields("规格")) = "" Then
Adodc2.Recordset.MoveNext

End If
For XhH = 0 To Adodc2.Recordset.RecordCount
Adodc4.Recordset.AddNew
Adodc4.Recordset.Fields("住院号") = DataGrid1.Columns("住院号").CellValue(DataGrid1.Bookmark)
Adodc4.Recordset.Fields("床号") = DataGrid1.Columns("床位号").CellValue(DataGrid1.Bookmark)
Adodc4.Recordset.Fields("使用清单") = DataGrid2.Columns("医嘱名称").CellValue(DataGrid2.Bookmark)
Adodc4.Recordset.Fields("属性") = "长期医嘱"
Adodc4.Recordset.Fields("规格") = DataGrid2.Columns("规格").CellValue(DataGrid2.Bookmark)
If DataGrid2.Columns("执行频率") = "qd" Then
Adodc4.Recordset.Fields("数量") = DataGrid2.Columns("一次数量").CellValue(DataGrid2.Bookmark) * DataGrid2.Columns("执行天数").CellValue(DataGrid2.Bookmark)
End If

If DataGrid2.Columns("执行频率") = "bid" Then
Adodc4.Recordset.Fields("数量") = DataGrid2.Columns("一次数量").CellValue(DataGrid2.Bookmark) * DataGrid2.Columns("执行天数").CellValue(DataGrid2.Bookmark) * 2
End If

If DataGrid2.Columns("执行频率") = "tid" Then
Adodc4.Recordset.Fields("数量") = DataGrid2.Columns("一次数量").CellValue(DataGrid2.Bookmark) * DataGrid2.Columns("执行天数").CellValue(DataGrid2.Bookmark) * 3
End If
If Trim(DataGrid2.Columns("执行频率")) = "" Then
End If

Adodc4.Recordset.Fields("部门结算人") = 护士站.护士工作站MDI.StatusBar1.Panels(3).Text
Adodc4.Recordset.Fields("结算日期") = Date

Adodc2.Recordset.MoveNext

If Trim(Adodc2.Recordset.Fields("规格")) = "" Then
Adodc2.Recordset.MoveNext
End If

If Adodc2.Recordset.EOF = True Then
Adodc4.Recordset.Update
MsgBox "已完成导入长期医嘱清单！"
Exit For
End If
Next XhH
End Sub

Private Sub Command4_Click()
Adodc4.Recordset.MoveFirst
Printer.FillStyle = 0
Printer.ColorMode = 2
 Printer.ScaleMode = vbMillimeters
    Printer.Orientation = 1
    Printer.PaperSize = 13
    Printer.DrawStyle = 0
    Printer.CurrentX = 50
    Printer.CurrentY = 10
    Printer.FontSize = 16
    Printer.Font = 楷体
    Printer.FontBold = True
    Printer.Print "荒地镇卫生院 住院结算单"
   
    Printer.CurrentX = 10
    Printer.CurrentY = 23
    
    Printer.Font = 宋体
    Printer.FontSize = 12
     Printer.Print "患者姓名：" & DataGrid1.Columns("患者姓名").CellValue(DataGrid1.Bookmark) & Space(2) & "性别/年龄：" & DataGrid1.Columns("性别").CellValue(DataGrid1.Bookmark) & "/" & DataGrid1.Columns("年龄").CellValue(DataGrid1.Bookmark) & "岁" & "  " & "患者编号：" & DataGrid1.Columns("患者编号").CellValue(DataGrid1.Bookmark)
     Printer.CurrentX = 10
    Printer.CurrentY = 30
    Printer.Print "住院号：" & DataGrid1.Columns("住院号").CellValue(DataGrid1.Bookmark) & Space(2) & "入/出院日期：" & DataGrid1.Columns("入院日期").CellValue(DataGrid1.Bookmark) & "/" & DataGrid1.Columns("出院日期").CellValue(DataGrid1.Bookmark) & "     住院天数:" & DataGrid1.Columns("天数").CellValue(DataGrid1.Bookmark) & "   诊断：" & DataGrid1.Columns("诊断").CellValue(DataGrid1.Bookmark)
     Printer.Line (10, 35)-(170, 35)
     Printer.FontSize = 12
     Printer.FontBold = False
     
     For DY = 0 To Adodc4.Recordset.RecordCount
     Printer.CurrentX = 10
    Printer.CurrentY = 40 + DY * 5
    Printer.Print DataGrid4.Columns("使用清单").CellValue(DataGrid4.Bookmark)
    Printer.CurrentX = 50
    Printer.CurrentY = 40 + DY * 5
    Printer.Print DataGrid4.Columns("属性").CellValue(DataGrid4.Bookmark)
    Printer.CurrentX = 80
    Printer.CurrentY = 40 + DY * 5
    Printer.Print DataGrid4.Columns("规格").CellValue(DataGrid4.Bookmark)
    Printer.CurrentX = 100
    Printer.CurrentY = 40 + DY * 5
    Printer.Print DataGrid4.Columns("数量").CellValue(DataGrid4.Bookmark)
    Printer.CurrentX = 120
    Printer.CurrentY = 40 + DY * 5
    Printer.Print DataGrid4.Columns("总价").CellValue(DataGrid4.Bookmark)
    Adodc4.Recordset.MoveNext
    
    If Adodc4.Recordset.EOF = True Then
    Printer.Line (10, 45 + DY * 5)-(170, 45 + DY * 5)
    Printer.CurrentX = 10
    Printer.CurrentY = 50 + DY * 5
    Printer.FontBold = True
    Printer.Print "主治医生：" & DataGrid1.Columns("诊疗医生").CellValue(DataGrid1.Bookmark) & Space(3) & "结算护士：" & 护士站.护士工作站MDI.StatusBar1.Panels(3).Text
    Printer.CurrentX = 10
    Printer.CurrentY = 55 + DY * 5
    Printer.FontBold = True
    Printer.Print "住院部结算时间：" & Date & Space(5) & "合作医疗结算人："
    Printer.EndDoc
    Exit For
    End If
    
    Next DY
    
End Sub

Private Sub Command6_Click()
On Error Resume Next
Adodc4.Recordset.Delete
Adodc4.Recordset.Update
End Sub

Private Sub Timer1_Timer()
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim SQL As String
SQL = "Provider=sqloledb.1;Data Source=NURA\SQLexpress;Persist Security Info=true;user id=sa;password=sa;initial catalog=ghgl"
Con.Open SQL
Con.CursorLocation = adUseClient
Mrc.Open "select * from 长期医嘱 where 科室='" & 护士工作站MDI.StatusBar1.Panels(4).Text & " ' and 床号='" & Text1.Text & "'and 住院号='" & DataGrid1.Columns("住院号").CellValue(DataGrid1.Bookmark) & "'", Con, adOpenKeyset, adLockOptimistic
Set DataGrid2.DataSource = Mrc
Set Adodc2.Recordset = Mrc
Timer1.Interval = 0
Timer2.Interval = 100
End Sub

Private Sub Timer2_Timer()
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim SQL As String
SQL = "Provider=sqloledb.1;Data Source=NURA\SQLexpress;Persist Security Info=true;user id=sa;password=sa;initial catalog=ghgl"
Con.Open SQL
Con.CursorLocation = adUseClient
Mrc.Open "select * from 临时医嘱 where 科室='" & 护士工作站MDI.StatusBar1.Panels(4).Text & " ' and 住院号='" & DataGrid1.Columns("住院号").CellValue(DataGrid1.Bookmark) & "'", Con, adOpenKeyset, adLockOptimistic
Set DataGrid3.DataSource = Mrc
Set Adodc3.Recordset = Mrc
Timer2.Interval = 0
End Sub
