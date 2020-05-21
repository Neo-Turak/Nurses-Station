VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form 病床分配 
   Caption         =   "待入院患者单"
   ClientHeight    =   8745
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8745
   ScaleWidth      =   13560
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "病床分配.frx":0000
      Height          =   3495
      Left            =   240
      TabIndex        =   19
      Top             =   2640
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   6165
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   17
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
         DataField       =   "姓名"
         Caption         =   "姓名"
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
      BeginProperty Column03 
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
      BeginProperty Column04 
         DataField       =   "住院部"
         Caption         =   "住院部"
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
      BeginProperty Column06 
         DataField       =   "床号"
         Caption         =   "床号"
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
      BeginProperty Column08 
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
      BeginProperty Column09 
         DataField       =   "身份证号"
         Caption         =   "身份证号"
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
         DataField       =   "医疗证号"
         Caption         =   "医疗证号"
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
         DataField       =   "地址"
         Caption         =   "地址"
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
      BeginProperty Column13 
         DataField       =   "交款日期"
         Caption         =   "交款日期"
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
         DataField       =   "交款金额"
         Caption         =   "交款金额"
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
         DataField       =   "收款人姓名"
         Caption         =   "收款人姓名"
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
      BeginProperty Column16 
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
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   540.284
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   675.213
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   524.976
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   464.882
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "病床分配.frx":0015
      Height          =   2295
      Left            =   240
      TabIndex        =   17
      Top             =   6240
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   4048
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   10
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
      BeginProperty Column05 
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
      BeginProperty Column06 
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
      BeginProperty Column07 
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
      BeginProperty Column08 
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
      BeginProperty Column09 
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   884.976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1769.953
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   585.071
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   495
      Left            =   8160
      Top             =   6480
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=nura\sqlexpress"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=nura\sqlexpress"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "床位动态"
      Caption         =   "Adodc4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   2160
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\sqlexpress"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\sqlexpress"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "患者总表"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "患者信息"
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   13215
      Begin MSAdodcLib.Adodc Adodc5 
         Height          =   375
         Left            =   8640
         Top             =   2520
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
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
         Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQlexpress"
         OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQlexpress"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "用户表"
         Caption         =   "Adodc5"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.TextBox Text5 
         DataField       =   "住院号"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   29
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         DataField       =   "诊断"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   27
         Text            =   "4"
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox Text3 
         DataField       =   "合作医疗号"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         MaxLength       =   18
         TabIndex        =   21
         Text            =   "1"
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Frame Frame2 
         Caption         =   "排床："
         Height          =   2055
         Left            =   5640
         TabIndex        =   12
         Top             =   120
         Width           =   7455
         Begin VB.CommandButton Command4 
            Caption         =   "病 人 入 院"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5280
            TabIndex        =   32
            Top             =   840
            Width           =   2055
         End
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   31
            Text            =   "饮食类别"
            Top             =   1440
            Width           =   2535
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            ItemData        =   "病床分配.frx":002A
            Left            =   120
            List            =   "病床分配.frx":003A
            TabIndex        =   30
            Text            =   "护理级别"
            Top             =   1440
            Width           =   2175
         End
         Begin MSAdodcLib.Adodc Adodc3 
            Height          =   330
            Left            =   240
            Top             =   2040
            Visible         =   0   'False
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   582
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
            Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=nura\sqlexpress"
            OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=nura\sqlexpress"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "床位动态"
            Caption         =   "Adodc3"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin VB.CommandButton Command3 
            Caption         =   "刷新"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5880
            TabIndex        =   16
            Top             =   1440
            Width           =   1095
         End
         Begin VB.CommandButton Command2 
            Caption         =   "分 配  床 位"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5280
            TabIndex        =   15
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1200
            MaxLength       =   3
            TabIndex        =   13
            Top             =   240
            Width           =   615
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   390
            Left            =   120
            TabIndex        =   24
            Top             =   840
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   688
            _Version        =   393216
            ListField       =   "用户名"
            Text            =   "请选择医生"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label10 
            BackColor       =   &H00C0C000&
            DataField       =   "所属医生"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   25
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label Label6 
            Caption         =   "床位号："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "查询"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4440
         TabIndex        =   11
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text1 
         DataField       =   "身份证号"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         MaxLength       =   18
         TabIndex        =   6
         Text            =   "1"
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label12 
         DataField       =   "患者编号"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   28
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label11 
         Caption         =   "诊断"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "医疗证号："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   20
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "5"
         DataField       =   "年龄"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "4"
         DataField       =   "性别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "年龄："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4440
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "性别："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4440
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "3"
         DataField       =   "患者编号"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "2"
         DataField       =   "姓名"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "身份证号："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "患者编号："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "患者姓名："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   8040
      Top             =   7680
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=nura\sqlexpress"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=nura\sqlexpress"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "住院单"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   375
      Left            =   3840
      TabIndex        =   23
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   375
      Left            =   2040
      TabIndex        =   22
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   375
      Left            =   360
      TabIndex        =   18
      Top             =   7800
      Width           =   1335
   End
End
Attribute VB_Name = "病床分配"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim SQL As String
SQL = "Provider=sqloledb.1;Data Source=NURA\SQLexpress;Persist Security Info=true;user id=sa;password=sa;initial catalog=ghgl"
Con.Open SQL
Con.CursorLocation = adUseClient
Mrc.Open "select * from 患者总表 where 合作医疗号='" & Label3.Caption & "'or 患者编号='" & Label3.Caption & "'", Con, adOpenKeyset, adLockOptimistic
Set Text1.DataSource = Mrc
Set Text3.DataSource = Mrc
End Sub

Private Sub Command2_Click()

Adodc1.Recordset.Fields("床号") = Text2.Text
Adodc1.Recordset.Fields("身份证号") = Text1.Text
Adodc1.Recordset.Fields("医疗证号") = Text3.Text
Adodc1.Recordset.Fields("诊疗医生") = DataCombo1.Text
Adodc1.Recordset.Fields("入院日期") = Date
Adodc1.Recordset.Fields("状态") = "待收费"
Adodc1.Recordset.Update
MsgBox " 已分配床位，提交到住院收费！", vbInformation, "成功"

End Sub

Private Sub Command3_Click()
Unload Me
Load Me
End Sub

Private Sub Command4_Click()
Adodc4.Recordset.Fields("患者编号") = Label3.Caption
Adodc4.Recordset.Fields("病床分区") = 住院医生工作站.医生工作站MDI.StatusBar1.Panels(4).Text
Adodc4.Recordset.Fields("住院号") = Text5.Text
Adodc4.Recordset.Fields("床位号") = Text2.Text
Adodc4.Recordset.Fields("患者姓名") = Label2.Caption
Adodc4.Recordset.Fields("性别") = Label4.Caption
Adodc4.Recordset.Fields("年龄") = Label5.Caption
Adodc4.Recordset.Fields("护理级别") = Combo1.Text
Adodc4.Recordset.Fields("饮食类别") = Text6.Text
Adodc4.Recordset.Fields("诊疗医生") = DataCombo1.Text
Adodc4.Recordset.Fields("诊断") = Text4.Text
Adodc4.Recordset.Fields("入院日期") = Date
MsgBox "住院成功！", vbInformation
End Sub

Private Sub Form_Activate()
On Error Resume Next
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim SQL As String
SQL = "Provider=sqloledb.1;Data Source=NURA\SQLexpress;Persist Security Info=true;user id=sa;password=sa;initial catalog=ghgl"
Con.Open SQL
Con.CursorLocation = adUseClient
Mrc.Open "select * from 用户表 where 科室='" & 住院医生工作站.医生工作站MDI.StatusBar1.Panels(4).Text & "'and 职位 like'%医生%'", Con, adOpenKeyset, adLockOptimistic
Set Adodc3.Recordset = Mrc
Set DataCombo1.RowSource = Mrc

Dim co As ADODB.Connection
Dim mr As ADODB.Recordset
Set co = New ADODB.Connection
Set mr = New ADODB.Recordset
Dim sqll As String
sqll = "Provider=sqloledb.1;Data Source=NURA\SQLexpress;Persist Security Info=true;user id=sa;password=sa;initial catalog=ghgl"
co.Open sqll
co.CursorLocation = adUseClient
mr.Open "select * from 床位动态 where 所属医生='" & 住院医生工作站.医生工作站MDI.StatusBar1.Panels(3).Text & "'order by '床位号'", co, adOpenKeyset, adLockOptimistic
Set DataGrid2.DataSource = mr
End Sub

Private Sub Form_Load()
On Error Resume Next
Label7.Caption = 护士工作站MDI.StatusBar1.Panels(3).Text '姓名
Label8.Caption = 护士工作站MDI.StatusBar1.Panels(4).Text   '科室
Label9.Caption = 护士工作站MDI.StatusBar1.Panels(5).Text   '职位

Label2.Caption = ""
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Text1.Text = ""
Text3.Text = ""
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim SQL As String
SQL = "Provider=sqloledb.1;Data Source=NURA\SQLexpress;Persist Security Info=true;user id=sa;password=sa;initial catalog=ghgl"
Con.Open SQL
Con.CursorLocation = adUseClient
Mrc.Open "select * from 住院单 where 状态='待排床'", Con, adOpenKeyset, adLockOptimistic
Set DataGrid1.DataSource = Mrc
Set Adodc1.Recordset = Mrc
Set Label2.DataSource = Mrc
Set Label3.DataSource = Mrc
Set Label4.DataSource = Mrc
Set Label5.DataSource = Mrc

End Sub

Private Sub Form_Resize()
Me.Height = 9100
Me.Width = 12000
End Sub

Private Sub Label10_Change()
DataCombo1.Text = Label10.Caption
End Sub

Private Sub Text2_LostFocus()
On Error Resume Next
'Adodc4.Recordset.Find "床位号=" & Text2.Text
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim SQL As String
SQL = "Provider=sqloledb.1;Data Source=NURA\SQLexpress;Persist Security Info=true;user id=sa;password=sa;initial catalog=ghgl"
Con.Open SQL
Con.CursorLocation = adUseClient
Mrc.Open "select * from 床位动态 Where 所属医生='" & 住院医生工作站.医生工作站MDI.StatusBar1.Panels(3).Text & "' and 床位号='" & Text2.Text & "'", Con, adOpenKeyset, adLockOptimistic
Set Adodc4.Recordset = Mrc
Set DataGrid2.DataSource = Mrc
Set Label10.DataSource = Mrc
End Sub

Private Sub Text6_Change()

End Sub
