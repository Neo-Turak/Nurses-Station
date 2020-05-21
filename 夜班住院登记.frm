VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form 夜班入院登记 
   BackColor       =   &H00FFFFFF&
   Caption         =   "夜班入院登记"
   ClientHeight    =   9195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9675
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   9675
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   4080
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   3360
      Top             =   9360
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
      RecordSource    =   "住院单"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "住院登记"
      Height          =   495
      Left            =   7560
      TabIndex        =   8
      Top             =   3960
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "夜班住院登记.frx":0000
      Height          =   4455
      Left            =   0
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4560
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   7858
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1725.165
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   884.976
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   689.953
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Height          =   2295
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   9015
      Begin VB.CommandButton Command3 
         Caption         =   "添加档案"
         Height          =   495
         Left            =   7080
         TabIndex        =   22
         Top             =   200
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         ForeColor       =   &H000000FF&
         Height          =   360
         ItemData        =   "夜班住院登记.frx":0015
         Left            =   6960
         List            =   "夜班住院登记.frx":001F
         TabIndex        =   1
         Text            =   "结算方式"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   4680
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   2880
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   240
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   2160
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   1560
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text4 
         Height          =   360
         Left            =   5880
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2880
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         Height          =   360
         Left            =   240
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   960
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "查询"
         Height          =   495
         Left            =   5400
         TabIndex        =   7
         Top             =   200
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   3000
         TabIndex        =   0
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "患者编号    性别  年龄      民族         家庭住址         结算方式 "
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   1440
         Width           =   8535
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "       患者姓名                身份证号           医疗证号"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   8535
      End
      Begin VB.Label Label1 
         Caption         =   "请输入身份证号或医疗号："
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2895
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   1931
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "请输入收费金额："
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3720
      TabIndex        =   14
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "请输入住院号："
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   4080
      Width           =   1815
   End
End
Attribute VB_Name = "夜班入院登记"
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
If Mid(Text1.Text, 1, 2) = "07" Then
Mrc.Open "select * from 患者总表 where 合作医疗号 like'%" & Text1.Text & "%'", Con, adOpenKeyset, adLockOptimistic
End If
If Mid(Text1.Text, 1, 2) = "65" Then
Mrc.Open "select * from 患者总表 where 身份证号='" & Text1.Text & "'", Con, adOpenKeyset, adLockOptimistic
End If
Set DataGrid2.DataSource = Mrc
Set Text2.DataSource = Mrc
Text2.DataField = "患者姓名"

Set Text3.DataSource = Mrc
Text3.DataField = "身份证号"

Set Text4.DataSource = Mrc
Text4.DataField = "合作医疗号"

Set Text9.DataSource = Mrc
Text9.DataField = "患者编号"

Set Text7.DataSource = Mrc
Text7.DataField = "性别"

Set Text8.DataSource = Mrc
Text8.DataField = "年龄"

Set Text10.DataSource = Mrc
Text10.DataField = "民族"

'Set Text11.DataSource = Mrc
'Text11.DataField = "家庭住址"


DataGrid2.Refresh
End Sub

Private Sub Command2_Click()
If Combo1.Text = "结算方式" Or Text5.Text = "" Or Text6.Text = "" Then
MsgBox " 请填写必要内容！", vbExclamation
Exit Sub
End If
On Error Resume Next

Adodc2.Recordset.AddNew
With Adodc2.Recordset
.Fields("患者编号") = Text9.Text
.Fields("姓名") = Text2.Text
.Fields("性别") = Text7.Text
.Fields("年龄") = Text8.Text
.Fields("住院部") = 护士工作站MDI.StatusBar1.Panels(4).Text
.Fields("住院号") = Text5.Text
.Fields("身份证号") = Text3.Text
.Fields("医疗证号") = Text4.Text
.Fields("地址") = Text11.Text
.Fields("入院日期") = Date
.Fields("交款日期") = Date
.Fields("交款金额") = Text6.Text
.Fields("收款人姓名") = 护士工作站MDI.StatusBar1.Panels(3).Text
.Fields("状态") = "待排床"
.Update
End With
End Sub

Private Sub Command3_Click()
MsgBox "档案管理功能在维修中！"
End Sub

Private Sub Form_Activate()
Me.Width = 9615
Me.Height = 9900
Text1.SetFocus

End Sub

Private Sub Form_Load()
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim SQL As String
SQL = "Provider=sqloledb.1;Data Source=NURA\SQLexpress;Persist Security Info=true;user id=sa;password=sa;initial catalog=ghgl"
Con.Open SQL
Con.CursorLocation = adUseClient
Mrc.Open "select * from 住院单 order by '住院号'", Con, adOpenKeyset, adLockOptimistic
Set DataGrid3.DataSource = Mrc
Set Adodc2.Recordset = Mrc
End Sub

Private Sub Text4_Change()
If Left(Text4.Text, 2) = "07" Then
Text11.Text = "荒地镇" & Mid(Text4.Text, 3, 2) & "村   组"
Else
End If
End Sub
