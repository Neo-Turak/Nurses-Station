VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form ҹ����Ժ�Ǽ� 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ҹ����Ժ�Ǽ�"
   ClientHeight    =   9195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9675
   BeginProperty Font 
      Name            =   "����"
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
      RecordSource    =   "סԺ��"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Caption         =   "סԺ�Ǽ�"
      Height          =   495
      Left            =   7560
      TabIndex        =   8
      Top             =   3960
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "ҹ��סԺ�Ǽ�.frx":0000
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
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   17
      BeginProperty Column00 
         DataField       =   "���߱��"
         Caption         =   "���߱��"
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
         DataField       =   "����"
         Caption         =   "����"
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
         DataField       =   "�Ա�"
         Caption         =   "�Ա�"
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
         DataField       =   "����"
         Caption         =   "����"
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
         DataField       =   "סԺ��"
         Caption         =   "סԺ��"
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
         DataField       =   "סԺ��"
         Caption         =   "סԺ��"
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
         DataField       =   "����"
         Caption         =   "����"
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
         DataField       =   "���"
         Caption         =   "���"
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
         DataField       =   "����ҽ��"
         Caption         =   "����ҽ��"
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
         DataField       =   "���֤��"
         Caption         =   "���֤��"
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
         DataField       =   "ҽ��֤��"
         Caption         =   "ҽ��֤��"
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
         DataField       =   "��ַ"
         Caption         =   "��ַ"
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
         DataField       =   "��Ժ����"
         Caption         =   "��Ժ����"
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
         DataField       =   "��������"
         Caption         =   "��������"
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
         DataField       =   "������"
         Caption         =   "������"
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
         DataField       =   "�տ�������"
         Caption         =   "�տ�������"
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
         DataField       =   "״̬"
         Caption         =   "״̬"
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
         Caption         =   "��ӵ���"
         Height          =   495
         Left            =   7080
         TabIndex        =   22
         Top             =   200
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         ForeColor       =   &H000000FF&
         Height          =   360
         ItemData        =   "ҹ��סԺ�Ǽ�.frx":0015
         Left            =   6960
         List            =   "ҹ��סԺ�Ǽ�.frx":001F
         TabIndex        =   1
         Text            =   "���㷽ʽ"
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
         Caption         =   "��ѯ"
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
         Caption         =   "���߱��    �Ա�  ����      ����         ��ͥסַ         ���㷽ʽ "
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   1440
         Width           =   8535
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "       ��������                ���֤��           ҽ��֤��"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   8535
      End
      Begin VB.Label Label1 
         Caption         =   "���������֤�Ż�ҽ�ƺţ�"
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
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Caption         =   "�������շѽ�"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3720
      TabIndex        =   14
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "������סԺ�ţ�"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   4080
      Width           =   1815
   End
End
Attribute VB_Name = "ҹ����Ժ�Ǽ�"
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
Mrc.Open "select * from �����ܱ� where ����ҽ�ƺ� like'%" & Text1.Text & "%'", Con, adOpenKeyset, adLockOptimistic
End If
If Mid(Text1.Text, 1, 2) = "65" Then
Mrc.Open "select * from �����ܱ� where ���֤��='" & Text1.Text & "'", Con, adOpenKeyset, adLockOptimistic
End If
Set DataGrid2.DataSource = Mrc
Set Text2.DataSource = Mrc
Text2.DataField = "��������"

Set Text3.DataSource = Mrc
Text3.DataField = "���֤��"

Set Text4.DataSource = Mrc
Text4.DataField = "����ҽ�ƺ�"

Set Text9.DataSource = Mrc
Text9.DataField = "���߱��"

Set Text7.DataSource = Mrc
Text7.DataField = "�Ա�"

Set Text8.DataSource = Mrc
Text8.DataField = "����"

Set Text10.DataSource = Mrc
Text10.DataField = "����"

'Set Text11.DataSource = Mrc
'Text11.DataField = "��ͥסַ"


DataGrid2.Refresh
End Sub

Private Sub Command2_Click()
If Combo1.Text = "���㷽ʽ" Or Text5.Text = "" Or Text6.Text = "" Then
MsgBox " ����д��Ҫ���ݣ�", vbExclamation
Exit Sub
End If
On Error Resume Next

Adodc2.Recordset.AddNew
With Adodc2.Recordset
.Fields("���߱��") = Text9.Text
.Fields("����") = Text2.Text
.Fields("�Ա�") = Text7.Text
.Fields("����") = Text8.Text
.Fields("סԺ��") = ��ʿ����վMDI.StatusBar1.Panels(4).Text
.Fields("סԺ��") = Text5.Text
.Fields("���֤��") = Text3.Text
.Fields("ҽ��֤��") = Text4.Text
.Fields("��ַ") = Text11.Text
.Fields("��Ժ����") = Date
.Fields("��������") = Date
.Fields("������") = Text6.Text
.Fields("�տ�������") = ��ʿ����վMDI.StatusBar1.Panels(3).Text
.Fields("״̬") = "���Ŵ�"
.Update
End With
End Sub

Private Sub Command3_Click()
MsgBox "������������ά���У�"
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
Mrc.Open "select * from סԺ�� order by 'סԺ��'", Con, adOpenKeyset, adLockOptimistic
Set DataGrid3.DataSource = Mrc
Set Adodc2.Recordset = Mrc
End Sub

Private Sub Text4_Change()
If Left(Text4.Text, 2) = "07" Then
Text11.Text = "�ĵ���" & Mid(Text4.Text, 3, 2) & "��   ��"
Else
End If
End Sub
