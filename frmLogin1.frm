VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmLogin1 
   BorderStyle     =   0  'None
   Caption         =   "登录"
   ClientHeight    =   4620
   ClientLeft      =   2790
   ClientTop       =   3105
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin1.frx":0000
   ScaleHeight     =   2729.648
   ScaleMode       =   0  'User
   ScaleWidth      =   5281.57
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1800
      Top             =   4200
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
      CommandType     =   8
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
      RecordSource    =   ""
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
   Begin VB.CommandButton CmdCancel 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   3
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "住院护士工作站"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   1200
      Width           =   2895
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   375
      Left            =   900
      TabIndex        =   8
      Top             =   420
      Width           =   315
      VariousPropertyBits=   19
      Size            =   "556;661"
      FontName        =   "宋体"
      FontHeight      =   180
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      DataField       =   "科室"
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
      Left            =   1200
      TabIndex        =   7
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      DataField       =   "用户名"
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
      Left            =   2160
      TabIndex        =   6
      Top             =   1819
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      DataField       =   "密码"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      DataField       =   "职位"
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
      Left            =   3480
      TabIndex        =   4
      Top             =   3000
      Width           =   1335
   End
   Begin MSForms.TextBox TxtPassword 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   2280
      Width           =   3255
      VariousPropertyBits=   746604563
      Size            =   "5741;661"
      PasswordChar    =   42
      SpecialEffect   =   0
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextUserName 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1785
      Width           =   855
      VariousPropertyBits=   746604563
      Size            =   "1508;661"
      SpecialEffect   =   0
      FontName        =   "宋体"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmLogin1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SkinH_AttachEx Lib "D:\Users\NURA\vb 37款皮肤\SkinH_VB6.dll" (ByVal lpSkinFile As String, ByVal lpPasswd As String) As Long

Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'用于将CreateRoundRectRgn创建的圆角区域赋给窗体
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'用于创建一个圆角矩形，该矩形由X1，Y1-X2，Y2确定，并由X3，Y3确定的椭圆描述圆角弧度。
'参数 类型及说明：
'X1,Y1 Long，矩形左上角的X，Y坐标
'X2,Y2 Long，矩形右下角的X，Y坐标
'X3 Long，圆角椭圆的宽。其范围从0（没有圆角）到矩形宽（全圆）
'Y3 Long，圆角椭圆的高。其范围从0（没有圆角）到矩形高（全圆）
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'将CreateRoundRectRgn创建的区域删除，这是必要的，否则不必要的占用电脑内存
Dim outrgn As Long
'接下来声明一个全局变量,用来获得区域句柄
Option Explicit
Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    '设置全局变量为 false
    '不提示失败的登录
    LoginSucceeded = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
If Label1.Caption = "护士" Or Label1.Caption = "护士长" And Label4.Caption = "内科" Or Label4.Caption = "妇科" Then
If Trim(Label2.Caption) = Trim(TxtPassword.Text) Then
    '检查正确的密码
        LoginSucceeded = True
        Me.Hide
          护士工作站MDI.Show

      护士工作站MDI.StatusBar1.Panels(3) = Label3.Caption
        护士工作站MDI.StatusBar1.Panels(4) = Label4.Caption
       护士工作站MDI.StatusBar1.Panels(5) = Label1.Caption
    Else
        MsgBox "无效的密码或用户名，请重试!", , "登录"
        TextUserName.SetFocus
        SendKeys "{Home}+{End}"
        End If
        Else
        MsgBox "不符合的用户类型！"
        End If
End Sub

Private Sub CommandButton1_Click()
Dim x As Integer, y As Integer, Z As Integer
Z = (Me.Width - 4755) / 2
y = Me.Width / 2
x = Me.Height / 2   '高度
frmADODBLogon.Left = Me.Left + Z
frmADODBLogon.Top = VB.Screen.Height / 2 + x
frmADODBLogon.Show
End Sub

Private Sub Form_Activate() '窗体Activate()事件
Call rgnform(Me, 50, 50) '调用子过程
'SkinH_AttachEx "D:\Users\NURA\Desktop\电子病历\皮肤\晶莹洁白.she", "" '皮肤设置
End Sub
Private Sub Form_Load()
Dim x As Integer, y As Integer
x = Screen.Width / Screen.TwipsPerPixelX
y = Screen.Height / Screen.TwipsPerPixelY

End Sub

Private Sub Form_Unload(Cancel As Integer) '窗体Unload事件
DeleteObject outrgn '将圆角区域使用的所有系统资源释放
End Sub
Private Sub rgnform(ByVal frmbox As Form, ByVal fw As Long, ByVal fh As Long) '子过程，改变参数fw和fh的值可实现圆角
Dim w As Long, h As Long
w = frmbox.ScaleX(frmbox.Width, vbTwips, vbPixels)
h = frmbox.ScaleY(frmbox.Height, vbTwips, vbPixels)
outrgn = CreateRoundRectRgn(0, 0, w, h, fw, fh)
Call SetWindowRgn(frmbox.hwnd, outrgn, True)
End Sub
Private Sub TextUserName_LostFocus()
Dim Conn As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Conn = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim ConnectString As String
ConnectString = "Provider=SQLOLEDB.1;password=sa;Persist Security Info=False;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLexpress"
Conn.Open ConnectString
Conn.CursorLocation = adUseClient
Mrc.Open "select * from 用户表 where ID='" & TextUserName.Text & "'", Conn, adOpenKeyset, adLockOptimistic
    Set Label1.DataSource = Mrc
    Set Label2.DataSource = Mrc
    Set Label3.DataSource = Mrc
    Set Label2.DataSource = Mrc
    Set Label4.DataSource = Mrc
End Sub
