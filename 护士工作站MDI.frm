VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl32.ocx"
Begin VB.MDIForm 护士工作站MDI 
   BackColor       =   &H8000000C&
   Caption         =   "护士工作站"
   ClientHeight    =   9840
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   14370
   Icon            =   "护士工作站MDI.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "护士工作站MDI.frx":1082
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   9345
      Width           =   14370
      _ExtentX        =   25347
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2858
            TextSave        =   "2016-06-20"
            Object.ToolTipText     =   "日期"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "16:02"
            Object.ToolTipText     =   "时间"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
            Object.ToolTipText     =   "当前用户"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.ToolTipText     =   "科室"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6615
            Object.ToolTipText     =   "属性"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   6615
            Text            =   "莎车县荒地镇卫生院"
            TextSave        =   "莎车县荒地镇卫生院"
            Object.ToolTipText     =   "医院名称"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Menu 常规管理 
      Caption         =   "常规管理(&Q)"
      Index           =   1
      Begin VB.Menu 分配 
         Caption         =   "病床分配"
         Shortcut        =   {F2}
      End
      Begin VB.Menu 医嘱 
         Caption         =   "医嘱执行"
         Shortcut        =   {F3}
      End
      Begin VB.Menu 床位动态 
         Caption         =   "床位动态"
         Shortcut        =   {F4}
      End
      Begin VB.Menu 收费 
         Caption         =   "预交款管理"
      End
      Begin VB.Menu 夜班住院管理 
         Caption         =   "夜班住院管理"
         Shortcut        =   {F6}
      End
      Begin VB.Menu 体温表 
         Caption         =   "体温表"
      End
      Begin VB.Menu 出院结算 
         Caption         =   "出院结算(&Y)"
      End
      Begin VB.Menu 查询 
         Caption         =   "病案查询"
      End
   End
   Begin VB.Menu 系统 
      Caption         =   "系统设置(&W)"
      Index           =   3
      Begin VB.Menu 病床 
         Caption         =   "病床管理"
      End
   End
   Begin VB.Menu 个性 
      Caption         =   "个性化设置(&E)"
      Begin VB.Menu 口令 
         Caption         =   "修改口令"
      End
      Begin VB.Menu 皮肤 
         Caption         =   "皮肤设置"
      End
   End
   Begin VB.Menu 其它 
      Caption         =   "其它(&X)"
      Begin VB.Menu 屏保 
         Caption         =   "屏幕保护"
      End
      Begin VB.Menu 退出 
         Caption         =   "退出系统"
      End
   End
End
Attribute VB_Name = "护士工作站MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub 出院结算_Click()
护士站.出院结算.Show
End Sub

Private Sub 床位动态_Click()
护士站.床位动态.Show
End Sub

Private Sub 分配_Click()
病床分配.Show
End Sub

Private Sub 口令_Click()
密码修改.Show
End Sub

Private Sub 体温表_Click()
护士站.体温表.Show
End Sub

Private Sub 夜班住院管理_Click()
夜班入院登记.Show
End Sub

Private Sub 医嘱_Click()
医嘱执行.Show
End Sub
