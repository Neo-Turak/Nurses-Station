VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl32.ocx"
Begin VB.MDIForm ��ʿ����վMDI 
   BackColor       =   &H8000000C&
   Caption         =   "��ʿ����վ"
   ClientHeight    =   9840
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   14370
   Icon            =   "��ʿ����վMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "��ʿ����վMDI.frx":1082
   StartUpPosition =   3  '����ȱʡ
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
            Object.ToolTipText     =   "����"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "16:02"
            Object.ToolTipText     =   "ʱ��"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
            Object.ToolTipText     =   "��ǰ�û�"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.ToolTipText     =   "����"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6615
            Object.ToolTipText     =   "����"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   6615
            Text            =   "ɯ���ػĵ�������Ժ"
            TextSave        =   "ɯ���ػĵ�������Ժ"
            Object.ToolTipText     =   "ҽԺ����"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Menu ������� 
      Caption         =   "�������(&Q)"
      Index           =   1
      Begin VB.Menu ���� 
         Caption         =   "��������"
         Shortcut        =   {F2}
      End
      Begin VB.Menu ҽ�� 
         Caption         =   "ҽ��ִ��"
         Shortcut        =   {F3}
      End
      Begin VB.Menu ��λ��̬ 
         Caption         =   "��λ��̬"
         Shortcut        =   {F4}
      End
      Begin VB.Menu �շ� 
         Caption         =   "Ԥ�������"
      End
      Begin VB.Menu ҹ��סԺ���� 
         Caption         =   "ҹ��סԺ����"
         Shortcut        =   {F6}
      End
      Begin VB.Menu ���±� 
         Caption         =   "���±�"
      End
      Begin VB.Menu ��Ժ���� 
         Caption         =   "��Ժ����(&Y)"
      End
      Begin VB.Menu ��ѯ 
         Caption         =   "������ѯ"
      End
   End
   Begin VB.Menu ϵͳ 
      Caption         =   "ϵͳ����(&W)"
      Index           =   3
      Begin VB.Menu ���� 
         Caption         =   "��������"
      End
   End
   Begin VB.Menu ���� 
      Caption         =   "���Ի�����(&E)"
      Begin VB.Menu ���� 
         Caption         =   "�޸Ŀ���"
      End
      Begin VB.Menu Ƥ�� 
         Caption         =   "Ƥ������"
      End
   End
   Begin VB.Menu ���� 
      Caption         =   "����(&X)"
      Begin VB.Menu ���� 
         Caption         =   "��Ļ����"
      End
      Begin VB.Menu �˳� 
         Caption         =   "�˳�ϵͳ"
      End
   End
End
Attribute VB_Name = "��ʿ����վMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ��Ժ����_Click()
��ʿվ.��Ժ����.Show
End Sub

Private Sub ��λ��̬_Click()
��ʿվ.��λ��̬.Show
End Sub

Private Sub ����_Click()
��������.Show
End Sub

Private Sub ����_Click()
�����޸�.Show
End Sub

Private Sub ���±�_Click()
��ʿվ.���±�.Show
End Sub

Private Sub ҹ��סԺ����_Click()
ҹ����Ժ�Ǽ�.Show
End Sub

Private Sub ҽ��_Click()
ҽ��ִ��.Show
End Sub
