VERSION 5.00
Begin VB.Form 床头卡打印 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "床头卡打印"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5985
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   5985
   Begin VB.CommandButton Command1 
      Caption         =   "打印"
      Height          =   495
      Left            =   1800
      TabIndex        =   20
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      Height          =   405
      Left            =   1560
      TabIndex        =   19
      Text            =   "Text7"
      Top             =   4200
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Height          =   405
      ItemData        =   "床头卡打印.frx":0000
      Left            =   1560
      List            =   "床头卡打印.frx":0010
      TabIndex        =   17
      Top             =   3480
      Width           =   2775
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   1080
      TabIndex        =   15
      Text            =   "Text6"
      Top             =   2760
      Width           =   3975
   End
   Begin VB.TextBox Text5 
      Height          =   405
      Left            =   3960
      TabIndex        =   13
      Text            =   "Text5"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   1080
      TabIndex        =   11
      Text            =   "Text4"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   1080
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   4200
      TabIndex        =   7
      Text            =   "2"
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   1080
      TabIndex        =   5
      Text            =   "1"
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3840
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "饭食类别："
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "护理级别："
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "诊断："
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "年龄："
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "性别："
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "姓名："
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "入院时间："
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
      Index           =   3
      Left            =   3000
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "住院号："
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
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "床号："
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
      Index           =   1
      Left            =   3120
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "可别："
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
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "床头卡打印"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
