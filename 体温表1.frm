VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form 体温表 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "三测单"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17280
   DrawStyle       =   2  'Dot
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   17280
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Height          =   6855
      Index           =   1
      Left            =   13560
      TabIndex        =   166
      Top             =   0
      Width           =   3495
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   35
         Left            =   2880
         TabIndex        =   274
         Top             =   6360
         Width           =   495
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   34
         Left            =   2880
         TabIndex        =   273
         Top             =   6000
         Width           =   495
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   33
         Left            =   2880
         TabIndex        =   272
         Top             =   5640
         Width           =   495
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   32
         Left            =   2880
         TabIndex        =   271
         Top             =   5280
         Width           =   495
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   31
         Left            =   2880
         TabIndex        =   270
         Top             =   4920
         Width           =   495
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   30
         Left            =   2880
         TabIndex        =   269
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   29
         Left            =   2880
         TabIndex        =   268
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   28
         Left            =   2880
         TabIndex        =   267
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   27
         Left            =   2880
         TabIndex        =   266
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   26
         Left            =   2880
         TabIndex        =   265
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   25
         Left            =   2880
         TabIndex        =   264
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   24
         Left            =   2880
         TabIndex        =   263
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   23
         Left            =   2880
         TabIndex        =   262
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   22
         Left            =   2880
         TabIndex        =   261
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   21
         Left            =   2880
         TabIndex        =   260
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   20
         Left            =   2880
         TabIndex        =   259
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   19
         Left            =   2880
         TabIndex        =   258
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   18
         Left            =   2880
         TabIndex        =   257
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   17
         Left            =   1080
         TabIndex        =   256
         Top             =   6360
         Width           =   615
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   16
         Left            =   1080
         TabIndex        =   255
         Top             =   6000
         Width           =   615
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   15
         Left            =   1080
         TabIndex        =   254
         Top             =   5640
         Width           =   615
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   14
         Left            =   1080
         TabIndex        =   253
         Top             =   5280
         Width           =   615
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   13
         Left            =   1080
         TabIndex        =   252
         Top             =   4920
         Width           =   615
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   12
         Left            =   1080
         TabIndex        =   251
         Top             =   4560
         Width           =   615
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   11
         Left            =   1080
         TabIndex        =   250
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   10
         Left            =   1080
         TabIndex        =   249
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   9
         Left            =   1080
         TabIndex        =   248
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   8
         Left            =   1080
         TabIndex        =   247
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   7
         Left            =   1080
         TabIndex        =   246
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   6
         Left            =   1080
         TabIndex        =   245
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   5
         Left            =   1080
         TabIndex        =   244
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   4
         Left            =   1080
         TabIndex        =   243
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   3
         Left            =   1080
         TabIndex        =   242
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   2
         Left            =   1080
         TabIndex        =   241
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   240
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF80FF&
         Caption         =   "0"
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   239
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   35
         Left            =   2280
         TabIndex        =   238
         Top             =   6360
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   34
         Left            =   2280
         TabIndex        =   237
         Top             =   6000
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   33
         Left            =   2280
         TabIndex        =   236
         Top             =   5640
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   32
         Left            =   2280
         TabIndex        =   235
         Top             =   5280
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   31
         Left            =   2280
         TabIndex        =   234
         Top             =   4920
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   30
         Left            =   2280
         TabIndex        =   233
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   29
         Left            =   2280
         TabIndex        =   232
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   28
         Left            =   2280
         TabIndex        =   231
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   27
         Left            =   2280
         TabIndex        =   230
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   26
         Left            =   2280
         TabIndex        =   229
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   25
         Left            =   2280
         TabIndex        =   228
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   24
         Left            =   2280
         TabIndex        =   227
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   23
         Left            =   2280
         TabIndex        =   226
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   22
         Left            =   2280
         TabIndex        =   225
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   21
         Left            =   2280
         TabIndex        =   224
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   20
         Left            =   2280
         TabIndex        =   223
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   19
         Left            =   2280
         TabIndex        =   222
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   18
         Left            =   2280
         TabIndex        =   221
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   17
         Left            =   600
         TabIndex        =   220
         Top             =   6360
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   16
         Left            =   600
         TabIndex        =   219
         Top             =   6000
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   15
         Left            =   600
         TabIndex        =   218
         Top             =   5640
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   14
         Left            =   600
         TabIndex        =   217
         Top             =   5280
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   13
         Left            =   600
         TabIndex        =   216
         Top             =   4920
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   12
         Left            =   600
         TabIndex        =   215
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   11
         Left            =   600
         TabIndex        =   214
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   10
         Left            =   600
         TabIndex        =   213
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   9
         Left            =   600
         TabIndex        =   212
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   8
         Left            =   600
         TabIndex        =   211
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   7
         Left            =   600
         TabIndex        =   210
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   6
         Left            =   600
         TabIndex        =   209
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   5
         Left            =   600
         TabIndex        =   208
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   4
         Left            =   600
         TabIndex        =   207
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   3
         Left            =   600
         TabIndex        =   206
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   2
         Left            =   600
         TabIndex        =   205
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   204
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   203
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D6MB24"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   35
         Left            =   1800
         TabIndex        =   202
         Top             =   6360
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D6MB20"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   34
         Left            =   1800
         TabIndex        =   201
         Top             =   6000
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D6MB16"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   33
         Left            =   1800
         TabIndex        =   200
         Top             =   5640
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D6MB12"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   32
         Left            =   1800
         TabIndex        =   199
         Top             =   5280
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D6MB8"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   31
         Left            =   1800
         TabIndex        =   198
         Top             =   4920
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D6MB4"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   30
         Left            =   1800
         TabIndex        =   197
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D5MB24"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   29
         Left            =   1800
         TabIndex        =   196
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D5MB20"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   28
         Left            =   1800
         TabIndex        =   195
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D5MB16"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   27
         Left            =   1800
         TabIndex        =   194
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D5MB12"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   26
         Left            =   1800
         TabIndex        =   193
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D5MB8"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   25
         Left            =   1800
         TabIndex        =   192
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D5MB4"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   24
         Left            =   1800
         TabIndex        =   191
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D4MB24"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   23
         Left            =   1800
         TabIndex        =   190
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D4MB20"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   22
         Left            =   1800
         TabIndex        =   189
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D4MB16"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   21
         Left            =   1800
         TabIndex        =   188
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D4MB12"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   20
         Left            =   1800
         TabIndex        =   187
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D4MB8"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   19
         Left            =   1800
         TabIndex        =   186
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D4MB4"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   18
         Left            =   1800
         TabIndex        =   185
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D3MB24"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   17
         Left            =   120
         TabIndex        =   184
         Top             =   6360
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D3MB20"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   16
         Left            =   120
         TabIndex        =   183
         Top             =   6000
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D3MB16"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   15
         Left            =   120
         TabIndex        =   182
         Top             =   5640
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D3MB12"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   14
         Left            =   120
         TabIndex        =   181
         Top             =   5280
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D3MB8"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   13
         Left            =   120
         TabIndex        =   180
         Top             =   4920
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D3MB4"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   12
         Left            =   120
         TabIndex        =   179
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D2MB24"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   11
         Left            =   120
         TabIndex        =   178
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D2MB20"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   10
         Left            =   120
         TabIndex        =   177
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D2MB16"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   176
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D2MB12"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   175
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D2MB8"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   174
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D2MB4"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   173
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D1MB24"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   172
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D1MB20"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   171
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D1MB16"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   170
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D1MB12"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   169
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D1MB8"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   168
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label12 
         DataField       =   "D1MB4"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   167
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Height          =   6855
      Index           =   0
      Left            =   10200
      TabIndex        =   48
      Top             =   0
      Width           =   3375
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   35
         Left            =   2760
         TabIndex        =   165
         Top             =   6360
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   34
         Left            =   2760
         TabIndex        =   164
         Top             =   6000
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   33
         Left            =   2760
         TabIndex        =   163
         Top             =   5640
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   32
         Left            =   2760
         TabIndex        =   162
         Top             =   5280
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   31
         Left            =   2760
         TabIndex        =   161
         Top             =   4920
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   30
         Left            =   2760
         TabIndex        =   160
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   29
         Left            =   2760
         TabIndex        =   159
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   28
         Left            =   2760
         TabIndex        =   158
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   27
         Left            =   2760
         TabIndex        =   157
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   26
         Left            =   2760
         TabIndex        =   156
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   25
         Left            =   2760
         TabIndex        =   155
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   24
         Left            =   2760
         TabIndex        =   154
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   23
         Left            =   2760
         TabIndex        =   153
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   22
         Left            =   2760
         TabIndex        =   152
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   21
         Left            =   2760
         TabIndex        =   151
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   20
         Left            =   2760
         TabIndex        =   150
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   19
         Left            =   2760
         TabIndex        =   149
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   18
         Left            =   2760
         TabIndex        =   148
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   17
         Left            =   1200
         TabIndex        =   147
         Top             =   6360
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   16
         Left            =   1200
         TabIndex        =   146
         Top             =   6000
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   15
         Left            =   1200
         TabIndex        =   145
         Top             =   5640
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   14
         Left            =   1200
         TabIndex        =   144
         Top             =   5280
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   13
         Left            =   1200
         TabIndex        =   143
         Top             =   4920
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   12
         Left            =   1200
         TabIndex        =   142
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   11
         Left            =   1200
         TabIndex        =   141
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   10
         Left            =   1200
         TabIndex        =   140
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   9
         Left            =   1200
         TabIndex        =   139
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   8
         Left            =   1200
         TabIndex        =   138
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   7
         Left            =   1200
         TabIndex        =   137
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   6
         Left            =   1200
         TabIndex        =   136
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   5
         Left            =   1200
         TabIndex        =   135
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   4
         Left            =   1200
         TabIndex        =   134
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   3
         Left            =   1200
         TabIndex        =   133
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   2
         Left            =   1200
         TabIndex        =   132
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   131
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   130
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   35
         Left            =   2280
         TabIndex        =   129
         Top             =   6360
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   34
         Left            =   2280
         TabIndex        =   128
         Top             =   6000
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   33
         Left            =   2280
         TabIndex        =   127
         Top             =   5640
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   32
         Left            =   2280
         TabIndex        =   126
         Top             =   5280
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   31
         Left            =   2280
         TabIndex        =   125
         Top             =   4920
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   30
         Left            =   2280
         TabIndex        =   124
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   29
         Left            =   2280
         TabIndex        =   123
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   28
         Left            =   2280
         TabIndex        =   122
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   27
         Left            =   2280
         TabIndex        =   121
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   26
         Left            =   2280
         TabIndex        =   120
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   25
         Left            =   2280
         TabIndex        =   119
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   24
         Left            =   2280
         TabIndex        =   118
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   23
         Left            =   2280
         TabIndex        =   117
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   22
         Left            =   2280
         TabIndex        =   116
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   21
         Left            =   2280
         TabIndex        =   115
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   20
         Left            =   2280
         TabIndex        =   114
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   19
         Left            =   2280
         TabIndex        =   113
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   18
         Left            =   2280
         TabIndex        =   112
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   17
         Left            =   720
         TabIndex        =   111
         Top             =   6360
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   16
         Left            =   720
         TabIndex        =   110
         Top             =   6000
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   15
         Left            =   720
         TabIndex        =   109
         Top             =   5640
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   14
         Left            =   720
         TabIndex        =   108
         Top             =   5280
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   13
         Left            =   720
         TabIndex        =   107
         Top             =   4920
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   12
         Left            =   720
         TabIndex        =   106
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   11
         Left            =   720
         TabIndex        =   105
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   10
         Left            =   720
         TabIndex        =   104
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   9
         Left            =   720
         TabIndex        =   103
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   8
         Left            =   720
         TabIndex        =   102
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   7
         Left            =   720
         TabIndex        =   101
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   6
         Left            =   720
         TabIndex        =   100
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   5
         Left            =   720
         TabIndex        =   99
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   4
         Left            =   720
         TabIndex        =   98
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   3
         Left            =   720
         TabIndex        =   97
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   96
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   95
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H000040C0&
         Caption         =   "0"
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   94
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label8 
         DataField       =   "d6tw24"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   35
         Left            =   1800
         TabIndex        =   92
         Top             =   6360
         Width           =   495
      End
      Begin VB.Label Label8 
         DataField       =   "d6tw20"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   34
         Left            =   1800
         TabIndex        =   91
         Top             =   6000
         Width           =   495
      End
      Begin VB.Label Label8 
         DataField       =   "d6tw16"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   33
         Left            =   1800
         TabIndex        =   90
         Top             =   5640
         Width           =   495
      End
      Begin VB.Label Label8 
         DataField       =   "d6tw12"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   32
         Left            =   1800
         TabIndex        =   89
         Top             =   5280
         Width           =   495
      End
      Begin VB.Label Label8 
         DataField       =   "d6tw8"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   31
         Left            =   1800
         TabIndex        =   88
         Top             =   4920
         Width           =   495
      End
      Begin VB.Label Label8 
         DataField       =   "d6tw4"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   30
         Left            =   1800
         TabIndex        =   87
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label Label8 
         DataField       =   "d5tw24"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   29
         Left            =   1800
         TabIndex        =   86
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label Label8 
         DataField       =   "d5tw20"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   28
         Left            =   1800
         TabIndex        =   85
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label Label8 
         DataField       =   "d5tw16"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   27
         Left            =   1800
         TabIndex        =   84
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label Label8 
         DataField       =   "d5tw12"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   26
         Left            =   1800
         TabIndex        =   83
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label Label8 
         DataField       =   "d5tw8"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   25
         Left            =   1800
         TabIndex        =   82
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label8 
         DataField       =   "d5tw4"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   24
         Left            =   1800
         TabIndex        =   81
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label8 
         DataField       =   "d4tw24"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   23
         Left            =   1800
         TabIndex        =   80
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label8 
         DataField       =   "d2tw20"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   22
         Left            =   1800
         TabIndex        =   79
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label8 
         DataField       =   "d4tw16"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   21
         Left            =   1800
         TabIndex        =   78
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label8 
         DataField       =   "d4tw12"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   20
         Left            =   1800
         TabIndex        =   77
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label8 
         DataField       =   "d4tw8"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   19
         Left            =   1800
         TabIndex        =   76
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label8 
         DataField       =   "d4tw4"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   18
         Left            =   1800
         TabIndex        =   75
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label8 
         DataField       =   "d3tw24"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   17
         Left            =   240
         TabIndex        =   74
         Top             =   6360
         Width           =   500
      End
      Begin VB.Label Label8 
         DataField       =   "D3TW20"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   16
         Left            =   240
         TabIndex        =   73
         Top             =   6000
         Width           =   500
      End
      Begin VB.Label Label8 
         DataField       =   "D3TW16"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   15
         Left            =   240
         TabIndex        =   72
         Top             =   5640
         Width           =   500
      End
      Begin VB.Label Label8 
         DataField       =   "D3TW12"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   14
         Left            =   240
         TabIndex        =   71
         Top             =   5280
         Width           =   500
      End
      Begin VB.Label Label8 
         DataField       =   "D3TW8"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   13
         Left            =   240
         TabIndex        =   70
         Top             =   4920
         Width           =   500
      End
      Begin VB.Label Label8 
         DataField       =   "D3TW4"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   12
         Left            =   240
         TabIndex        =   69
         Top             =   4560
         Width           =   500
      End
      Begin VB.Label Label8 
         DataField       =   "d2tw24"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   11
         Left            =   240
         TabIndex        =   68
         Top             =   4200
         Width           =   500
      End
      Begin VB.Label Label8 
         DataField       =   "D2TW20"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   10
         Left            =   240
         TabIndex        =   67
         Top             =   3840
         Width           =   500
      End
      Begin VB.Label Label8 
         DataField       =   "D2TW16"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   9
         Left            =   240
         TabIndex        =   66
         Top             =   3480
         Width           =   500
      End
      Begin VB.Label Label8 
         DataField       =   "D2TW12"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   65
         Top             =   3120
         Width           =   500
      End
      Begin VB.Label Label8 
         DataField       =   "D2TW8"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   64
         Top             =   2760
         Width           =   500
      End
      Begin VB.Label Label8 
         DataField       =   "D2TW4"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   63
         Top             =   2400
         Width           =   500
      End
      Begin VB.Label Label8 
         DataField       =   "D1TW24"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   62
         Top             =   2040
         Width           =   500
      End
      Begin VB.Label Label8 
         DataField       =   "D1TW20"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   61
         Top             =   1680
         Width           =   500
      End
      Begin VB.Label Label8 
         DataField       =   "D1TW16"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   60
         Top             =   1320
         Width           =   500
      End
      Begin VB.Label Label8 
         DataField       =   "D1TW12"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   59
         Top             =   960
         Width           =   500
      End
      Begin VB.Label Label8 
         DataField       =   "D1TW8"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   58
         Top             =   600
         Width           =   500
      End
      Begin VB.Label Label8 
         DataField       =   "D1TW4"
         DataSource      =   "Adodc2"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   57
         Top             =   240
         Width           =   500
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   6240
      Top             =   7080
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "体温单"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   4920
      TabIndex        =   40
      Top             =   4560
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin VB.CommandButton Command2 
      Caption         =   "保  存"
      Height          =   495
      Left            =   5040
      TabIndex        =   9
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   6855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4815
      Begin VB.TextBox Text16 
         DataField       =   "入院日期"
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   1320
         Width           =   2175
      End
      Begin VB.CommandButton Command8 
         Caption         =   "记录单查询"
         Height          =   615
         Left            =   2520
         TabIndex        =   31
         Top             =   2280
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00D1815F&
         Caption         =   "基本档案："
         Enabled         =   0   'False
         Height          =   3495
         Left            =   120
         TabIndex        =   12
         Top             =   3120
         Width           =   4455
         Begin VB.TextBox Text9 
            Height          =   375
            Left            =   2880
            TabIndex        =   47
            Top             =   1320
            Width           =   1095
         End
         Begin VB.CommandButton Command7 
            Caption         =   "保存"
            Height          =   495
            Left            =   3000
            TabIndex        =   30
            Top             =   1920
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton Command6 
            Caption         =   "修改"
            Height          =   495
            Left            =   3000
            TabIndex        =   29
            Top             =   2640
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.ComboBox Combo3 
            Height          =   360
            ItemData        =   "体温表1.frx":0000
            Left            =   1080
            List            =   "体温表1.frx":000D
            TabIndex        =   26
            Top             =   2760
            Width           =   1695
         End
         Begin VB.TextBox Text14 
            Height          =   360
            Left            =   1080
            TabIndex        =   25
            Top             =   2280
            Width           =   1695
         End
         Begin VB.ComboBox Combo1 
            Height          =   360
            ItemData        =   "体温表1.frx":0023
            Left            =   1080
            List            =   "体温表1.frx":0048
            TabIndex        =   23
            Top             =   1800
            Width           =   1695
         End
         Begin VB.TextBox Text12 
            Height          =   375
            Left            =   720
            MaxLength       =   10
            TabIndex        =   22
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox Text11 
            Height          =   375
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   20
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox Text10 
            Height          =   375
            Left            =   720
            MaxLength       =   10
            TabIndex        =   18
            Top             =   800
            Width           =   1095
         End
         Begin VB.TextBox Text7 
            Height          =   375
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   16
            Top             =   300
            Width           =   1095
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   720
            MaxLength       =   10
            TabIndex        =   14
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "血压：        nmHg"
            Height          =   255
            Left            =   2280
            TabIndex        =   46
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "大便次数"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   28
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "皮试结果"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   27
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "皮试信息"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   24
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "尿量：        ml"
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   21
            Top             =   1320
            Width           =   2295
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "出量：         ml"
            Height          =   375
            Index           =   3
            Left            =   2280
            TabIndex        =   19
            Top             =   885
            Width           =   2055
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "入量：        ml"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "体重：         kg"
            Height          =   375
            Index           =   1
            Left            =   2280
            TabIndex        =   15
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "身高：        cm"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "新建档案"
         Height          =   615
         Left            =   360
         TabIndex        =   11
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "查询"
         Height          =   495
         Left            =   3000
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         DataField       =   "床位号"
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   8
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   7
         Top             =   300
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         DataField       =   "患者姓名"
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   800
         Width           =   3135
      End
      Begin VB.Shape Shape1 
         Height          =   2775
         Left            =   120
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "床   号："
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "住 院 号:"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "入院日期："
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "病人姓名："
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "体温表打印"
      Height          =   495
      Left            =   8280
      TabIndex        =   0
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H8000000D&
      Caption         =   "记录："
      Height          =   4455
      Left            =   4920
      TabIndex        =   32
      Top             =   0
      Width           =   5295
      Begin VB.TextBox Text13 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   1
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   292
         Top             =   2760
         Width           =   2055
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Index           =   5
         Left            =   4680
         TabIndex        =   291
         Text            =   "8"
         Top             =   2280
         Width           =   580
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Index           =   4
         Left            =   3888
         TabIndex        =   290
         Text            =   "8"
         Top             =   2280
         Width           =   580
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Index           =   3
         Left            =   3096
         TabIndex        =   289
         Text            =   "8"
         Top             =   2280
         Width           =   580
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Index           =   2
         Left            =   2304
         TabIndex        =   288
         Text            =   "8"
         Top             =   2280
         Width           =   580
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Index           =   1
         Left            =   1512
         TabIndex        =   287
         Text            =   "8"
         Top             =   2280
         Width           =   580
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Index           =   5
         Left            =   4680
         TabIndex        =   286
         Text            =   "6"
         Top             =   1800
         Width           =   580
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Index           =   4
         Left            =   3888
         TabIndex        =   285
         Text            =   "6"
         Top             =   1800
         Width           =   580
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Index           =   3
         Left            =   3096
         TabIndex        =   284
         Text            =   "6"
         Top             =   1800
         Width           =   580
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Index           =   2
         Left            =   2304
         TabIndex        =   283
         Text            =   "6"
         Top             =   1800
         Width           =   580
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Index           =   1
         Left            =   1512
         TabIndex        =   282
         Text            =   "6"
         Top             =   1800
         Width           =   580
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Index           =   5
         Left            =   4680
         TabIndex        =   281
         Text            =   "5"
         Top             =   1320
         Width           =   580
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Index           =   4
         Left            =   3888
         TabIndex        =   280
         Text            =   "5"
         Top             =   1320
         Width           =   580
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Index           =   3
         Left            =   3096
         TabIndex        =   279
         Text            =   "5"
         Top             =   1320
         Width           =   580
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Index           =   2
         Left            =   2304
         TabIndex        =   278
         Text            =   "5"
         Top             =   1320
         Width           =   580
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Index           =   1
         Left            =   1512
         TabIndex        =   277
         Text            =   "5"
         Top             =   1320
         Width           =   580
      End
      Begin VB.CommandButton Command3 
         Caption         =   "调整"
         Height          =   495
         Left            =   1920
         TabIndex        =   93
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox Text13 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   0
         Left            =   720
         MaxLength       =   10
         TabIndex        =   44
         Top             =   2760
         Width           =   2055
      End
      Begin VB.ComboBox Combo5 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   360
         Left            =   1080
         TabIndex        =   43
         Text            =   "Combo5"
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   39
         Text            =   "8"
         Top             =   2280
         Width           =   580
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   38
         Text            =   "6"
         Top             =   1800
         Width           =   580
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   37
         Text            =   "5"
         Top             =   1320
         Width           =   580
      End
      Begin VB.CommandButton Command9 
         Caption         =   "确定"
         Height          =   375
         Left            =   3240
         TabIndex        =   33
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label26 
         BackColor       =   &H0000FFFF&
         Caption         =   "日  期："
         Height          =   255
         Left            =   120
         TabIndex        =   293
         Top             =   540
         Width           =   855
      End
      Begin VB.Label Label25 
         BackColor       =   &H0000FFFF&
         Caption         =   "时间段："
         Height          =   255
         Left            =   120
         TabIndex        =   276
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label24 
         BackColor       =   &H000040C0&
         Caption         =   "  4     8     12     16    20     24"
         Height          =   255
         Left            =   720
         TabIndex        =   275
         Top             =   960
         Width           =   4575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "血压："
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   45
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   4560
         TabIndex        =   42
         Top             =   500
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "呼吸："
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   36
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "脉搏："
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   35
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "体温："
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   1320
         Width           =   855
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   5280
         Y1              =   1250
         Y2              =   1250
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1200
      Top             =   6960
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
      RecordSource    =   "床位动态"
      Caption         =   "Adodc1"
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
   Begin VB.Label Label23 
      Caption         =   "Label19"
      Height          =   375
      Left            =   2040
      TabIndex        =   56
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Label22 
      Caption         =   "Label18"
      Height          =   375
      Left            =   2040
      TabIndex        =   55
      Top             =   0
      Width           =   405
   End
   Begin VB.Label Label21 
      Caption         =   "Label17"
      Height          =   375
      Left            =   1320
      TabIndex        =   54
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Label20 
      Caption         =   "Label16"
      Height          =   375
      Left            =   1320
      TabIndex        =   53
      Top             =   0
      Width           =   405
   End
   Begin VB.Label Label19 
      Caption         =   "Label19"
      Height          =   375
      Left            =   720
      TabIndex        =   52
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Label18 
      Caption         =   "Label18"
      Height          =   375
      Left            =   720
      TabIndex        =   51
      Top             =   0
      Width           =   405
   End
   Begin VB.Label Label17 
      Caption         =   "Label17"
      Height          =   375
      Left            =   0
      TabIndex        =   50
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Label16 
      Caption         =   "Label16"
      Height          =   375
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Width           =   405
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   10200
      Picture         =   "体温表1.frx":007A
      Stretch         =   -1  'True
      Top             =   10920
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "体温表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command9_Click

If Text5.Text = "" Then
Text5.SetFocus
Else
End If

End If
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text9.SetFocus
End If
End Sub

Private Sub Combo5_Change()
Label7.Caption = DateDiff("d", Text16.Text, Combo5.Text) + 1
End Sub

Private Sub Combo5_LostFocus()
Label7.Caption = DateDiff("d", Text16.Text, Combo5.Text) + 1
End Sub

Private Sub Command1_Click()
On Error Resume Next
Printer.FillStyle = 0
Printer.ColorMode = 2
 Printer.ScaleMode = vbMillimeters
    Printer.Orientation = 1
    Printer.PaperSize = 13
    Printer.DrawStyle = 0
  Printer.PaintPicture Image1.Picture, 0, 0, 176, 250
  Printer.FontSize = 10
  
 ' For i = 0 To 250 Step 5
 '  Printer.CurrentX = 0
 ' Printer.CurrentY = i
  ' Printer.Print i
   
 '  Next i
   
 ' For y = 0 To 180 Step 5
 ' Printer.CurrentX = y
 ' Printer.CurrentY = 10
 ' Printer.Print y
  
 '  Next y
   
   Printer.CurrentX = 13
   Printer.CurrentY = 21
   Printer.Print Left(Text1.Text, 9)
   
   Printer.CurrentX = 65
   Printer.CurrentY = 21
   Printer.Print Text16.Text
   Printer.FontBold = True
   
   Printer.CurrentX = 150
   Printer.CurrentY = 14
   Printer.FontSize = 12
   Printer.FontUnderline = True
   Printer.Print Text2.Text
   
    Printer.CurrentX = 135
   Printer.CurrentY = 21
   Printer.Print Text3.Text
   Printer.FontUnderline = False
   Printer.FontBold = False
   
   Printer.FontSize = 10
   Printer.CurrentX = 20
   Printer.CurrentY = 27
   Printer.Print Text16.Text
   
   For w = 1 To 6
   Printer.CurrentX = 25 + w * 20
   Printer.CurrentY = 27
   Printer.Print Right(DateAdd("d", w, Text16.Text), 5)
   
   Printer.CurrentX = 25 + (w - 1) * 20
   Printer.CurrentY = 31
   Printer.Print w
   
   
   Printer.CurrentX = 32 + (w - 1) * 20
   Printer.CurrentY = 196
   Printer.Print Combo1.Text
   
   
   Printer.CurrentX = 32 + (w - 1) * 20
   Printer.CurrentY = 212
   Printer.Print Text12.Text
     
   Next w
   Printer.CurrentX = 31
   Printer.CurrentY = 226
   Printer.Print Text7.Text
   
   Printer.CurrentX = 30
   Printer.CurrentY = 232
   Printer.Print Text4.Text
   
   '**************************************基本常规信息结束************
   '体温模点化块
   
 For LBL = 0 To 35 Step 1
If Not Label8(LBL).Caption = "" Then   '存在数据时
 Printer.Circle (20 + LBL * 4, 155 - (Label8(LBL).Caption - 36) * 16), (1)
Else
End If
 '  x(lbl) = 20 + Val(lbl * 4)                                 '第一个坐标X
   '  y(lbl) = 155 - Val((Label8(lbl).Caption - 36) * 16)        '第一个坐标Y
   Next LBL
   
  For STU = 0 To 35 Step 1
  If Label10(STU + 1).Caption = "0" Then
  Exit For
  Else
 
    Printer.Line (Label10(STU).Caption, Label11(STU).Caption)-(Label10(STU + 1).Caption, Label11(STU + 1).Caption)
  End If
  Next STU
  '***************************************************************************
  '脉搏点化
  
  For lml = 0 To 35 Step 1
  If Not Label12(lml).Caption = "" Then
  Printer.DrawMode = 13
  Printer.DrawStyle = 0
  Printer.FillStyle = 1
  Printer.CurrentX = (20 + Val(lml * 4)) - 0.95
  Printer.CurrentY = (155 - Val((Label12(lml).Caption - 60) * 0.875)) - 2.2
  Printer.FontSize = 10
  Printer.Print "x"
  'Printer.Circle ((20 + Val(lml * 4)), (155 - Val((Label12(lml).Caption - 60) * 0.875))), (1)
  Else
  End If
  
  Next lml
  
  Printer.FontSize = 10
  
  For stmb = 0 To 35 Step 1
  If Label13(stmb + 1).Caption = "0" Then
  Exit For
  Else
  
    Printer.Line (Label13(stmb).Caption, Label14(stmb).Caption)-(Label13(stmb + 1).Caption, Label14(stmb + 1).Caption)
  End If
  Next stmb
  
  '*****************************************************************
  '呼吸模块
  For cc = 0 To 5 Step 1
  Printer.FontSize = 6
  Printer.DrawMode = 10
  Printer.CurrentX = (18 + Val(cc * 3.2))
  Printer.CurrentY = 190
  If Not Adodc2.Recordset.Fields("D1HX" & (cc + 1) * 4) = "" Then
  Printer.Print "|" & Adodc2.Recordset.Fields("D1HX" & (cc + 1) * 4) & "|"
  Else
  End If
  Next cc
  
  For QQ = 0 To 5 Step 1
  Printer.FontSize = 6
  Printer.DrawMode = 10
  Printer.CurrentX = (38 + Val(QQ * 3.2))
  Printer.CurrentY = 190
  If Not Adodc2.Recordset.Fields("D2HX" & (QQ + 1) * 4) = "" Then
  Printer.Print "|" & Adodc2.Recordset.Fields("D2HX" & (QQ + 1) * 4) & "|"
  Else
  End If
  Next QQ
  
  For WW = 0 To 5 Step 1
  Printer.FontSize = 6
  Printer.DrawMode = 10
  Printer.CurrentX = (58 + Val(WW * 3.2))
  Printer.CurrentY = 190
  If Not Adodc2.Recordset.Fields("D3HX" & (WW + 1) * 4) = "" Then
  Printer.Print "|" & Adodc2.Recordset.Fields("D3HX" & (WW + 1) * 4) & "|"
  Else
  End If
  Next WW
  
  For EE = 0 To 5 Step 1
  Printer.FontSize = 6
  Printer.DrawMode = 10
  Printer.CurrentX = (58 + Val(EE * 3.2))
  Printer.CurrentY = 190
  If Not Adodc2.Recordset.Fields("D4HX" & (EE + 1) * 4) = "" Then
  Printer.Print "|" & Adodc2.Recordset.Fields("D4HX" & (EE + 1) * 4) & "|"
  Else
  End If
  Next EE
  
  For RR = 0 To 5 Step 1
  Printer.FontSize = 6
  Printer.DrawMode = 10
  Printer.CurrentX = (78 + Val(RR * 3.2))
  Printer.CurrentY = 190
  If Not Adodc2.Recordset.Fields("D5HX" & (RR + 1) * 4) = "" Then
  Printer.Print "|" & Adodc2.Recordset.Fields("D5HX" & (RR + 1) * 4) & "|"
  Else
  End If
  Next RR
  
  For YY = 0 To 5 Step 1
  Printer.FontSize = 6
  Printer.DrawMode = 10
  Printer.CurrentX = (98 + Val(YY * 3.2))
  Printer.CurrentY = 190
  If Not Adodc2.Recordset.Fields("D6HX" & (YY + 1) * 4) = "" Then
  Printer.Print "|" & Adodc2.Recordset.Fields("D6HX" & (YY + 1) * 4) & "|"
  Else
  End If
  Next YY
  
   Printer.EndDoc
End Sub

Private Sub Command2_Click()
On Error Resume Next

For AA = 0 To 5
If Not Text5(AA).Text = "" Then

If Not (Val(Text5(AA).Text) < 42 And Val(Text5(AA).Text) > 34) Then
MsgBox "出现不在正常范围的体温记录或数据错误，请改正后保存！"
Exit Sub
End If

Else
End If

If Not Text5(AA).Text = "" Then

If Not (Text6(AA).Text <> "" And Val(Text6(0).Text) > 20 And Val(Text6(0).Text) < 180) Then
MsgBox "出现不在正常范围的脉搏记录或数据有空字符串，请改正后保存！"
Exit Sub
End If

Else
End If

Next AA
Adodc2.Recordset.Update
End Sub

Private Sub Command2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command2_Click
Combo2.SetFocus
End If
End Sub


Private Sub Command3_Click()
Dim s As Integer
s = 0
 For LBL = 0 To 35 Step 1
   If Label8(LBL).Caption <> "" Then    '存在数据时
s = s + 1
Label10(s - 1).Caption = (20 + Val(LBL * 4))
Label11(s - 1).Caption = (155 - Val((Label8(LBL).Caption - 36) * 16))
Else

End If
   '  x(lbl) = 20 + Val(lbl * 4)                                 '第一个坐标X
   '  y(lbl) = 155 - Val((Label8(lbl).Caption - 36) * 16)        '第一个坐标Y
   Next LBL
   
   Dim TT As Integer
   TT = 0
For MBD = 0 To 35 Step 1
If Label12(MBD).Caption <> "" Then
TT = TT + 1
Label13(TT - 1).Caption = (20 + Val(MBD * 4))
Label14(TT - 1).Caption = (155 - Val((Label12(MBD).Caption - 60) * 0.875))
Else

End If

Next MBD

End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim SQL As String
SQL = "Provider=sqloledb.1;Data Source=NURA\SQLexpress;Persist Security Info=true;user id=sa;password=sa;initial catalog=ghgl"
Con.Open SQL
Con.CursorLocation = adUseClient
Mrc.Open "select * from 床位动态 where 住院号='" & Text2.Text & "'", Con, adOpenKeyset, adLockOptimistic
Set Adodc1.Recordset = Mrc
Set Text1.DataSource = Mrc
Set Text16.DataSource = Mrc
Set Text3.DataSource = Mrc

End Sub

Private Sub Command7_Click()
Adodc2.Recordset.Fields("第一天") = Text16.Text
Adodc2.Recordset.Fields("第二天") = DateAdd("d", 1, Text16.Text)
Adodc2.Recordset.Fields("第三天") = DateAdd("d", 2, Text16.Text)
Adodc2.Recordset.Fields("第四天") = DateAdd("d", 3, Text16.Text)
Adodc2.Recordset.Fields("第五天") = DateAdd("d", 4, Text16.Text)
Adodc2.Recordset.Fields("第六天") = DateAdd("d", 5, Text16.Text)
Adodc2.Recordset.Update
End Sub

Private Sub Command8_Click()
Combo5.Clear
Combo5.Text = Text16.Text
For w = 0 To 5
Combo5.AddItem DateAdd("d", w, Text16.Text)
Next w

On Error Resume Next
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim SQL As String
SQL = "Provider=sqloledb.1;Data Source=NURA\SQLexpress;Persist Security Info=true;user id=sa;password=sa;initial catalog=ghgl"
Con.Open SQL
Con.CursorLocation = adUseClient
Mrc.Open "select * from 体温单 where 住院号='" & Text2.Text & "'", Con, adOpenKeyset, adLockOptimistic
Set Adodc2.Recordset = Mrc
Set DataGrid1.DataSource = Mrc
If Adodc2.Recordset.RecordCount = 0 Then
'Set Adodc2.RecordSource = "体温记录"
Adodc2.Recordset.AddNew
Set Text4.DataSource = Adodc2
    Text4.DataField = "身高"
    
Set Text7.DataSource = Adodc2
    Text7.DataField = "体重"
    
    Set Text10.DataSource = Adodc2
    Text10.DataField = "入量"
    
     Set Text11.DataSource = Adodc2
    Text11.DataField = "出量"
    
     Set Text12.DataSource = Adodc2
    Text12.DataField = "尿量"
    
    Set Text9.DataSource = Adodc2
    Text9.DataField = "血压"
    
     Set Combo1.DataSource = Adodc2
    Combo1.DataField = "大便次数"
    
     Set Text14.DataSource = Adodc2
    Text14.DataField = "皮试信息"
    
    Set Combo3.DataSource = Adodc2
    Combo3.DataField = "皮试结果"
    
    Adodc2.Recordset.Fields("住院号") = Text2.Text
    Adodc2.Recordset.Fields("病人姓名") = Text1.Text
    Adodc2.Recordset.Fields("入院日期") = Text16.Text
    Adodc2.Recordset.Fields("床号") = Text3.Text
    Frame4.Enabled = True
    Command7.Visible = True
Else
Set Text4.DataSource = Adodc2
    Text4.DataField = "身高"
    
Set Text7.DataSource = Adodc2
    Text7.DataField = "体重"
    
    Set Text10.DataSource = Adodc2
    Text10.DataField = "入量"
    
     Set Text11.DataSource = Adodc2
    Text11.DataField = "出量"
    
     Set Text12.DataSource = Adodc2
    Text12.DataField = "尿量"
    
     Set Text9.DataSource = Adodc2
    Text9.DataField = "血压"
    
     Set Combo1.DataSource = Adodc2
    Combo1.DataField = "大便次数"
    
     Set Text14.DataSource = Adodc2
    Text14.DataField = "皮试信息"
    
    Set Combo3.DataSource = Adodc2
    Combo3.DataField = "皮试结果"
    Frame4.Enabled = False
    Label7.Caption = DateDiff("d", Text16.Text, Combo5.Text) + 1
End If
End Sub
Private Sub Command9_Click()
    '体温模块---------------------------------------------
    Set Text5(0).DataSource = Adodc2
    Text5(0).DataField = "D" & Label7.Caption & "TW4"
    Set Text5(1).DataSource = Adodc2
    Text5(1).DataField = "D" & Label7.Caption & "TW8"
    Set Text5(2).DataSource = Adodc2
    Text5(2).DataField = "D" & Label7.Caption & "TW12"
    Set Text5(3).DataSource = Adodc2
    Text5(3).DataField = "D" & Label7.Caption & "TW16"
    Set Text5(4).DataSource = Adodc2
    Text5(4).DataField = "D" & Label7.Caption & "TW20"
     Set Text5(5).DataSource = Adodc2
    Text5(5).DataField = "D" & Label7.Caption & "TW24"
    '脉搏模块--------------------------------------------
    Set Text6(0).DataSource = Adodc2
    Text6(0).DataField = "D" & Label7.Caption & "MB4"
    Set Text6(1).DataSource = Adodc2
    Text6(1).DataField = "D" & Label7.Caption & "MB8"
    Set Text6(2).DataSource = Adodc2
    Text6(2).DataField = "D" & Label7.Caption & "MB12"
    Set Text6(3).DataSource = Adodc2
    Text6(3).DataField = "D" & Label7.Caption & "MB16"
    Set Text6(4).DataSource = Adodc2
    Text6(4).DataField = "D" & Label7.Caption & "MB20"
    Set Text6(5).DataSource = Adodc2
    Text6(5).DataField = "D" & Label7.Caption & "MB24"
    '呼吸模块-------------------------------------------
    Set Text8(0).DataSource = Adodc2
    Text8(0).DataField = "D" & Label7.Caption & "HX4"
    Set Text8(1).DataSource = Adodc2
    Text8(1).DataField = "D" & Label7.Caption & "HX8"
    Set Text8(2).DataSource = Adodc2
    Text8(2).DataField = "D" & Label7.Caption & "HX12"
    Set Text8(3).DataSource = Adodc2
    Text8(3).DataField = "D" & Label7.Caption & "HX16"
    Set Text8(4).DataSource = Adodc2
    Text8(4).DataField = "D" & Label7.Caption & "HX20"
    Set Text8(5).DataSource = Adodc2
    Text8(5).DataField = "D" & Label7.Caption & "HX24"
    '血压模块-----------------------------------------
    Set Text13(0).DataSource = Adodc2
    Text13(0).DataField = "D" & Label7.Caption & "血压1"
    Set Text13(1).DataSource = Adodc2
    Text13(1).DataField = "D" & Label7.Caption & "血压2"
End Sub


Private Sub Form_Load()
Me.Width = 4815
Me.Height = 3465

End Sub

Private Sub Text1_Change()
Command8.Visible = True
Me.Width = 4950
Me.Height = 7365
Command8.SetFocus
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command2.SetFocus
End If
End Sub

Private Sub Text16_Change()
Combo5.Text = Text16.Text
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command4_Click
End If
End Sub
Private Sub Text4_Change()
Me.Width = 10380
Me.Height = 7440

End Sub


