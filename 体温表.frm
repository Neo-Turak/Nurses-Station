VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form ���±� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ⵥ"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11415
   DrawStyle       =   2  'Dot
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
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
      Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=nura\sqlexpress"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=nura\sqlexpress"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "���µ�"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "���±�.frx":0000
      Height          =   2415
      Left            =   4920
      TabIndex        =   54
      Top             =   4560
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4260
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      ColumnCount     =   68
      BeginProperty Column00 
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
         DataField       =   "Ѫѹ"
         Caption         =   "Ѫѹ"
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
      BeginProperty Column06 
         DataField       =   "Ƥ����Ϣ"
         Caption         =   "Ƥ����Ϣ"
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
         DataField       =   "Ƥ�Խ��"
         Caption         =   "Ƥ�Խ��"
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
         DataField       =   "��һ��"
         Caption         =   "��һ��"
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
         DataField       =   "D1����"
         Caption         =   "D1����"
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
         DataField       =   "D1����"
         Caption         =   "D1����"
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
         DataField       =   "D1����"
         Caption         =   "D1����"
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
         DataField       =   "D1ʱ��"
         Caption         =   "D1ʱ��"
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
         DataField       =   "d1���²�λ"
         Caption         =   "d1���²�λ"
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
         DataField       =   "d1����"
         Caption         =   "d1����"
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
         DataField       =   "d1������"
         Caption         =   "d1������"
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
         DataField       =   "d1�¼�"
         Caption         =   "d1�¼�"
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
      BeginProperty Column17 
         DataField       =   "d1�¼�ʱ��"
         Caption         =   "d1�¼�ʱ��"
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
      BeginProperty Column18 
         DataField       =   "�ڶ���"
         Caption         =   "�ڶ���"
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
      BeginProperty Column19 
         DataField       =   "D2����"
         Caption         =   "D2����"
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
      BeginProperty Column20 
         DataField       =   "D2����"
         Caption         =   "D2����"
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
      BeginProperty Column21 
         DataField       =   "D2����"
         Caption         =   "D2����"
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
      BeginProperty Column22 
         DataField       =   "D2ʱ��"
         Caption         =   "D2ʱ��"
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
      BeginProperty Column23 
         DataField       =   "d2���²�λ"
         Caption         =   "d2���²�λ"
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
      BeginProperty Column24 
         DataField       =   "d2����"
         Caption         =   "d2����"
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
      BeginProperty Column25 
         DataField       =   "d2������"
         Caption         =   "d2������"
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
      BeginProperty Column26 
         DataField       =   "d2�¼�"
         Caption         =   "d2�¼�"
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
      BeginProperty Column27 
         DataField       =   "d2�¼�ʱ��"
         Caption         =   "d2�¼�ʱ��"
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
      BeginProperty Column28 
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
      BeginProperty Column29 
         DataField       =   "D3����"
         Caption         =   "D3����"
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
      BeginProperty Column30 
         DataField       =   "D3����"
         Caption         =   "D3����"
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
      BeginProperty Column31 
         DataField       =   "D3����"
         Caption         =   "D3����"
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
      BeginProperty Column32 
         DataField       =   "D3ʱ��"
         Caption         =   "D3ʱ��"
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
      BeginProperty Column33 
         DataField       =   "d3���²�λ"
         Caption         =   "d3���²�λ"
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
      BeginProperty Column34 
         DataField       =   "d3����"
         Caption         =   "d3����"
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
      BeginProperty Column35 
         DataField       =   "d3������"
         Caption         =   "d3������"
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
      BeginProperty Column36 
         DataField       =   "d3�¼�"
         Caption         =   "d3�¼�"
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
      BeginProperty Column37 
         DataField       =   "d3�¼�ʱ��"
         Caption         =   "d3�¼�ʱ��"
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
      BeginProperty Column38 
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
      BeginProperty Column39 
         DataField       =   "D4����"
         Caption         =   "D4����"
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
      BeginProperty Column40 
         DataField       =   "D4����"
         Caption         =   "D4����"
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
      BeginProperty Column41 
         DataField       =   "D4����"
         Caption         =   "D4����"
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
      BeginProperty Column42 
         DataField       =   "D4ʱ��"
         Caption         =   "D4ʱ��"
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
      BeginProperty Column43 
         DataField       =   "d4���²�λ"
         Caption         =   "d4���²�λ"
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
      BeginProperty Column44 
         DataField       =   "d4����"
         Caption         =   "d4����"
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
      BeginProperty Column45 
         DataField       =   "d4������"
         Caption         =   "d4������"
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
      BeginProperty Column46 
         DataField       =   "d4�¼�"
         Caption         =   "d4�¼�"
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
      BeginProperty Column47 
         DataField       =   "d4�¼�ʱ��"
         Caption         =   "d4�¼�ʱ��"
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
      BeginProperty Column48 
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
      BeginProperty Column49 
         DataField       =   "D5����"
         Caption         =   "D5����"
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
      BeginProperty Column50 
         DataField       =   "D5����"
         Caption         =   "D5����"
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
      BeginProperty Column51 
         DataField       =   "D5����"
         Caption         =   "D5����"
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
      BeginProperty Column52 
         DataField       =   "D5ʱ��"
         Caption         =   "D5ʱ��"
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
      BeginProperty Column53 
         DataField       =   "d5���²�λ"
         Caption         =   "d5���²�λ"
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
      BeginProperty Column54 
         DataField       =   "d5����"
         Caption         =   "d5����"
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
      BeginProperty Column55 
         DataField       =   "d5������"
         Caption         =   "d5������"
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
      BeginProperty Column56 
         DataField       =   "d5�¼�"
         Caption         =   "d5�¼�"
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
      BeginProperty Column57 
         DataField       =   "d5�¼�ʱ��"
         Caption         =   "d5�¼�ʱ��"
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
      BeginProperty Column58 
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
      BeginProperty Column59 
         DataField       =   "D6����"
         Caption         =   "D6����"
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
      BeginProperty Column60 
         DataField       =   "D6����"
         Caption         =   "D6����"
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
      BeginProperty Column61 
         DataField       =   "D6����"
         Caption         =   "D6����"
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
      BeginProperty Column62 
         DataField       =   "D6ʱ��"
         Caption         =   "D6ʱ��"
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
      BeginProperty Column63 
         DataField       =   "d6���²�λ"
         Caption         =   "d6���²�λ"
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
      BeginProperty Column64 
         DataField       =   "d6����"
         Caption         =   "d6����"
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
      BeginProperty Column65 
         DataField       =   "d6������"
         Caption         =   "d6������"
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
      BeginProperty Column66 
         DataField       =   "d6�¼�"
         Caption         =   "d6�¼�"
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
      BeginProperty Column67 
         DataField       =   "d6�¼�ʱ��"
         Caption         =   "d6�¼�ʱ��"
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
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   615.118
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   2534.74
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   2534.74
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   2534.74
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column24 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column25 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column26 
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column27 
            ColumnWidth     =   2534.74
         EndProperty
         BeginProperty Column28 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column29 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column30 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column31 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column32 
            ColumnWidth     =   2534.74
         EndProperty
         BeginProperty Column33 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column34 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column35 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column36 
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column37 
            ColumnWidth     =   2534.74
         EndProperty
         BeginProperty Column38 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column39 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column40 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column41 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column42 
            ColumnWidth     =   2534.74
         EndProperty
         BeginProperty Column43 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column44 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column45 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column46 
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column47 
            ColumnWidth     =   2534.74
         EndProperty
         BeginProperty Column48 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column49 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column50 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column51 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column52 
            ColumnWidth     =   2534.74
         EndProperty
         BeginProperty Column53 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column54 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column55 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column56 
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column57 
            ColumnWidth     =   2534.74
         EndProperty
         BeginProperty Column58 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column59 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column60 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column61 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column62 
            ColumnWidth     =   2534.74
         EndProperty
         BeginProperty Column63 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column64 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column65 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column66 
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column67 
            ColumnWidth     =   2534.74
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���"
      Height          =   495
      Left            =   8040
      TabIndex        =   10
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��  ��"
      Height          =   495
      Left            =   5040
      TabIndex        =   9
      Top             =   3000
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
         DataField       =   "��Ժ����"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1320
         TabIndex        =   55
         Top             =   1320
         Width           =   2175
      End
      Begin VB.CommandButton Command8 
         Caption         =   "��¼����ѯ"
         Height          =   615
         Left            =   2520
         TabIndex        =   34
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00D1815F&
         Caption         =   "����������"
         Enabled         =   0   'False
         Height          =   3495
         Left            =   120
         TabIndex        =   13
         Top             =   3120
         Width           =   4455
         Begin VB.CommandButton Command7 
            Caption         =   "����"
            Height          =   495
            Left            =   3000
            TabIndex        =   33
            Top             =   1920
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton Command6 
            Caption         =   "�޸�"
            Height          =   495
            Left            =   3000
            TabIndex        =   32
            Top             =   2640
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.ComboBox Combo3 
            Height          =   360
            ItemData        =   "���±�.frx":0015
            Left            =   1080
            List            =   "���±�.frx":0022
            TabIndex        =   29
            Top             =   2760
            Width           =   1695
         End
         Begin VB.TextBox Text14 
            Height          =   360
            Left            =   1080
            TabIndex        =   28
            Text            =   "Text14"
            Top             =   2280
            Width           =   1695
         End
         Begin VB.ComboBox Combo1 
            Height          =   360
            ItemData        =   "���±�.frx":0038
            Left            =   1080
            List            =   "���±�.frx":005A
            TabIndex        =   26
            Top             =   1800
            Width           =   1695
         End
         Begin VB.TextBox Text13 
            Height          =   375
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   25
            Text            =   "Text13"
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox Text12 
            Height          =   375
            Left            =   720
            MaxLength       =   10
            TabIndex        =   23
            Text            =   "Text12"
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox Text11 
            Height          =   375
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   21
            Text            =   "Text11"
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox Text10 
            Height          =   375
            Left            =   720
            MaxLength       =   10
            TabIndex        =   19
            Text            =   "Text10"
            Top             =   800
            Width           =   1095
         End
         Begin VB.TextBox Text7 
            Height          =   375
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   17
            Text            =   "Text7"
            Top             =   300
            Width           =   1095
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   720
            MaxLength       =   10
            TabIndex        =   15
            Text            =   "Text4"
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "������"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   31
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Ƥ�Խ��"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   30
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Ƥ����Ϣ"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   27
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Ѫѹ��       mmHg"
            Height          =   375
            Index           =   5
            Left            =   2280
            TabIndex        =   24
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "������        ml"
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   22
            Top             =   1320
            Width           =   2295
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "������         ml"
            Height          =   375
            Index           =   3
            Left            =   2280
            TabIndex        =   20
            Top             =   885
            Width           =   2055
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "������        ml"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   18
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "���أ�         kg"
            Height          =   375
            Index           =   1
            Left            =   2280
            TabIndex        =   16
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "��ߣ�        cm"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "�½�����"
         Height          =   615
         Left            =   360
         TabIndex        =   12
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "��ѯ"
         Height          =   495
         Left            =   3000
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         DataField       =   "��λ��"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   8
         Text            =   "Text3"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   300
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         DataField       =   "��������"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Text            =   "Text1"
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
         Caption         =   "��   �ţ�"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ס Ժ ��:"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ���ڣ�"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "����������"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���±��ӡ"
      Height          =   495
      Left            =   6480
      TabIndex        =   0
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H8000000D&
      Caption         =   "��¼��"
      Height          =   4455
      Left            =   4920
      TabIndex        =   35
      Top             =   0
      Width           =   5295
      Begin VB.TextBox Text15 
         DataSource      =   "Adodc2"
         Height          =   375
         Left            =   3840
         TabIndex        =   53
         Text            =   "Text15"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         DataSource      =   "Adodc2"
         Height          =   375
         Left            =   3840
         TabIndex        =   49
         Text            =   "Text9"
         Top             =   1680
         Width           =   735
      End
      Begin VB.ComboBox Combo4 
         DataSource      =   "Adodc2"
         Height          =   360
         ItemData        =   "���±�.frx":0086
         Left            =   3840
         List            =   "���±�.frx":0090
         TabIndex        =   47
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         DataSource      =   "Adodc2"
         Height          =   375
         Left            =   720
         TabIndex        =   44
         Text            =   "Text8"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         DataSource      =   "Adodc2"
         Height          =   375
         Left            =   720
         TabIndex        =   43
         Text            =   "Text6"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         DataSource      =   "Adodc2"
         Height          =   375
         Left            =   720
         TabIndex        =   42
         Text            =   "Text5"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         Caption         =   "ȷ��"
         Height          =   375
         Left            =   3600
         TabIndex        =   38
         Top             =   480
         Width           =   855
      End
      Begin VB.ComboBox Combo2 
         Height          =   360
         ItemData        =   "���±�.frx":00A0
         Left            =   1920
         List            =   "���±�.frx":00BA
         TabIndex        =   37
         Top             =   480
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   117309441
         CurrentDate     =   42526
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   4680
         TabIndex        =   56
         Top             =   520
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "���ʣ�"
         Height          =   375
         Index           =   9
         Left            =   2880
         TabIndex        =   52
         Top             =   2200
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "/��"
         Height          =   375
         Index           =   8
         Left            =   2040
         TabIndex        =   51
         Top             =   2200
         Width           =   495
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   375
         Index           =   7
         Left            =   4680
         TabIndex        =   50
         Top             =   1740
         Width           =   375
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "�����£�"
         Height          =   375
         Index           =   6
         Left            =   2760
         TabIndex        =   48
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "���²�λ��"
         Height          =   375
         Index           =   5
         Left            =   2760
         TabIndex        =   46
         Top             =   1245
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   375
         Index           =   4
         Left            =   2040
         TabIndex        =   45
         Top             =   1250
         Width           =   375
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   41
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   40
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "���£�"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   1200
         Width           =   855
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   5280
         Y1              =   960
         Y2              =   960
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1200
      Top             =   6960
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
      Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=nura\sqlexpress"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=nura\sqlexpress"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "��λ��̬"
      Caption         =   "Adodc1"
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
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   10200
      Picture         =   "���±�.frx":00E6
      Stretch         =   -1  'True
      Top             =   10920
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "���±�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Printer.FillStyle = 0
 Printer.ScaleMode = vbMillimeters
    Printer.Orientation = 1
    Printer.PaperSize = 13
    Printer.DrawStyle = 0
  Printer.PaintPicture Image1.Picture, 0, 0, 176, 250
  Printer.FontSize = 8
  
   For i = 0 To 250 Step 10
   Printer.CurrentX = 0
   Printer.CurrentY = i
   Printer.Print i
   Next i
   
   For y = 0 To 180 Step 10
   Printer.CurrentX = y
   Printer.CurrentY = 0
   Printer.Print y
   Next y
   
   Printer.CurrentX = 20
   Printer.CurrentY = 22
   Printer.Print Text1.Text
   
   Printer.CurrentX = 75
   Printer.CurrentY = 22
   Printer.Print DTPicker1.Value
   Printer.FontBold = True
   Printer.CurrentX = 155
   Printer.CurrentY = 22
   Printer.Print Text3.Text
   
   Printer.CurrentX = 155
   Printer.CurrentY = 15
   Printer.Print Text2.Text
   
   Printer.FontBold = False
   
    Printer.EndDoc
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.Update
End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim SQL As String
SQL = "Provider=sqloledb.1;Data Source=NURA\SQLEXPRESS;Persist Security Info=true;user id=sa;password=sa;initial catalog=ghgl"
Con.Open SQL
Con.CursorLocation = adUseClient
Mrc.Open "select * from ��λ��̬ where סԺ��='" & Text2.Text & "'", Con, adOpenKeyset, adLockOptimistic
Set Adodc1.Recordset = Mrc
End Sub

Private Sub Command7_Click()
Adodc2.Recordset.Update
End Sub

Private Sub Command8_Click()
On Error Resume Next
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim SQL As String
SQL = "Provider=sqloledb.1;Data Source=NURA\SQLEXPRESS;Persist Security Info=true;user id=sa;password=sa;initial catalog=ghgl"
Con.Open SQL
Con.CursorLocation = adUseClient
Mrc.Open "select * from ���µ� where סԺ��='" & Text2.Text & "'", Con, adOpenKeyset, adLockOptimistic
Set Adodc2.Recordset = Mrc
If Adodc2.Recordset.RecordCount = 0 Then
'Set Adodc2.RecordSource = "���¼�¼"
Adodc2.Recordset.AddNew
Set Text4.DataSource = Adodc2
    Text4.DataField = "���"
    
Set Text7.DataSource = Adodc2
    Text7.DataField = "����"
    
    Set Text10.DataSource = Adodc2
    Text10.DataField = "����"
    
     Set Text11.DataSource = Adodc2
    Text11.DataField = "����"
    
     Set Text12.DataSource = Adodc2
    Text12.DataField = "����"
    
     Set Text13.DataSource = Adodc2
    Text13.DataField = "Ѫѹ"
    
     Set Combo1.DataSource = Adodc2
    Combo1.DataField = "������"
    
     Set Text14.DataSource = Adodc2
    Text14.DataField = "Ƥ����Ϣ"
    
    Set Combo3.DataSource = Adodc2
    Combo3.DataField = "Ƥ�Խ��"
    
    Adodc2.Recordset.Fields("סԺ��") = Text2.Text
    Adodc2.Recordset.Fields("��������") = Text1.Text
    Adodc2.Recordset.Fields("��Ժ����") = Text16.Text
    Adodc2.Recordset.Fields("����") = Text3.Text
    Frame4.Enabled = True
    Command7.Visible = True
Else
Set Text4.DataSource = Adodc2
    Text4.DataField = "���"
    
Set Text7.DataSource = Adodc2
    Text7.DataField = "����"
    
    Set Text10.DataSource = Adodc2
    Text10.DataField = "����"
    
     Set Text11.DataSource = Adodc2
    Text11.DataField = "����"
    
     Set Text12.DataSource = Adodc2
    Text12.DataField = "����"
    
      Set Text13.DataSource = Adodc2
    Text13.DataField = "Ѫѹ"
    
     Set Combo1.DataSource = Adodc2
    Combo1.DataField = "������"
    
     Set Text14.DataSource = Adodc2
    Text14.DataField = "Ƥ����Ϣ"
    
    Set Combo3.DataSource = Adodc2
    Combo3.DataField = "Ƥ�Խ��"
    Frame4.Enabled = False
    Label7.Caption = DateDiff("d", Text16.Text, DTPicker2.Value) + 1
End If
End Sub

Private Sub Command9_Click()
    Set Text5.DataSource = Adodc2
    Text5.DataField = "D" & Label7.Caption & "TW" & Val(Combo2.Text)
    
    Set Text6.DataSource = Adodc2
    Text6.DataField = "D" & Label7.Caption & "MB" & Val(Combo2.Text)
    
   
    Set Text8.DataSource = Adodc2
    Text8.DataField = "D" & Label7.Caption & "HX" & Val(Combo2.Text)
    
    Set Text15.DataSource = Adodc2
    Text15.DataField = "D" & Label7.Caption & "XL" & Val(Combo2.Text)
End Sub

Private Sub DTPicker2_Change()
Label7.Caption = DateDiff("d", Text16.Text, DTPicker2.Value) + 1
End Sub

Private Sub Text16_Change()
DTPicker2.Value = Text16.Text
End Sub
