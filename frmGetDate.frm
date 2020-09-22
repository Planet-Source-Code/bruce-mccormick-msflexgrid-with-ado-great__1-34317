VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmGetDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select A Date"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdToday 
      Caption         =   "&Today"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2895
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4455
      _Version        =   524288
      _ExtentX        =   7858
      _ExtentY        =   5106
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2001
      Month           =   7
      Day             =   25
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmGetDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public datSelDate As Date

Private Sub Calendar1_Click()
    datSelDate = Calendar1.Value
End Sub

Private Sub cmdOk_Click()
    Me.Hide
End Sub

Private Sub cmdToday_Click()
    Calendar1.Value = Now
    Call Calendar1_Click
End Sub

Public Property Get SelDate() As Date
    SelDate = datSelDate
End Property

Public Property Let SelDate(ByVal vDate As Date)
    datSelDate = vDate
    Calendar1.Value = datSelDate
End Property

Public Sub Go()
    '  frmGetDate.SelDate = Now ' #7/3/1979#
    Me.Show vbModal
End Sub
