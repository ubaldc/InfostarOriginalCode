VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "reSize32.OCX"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "FPSPRu70.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frmpreds 
   Caption         =   "Predictions"
   ClientHeight    =   5850
   ClientLeft      =   360
   ClientTop       =   780
   ClientWidth     =   9765
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial Unicode MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmpreds.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   9765
   Begin FPUSpreadADO.fpSpread gridpreds 
      Height          =   3975
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   7935
      _Version        =   458752
      _ExtentX        =   13996
      _ExtentY        =   7011
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   50
      MaxRows         =   50
      SpreadDesigner  =   "frmpreds.frx":0442
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   840
      Top             =   4920
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5850
      FormDesignWidth =   9765
   End
   Begin HexUniControls.ctlUniButtonImageXP Command1 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   5040
      Width           =   1335
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmpreds.frx":0651
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmpreds.frx":0679
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmpreds.frx":0699
   End
End
Attribute VB_Name = "frmpreds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
frmpreds.Visible = False
End Sub











