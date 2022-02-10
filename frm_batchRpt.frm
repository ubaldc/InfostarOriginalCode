VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_batchRpt 
   Caption         =   "Batch Report"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11010
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   11010
   StartUpPosition =   2  'CenterScreen
   Begin FPUSpreadADO.fpSpread ss1 
      Height          =   6255
      Left            =   240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Width           =   10455
      _Version        =   458752
      _ExtentX        =   18441
      _ExtentY        =   11033
      _StockProps     =   64
      ColHeaderDisplay=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   -2147483633
      MaxCols         =   107
      MaxRows         =   5000
      OperationMode   =   1
      RetainSelBlock  =   0   'False
      SelectBlockOptions=   0
      SpreadDesigner  =   "frm_batchRpt.frx":0000
      UserResize      =   1
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   240
      Top             =   7080
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   8070
      FormDesignWidth =   11010
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_exit 
      Cancel          =   -1  'True
      Height          =   650
      Left            =   6480
      TabIndex        =   0
      Top             =   7080
      Width           =   2000
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_batchRpt.frx":0274
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_batchRpt.frx":029C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_batchRpt.frx":02BC
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_delete 
      Height          =   650
      Left            =   2400
      TabIndex        =   1
      Top             =   7080
      Width           =   2000
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_batchRpt.frx":02D8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_batchRpt.frx":030E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_batchRpt.frx":032E
   End
   Begin HexUniControls.ctlUniLabel lbl_batchName 
      Height          =   405
      Left            =   1440
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   714
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_batchRpt.frx":034A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_batchRpt.frx":037E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_batchRpt.frx":039E
   End
   Begin HexUniControls.ctlUniLabel lbl_batchNameFile 
      Height          =   405
      Left            =   4414
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   714
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_batchRpt.frx":03BA
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   0
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frm_batchRpt.frx":03DA
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_batchRpt.frx":03FA
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   240
      Top             =   7440
      _ExtentX        =   794
      _ExtentY        =   794
      PersistentTip   =   0   'False
      ActivateTips    =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TipBgColor      =   -2147483624
      TipFgColor      =   -2147483625
      Style           =   -1
      Transparency    =   0
   End
   Begin HexUniControls.ctlUniFormCaption uniTitle 
      Left            =   840
      Top             =   7080
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_batchRpt.frx":0416
   End
End
Attribute VB_Name = "frm_batchRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_restartLoop As Boolean

Private m_batchRptFile As String

Public Function create_batch_rpt_file(uniFile As clsUniFile, fileName As String) As Boolean
  Dim printStrg As String
  Dim nn As Integer
  
  If (uniFile.OpenFileWrite(fileName) = True) Then
    uniFile.WriteBOM fe_UTF16LE
      
    ' Build header
    ' Column label Date-Time (military format)
    printStrg = (Chr(34) & MLSupport.GSS("Headers", "dateTime", "Date-Time") & Chr(34))
    
    ' Column label System Serial Number
    printStrg = printStrg & "," & (Chr(34) & MLSupport.GSS("Headers", "serNum", "Serial No.") & Chr(34))
    
    ' Column label Product
    printStrg = printStrg & "," & (Chr(34) & MLSupport.GSS("Headers", "product", "Product") & Chr(34))

    ' Column label Sample ID
    printStrg = printStrg & "," & (Chr(34) & MLSupport.GSS("Headers", "sampleID", "Sample ID") & Chr(34))

    ' Column label Comment
    printStrg = printStrg & "," & (Chr(34) & MLSupport.GSS("Headers", "comment", "Comment") & Chr(34))

    ' Column label Status
    printStrg = printStrg & "," & (Chr(34) & MLSupport.GSS("Headers", "status", "Status") & Chr(34))
  
    ' Column label User Inputs
    For nn = 1 To MAX_MAN_INPUTS
      If (CSVUserInputs(nn) = True) Then
        printStrg = printStrg & "," & (Chr(34) & MLSupport.GSS("Headers", "input", "Input") & " " & nn & Chr(34))
      End If
    Next nn
  
    ' Column label Load Tower
    printStrg = printStrg & "," & (Chr(34) & MLSupport.GSS("Headers", "loadTower", "Load Tower") & Chr(34))

    ' Column label Cup Number
    printStrg = printStrg & "," & (Chr(34) & MLSupport.GSS("Headers", "cupNum", "Cup #") & Chr(34))

    ' Column label Unload Tower
    printStrg = printStrg & "," & (Chr(34) & MLSupport.GSS("Headers", "unloadTower", "Unload Tower") & Chr(34))

    ' Column label Cup Number
    printStrg = printStrg & "," & (Chr(34) & MLSupport.GSS("Headers", "cupNum", "Cup #") & Chr(34))
  
    ' Column label property info
    For nn = 1 To MAX_NUM_PROPS
      ' Property Name
      printStrg = printStrg & "," & (Chr(34) & MLSupport.GSS("Headers", "property", "Property") & " " & nn & Chr(34))
      
      ' Property Value
      printStrg = printStrg & "," & (Chr(34) & MLSupport.GSS("Headers", "value", "Value") & " " & nn & Chr(34))
     
      ' Property M-Distance
      printStrg = printStrg & "," & (Chr(34) & MLSupport.GSS("Headers", "mDist", "M-Dist") & " " & nn & Chr(34))
     
      ' Property S-Residual
      printStrg = printStrg & "," & (Chr(34) & MLSupport.GSS("Headers", "sResid", "S-Resid") & " " & nn & Chr(34))

      ' Property Outlier
      printStrg = printStrg & "," & (Chr(34) & MLSupport.GSS("Headers", "outlier", "Outlier") & " " & nn & Chr(34))
      
      ' Property Neighborhood Distance
      printStrg = printStrg & "," & (Chr(34) & MLSupport.GSS("Headers", "nD", "ND") & " " & nn & Chr(34))
     
      ' Property Intercept
      printStrg = printStrg & "," & (Chr(34) & MLSupport.GSS("Headers", "intercept", "Intercept") & " " & nn & Chr(34))
      
      ' Property Slope
      printStrg = printStrg & "," & (Chr(34) & MLSupport.GSS("Headers", "slope", "Slope") & " " & nn & Chr(34))
    Next nn
  End If

  uniFile.WriteUnicodeLine printStrg
  create_batch_rpt_file = True
End Function

Public Sub show_report(batchRptFile As String)
  Dim ret As Boolean

  m_batchRptFile = batchRptFile
  lbl_batchNameFile.Caption = CFile.st_FileNameNoExt(batchRptFile)
  
  On Error Resume Next
  ret = frm_batchRpt.ss1.LoadTextFile(m_batchRptFile, "", ",", Chr(13) + Chr(10), LoadTextFileColHeaders, "")
  frm_batchRpt.ss1.OperationMode = OperationModeRead
  
  If (ret = True) Then
    frm_batchRpt.ss1.Row = 0
    frm_batchRpt.ss1.Col = 0
    frm_batchRpt.ss1.Row2 = 0
    frm_batchRpt.ss1.Col2 = frm_batchRpt.ss1.MaxCols
    frm_batchRpt.ss1.BlockMode = True
    frm_batchRpt.ss1.Font.Bold = True
    frm_batchRpt.ss1.Font.Name = "Arial Unicode MS"
    frm_batchRpt.ss1.Font.Size = 10
    frm_batchRpt.ss1.BlockMode = False
    
    frm_batchRpt.ss1.Row = 1
    frm_batchRpt.ss1.Col = 0
    frm_batchRpt.ss1.Row2 = frm_batchRpt.ss1.MaxRows
    frm_batchRpt.ss1.Col2 = frm_batchRpt.ss1.MaxCols
    frm_batchRpt.ss1.BlockMode = True
    frm_batchRpt.ss1.Font.Bold = False
    frm_batchRpt.ss1.Font.Name = "Arial Unicode MS"
    frm_batchRpt.ss1.Font.Size = 10
    frm_batchRpt.ss1.BlockMode = False
  End If
End Sub

Private Sub cmd_exit_Click()
  
  If (m_restartLoop = True) Then
    m_restartLoop = False
    Call unity_main.restart_loop(LOG_DBG_LEVEL3, "Batch Report screen Exit button selected")
  Else
    unity_main.errorstring = "Batch Report screen Exit button selected"
    unity_main.write_error (LOG_DBG_LEVEL3)
  End If
  
  Unload frm_batchRpt
End Sub

Private Sub cmd_delete_Click()
  Dim optVal As Integer
  Dim errMsg As String
  Dim uniMsg As String
  
  unity_main.errorstring = "Batch Report screen Delete File button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  On Error GoTo NO_DELETE
  uniMsg = MLSupport.GGS_Params("fileMsg1", "Are you sure you want to delete %1?", m_batchRptFile)
  optVal = CWrap.ShowMessageBoxW(uniMsg, vbYesNo)
  
  If (optVal = vbYes) Then
    CFile.st_RmFile m_batchRptFile
    unity_main.errorstring = ("User deleted batch file: " & m_batchRptFile)
    unity_main.write_error (LOG_DBG_LEVEL1)
    Unload frm_batchRpt
  End If

  Exit Sub
  
NO_DELETE:
  errMsg = (m_batchRptFile & " file delete error. " & Error$)
  unity_main.errorstring = errMsg
  unity_main.write_error (LOG_DBG_LEVEL1)
  uniMsg = MLSupport.GGS_Params("fileErrMsg8", "%1 file delete error. %2", m_batchRptFile, Error$)
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
End Sub






