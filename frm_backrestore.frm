VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_backrestore 
   Caption         =   "File Management/Reference Utilities"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4125
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
   Icon            =   "frm_backrestore.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   4125
   StartUpPosition =   3  'Windows Default
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3120
      Top             =   7800
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   7770
      FormDesignWidth =   4125
   End
   Begin HexUniControls.ctlUniFileBoxXP lst_bu 
      Height          =   870
      Left            =   480
      TabIndex        =   7
      Top             =   7800
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tip             =   "frm_backrestore.frx":0442
      Path            =   ""
      Pattern         =   "*.*"
      PatternAlsoForDirs=   0   'False
      ReadOnly        =   -1  'True
      System          =   0   'False
      Hidden          =   0   'False
      PermitNavigation=   -1  'True
      MultiSelect     =   0
      HScroll         =   -1  'True
      ShowFullPath    =   0   'False
      DisplayMode     =   1
      MousePointer    =   0
      MouseIcon       =   "frm_backrestore.frx":0462
   End
   Begin HexUniControls.ctlUniListBoxXP lst_wp 
      Height          =   840
      Left            =   1920
      TabIndex        =   6
      Top             =   7800
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TrapTab         =   0   'False
      Tip             =   "frm_backrestore.frx":047E
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_backrestore.frx":049E
      ManualStart     =   0   'False
      Columns         =   0
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_pptRef 
      Height          =   795
      Left            =   360
      TabIndex        =   5
      Top             =   3960
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   1402
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_backrestore.frx":04BA
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_backrestore.frx":0534
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_backrestore.frx":0554
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_modelRestore 
      Height          =   795
      Left            =   360
      TabIndex        =   4
      Top             =   3000
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   1402
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_backrestore.frx":0570
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_backrestore.frx":05BE
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_backrestore.frx":05DE
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_modelBackup 
      Height          =   795
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   1402
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_backrestore.frx":05FA
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_backrestore.frx":0646
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_backrestore.frx":0666
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_exit 
      Cancel          =   -1  'True
      Height          =   795
      Left            =   360
      TabIndex        =   0
      Top             =   6840
      Width           =   3405
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_backrestore.frx":0682
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_backrestore.frx":06AA
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_backrestore.frx":06CA
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cfgRestore 
      Height          =   795
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   1402
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_backrestore.frx":06E6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_backrestore.frx":0752
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_backrestore.frx":0772
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cfgBackup 
      Height          =   795
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   3405
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_backrestore.frx":078E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_backrestore.frx":07F4
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_backrestore.frx":0814
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_pptExtRef 
      Height          =   795
      Left            =   360
      TabIndex        =   8
      Top             =   4920
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   1402
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_backrestore.frx":0830
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_backrestore.frx":08AA
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_backrestore.frx":08CA
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_intRefCalMgmt 
      Height          =   795
      Left            =   360
      TabIndex        =   9
      Top             =   5880
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   1402
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_backrestore.frx":08E6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_backrestore.frx":0958
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_backrestore.frx":0978
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   0
      Top             =   7800
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
      Left            =   3600
      Top             =   7800
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_backrestore.frx":0994
   End
End
Attribute VB_Name = "frm_backrestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type STARTUPINFO
      cb As Long
      lpReserved As String
      lpDesktop As String
      lpTitle As String
      dwX As Long
      dwY As Long
      dwXSize As Long
      dwYSize As Long
      dwXCountChars As Long
      dwYCountChars As Long
      dwFillAttribute As Long
      dwFlags As Long
      wShowWindow As Integer
      cbReserved2 As Integer
      lpReserved2 As Long
      hStdInput As Long
      hStdOutput As Long
      hStdError As Long
   End Type

   Private Type PROCESS_INFORMATION
      hProcess As Long
      hThread As Long
      dwProcessID As Long
      dwThreadID As Long
   End Type
 Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
      hHandle As Long, ByVal dwMilliseconds As Long) As Long

   Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
      lpApplicationName As String, ByVal lpCommandLine As String, ByVal _
      lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
      ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
      ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
      lpStartupInfo As STARTUPINFO, lpProcessInformation As _
      PROCESS_INFORMATION) As Long

   Private Declare Function CloseHandle Lib "kernel32" _
      (ByVal hObject As Long) As Long

   Private Declare Function GetExitCodeProcess Lib "kernel32" _
      (ByVal hProcess As Long, lpExitCode As Long) As Long

   Private Const NORMAL_PRIORITY_CLASS = &H20&
   Private Const INFINITE = -1&

   Public Function ExecCmd(cmdline$)
      Dim proc As PROCESS_INFORMATION
      Dim start As STARTUPINFO
      Dim ret&
      
      ' Initialize the STARTUPINFO structure:
      start.cb = Len(start)

      ' Start the shelled application:
      ret& = CreateProcessA(vbNullString, cmdline$, 0&, 0&, 1&, _
         NORMAL_PRIORITY_CLASS, 0&, vbNullString, start, proc)

      ' Wait for the shelled application to finish:
         ret& = WaitForSingleObject(proc.hProcess, INFINITE)
         Call GetExitCodeProcess(proc.hProcess, ret&)
         Call CloseHandle(proc.hThread)
         Call CloseHandle(proc.hProcess)
         ExecCmd = ret&
   End Function


Public Sub backup_inis()
  Dim sourceDir As String
  Dim targetDir As String
  Dim cnt(6) As Integer
  Dim errMsg As String
  Dim uniMsg As String
  
  On Error Resume Next
  sourceDir = CFG_DIR
  targetDir = CFG_BKUP_DIR

  If (CreatePath(targetDir) = False) Then
    errMsg = (targetDir & " directory cannot be created")
    unity_main.errorstring = errMsg
    unity_main.write_error (LOG_DBG_LEVEL1)
    uniMsg = MLSupport.GGS_Params("frm_backrestore.errMsg1", "%1 directory cannot be created", targetDir)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  Else
    If (CreatePath(PRODUCTS_CFG_BKUP_DIR) = True) Then
      ' Backup config directory files
      CFile.st_RmFiles (targetDir & "*" & CFG_FILE_EXT)  'get rid of old backups
      Call copy_files(sourceDir, targetDir, CFG_FILE_EXT, cnt(1))
  
      CFile.st_RmFiles (targetDir & "*" & USER_INPUTS_CFG_FILE_EXT)  'get rid of old backups
      Call copy_files(sourceDir, targetDir, USER_INPUTS_CFG_FILE_EXT, cnt(2))
  
      CFile.st_RmFiles (targetDir & "*" & PW_FILE_EXT)  'get rid of old backups
      Call copy_files(sourceDir, targetDir, PW_FILE_EXT, cnt(3))
  
      ' Backup products directory files
      sourceDir = PRODUCTS_CFG_DIR
      targetDir = PRODUCTS_CFG_BKUP_DIR
      CFile.st_RmFiles (targetDir & "*" & CFG_FILE_EXT)  'get rid of old backups
      Call copy_files(sourceDir, targetDir, CFG_FILE_EXT, cnt(4))
      
      If (CreatePath(EXT_REFS_CFG_BKUP_DIR) = True) Then
        ' Backup external reference directory files
        sourceDir = EXT_REFS_CFG_DIR
        targetDir = EXT_REFS_CFG_BKUP_DIR
        CFile.st_RmFiles (targetDir & "*" & CFG_FILE_EXT)  'get rid of old backups
        Call copy_files(sourceDir, targetDir, CFG_FILE_EXT, cnt(5))
      
        If (CreatePath(SCAN_BATCHES_CFG_BKUP_DIR) = True) Then
          ' Backup batch scan directory files
          sourceDir = SCAN_BATCHES_CFG_DIR
          targetDir = SCAN_BATCHES_CFG_BKUP_DIR
          CFile.st_RmFiles (targetDir & "*" & SCAN_BATCH_FILE_EXT)  'get rid of old backups
          Call copy_files(sourceDir, targetDir, SCAN_BATCH_FILE_EXT, cnt(6))
      
          errMsg = ("Number of configuration files backed up = " & Str$(cnt(1) + cnt(2) + cnt(3) + cnt(4) + cnt(5) + cnt(6)))
          unity_main.errorstring = errMsg
          unity_main.write_error (LOG_DBG_LEVEL1)
          uniMsg = MLSupport.GGS_Params("frm_backrestore.msg2", "Number of configuration files backed up = %1", Str$(cnt(1) + cnt(2) + cnt(3) + cnt(4) + cnt(5) + cnt(6)))
          CWrap.ShowMessageBoxW uniMsg
        Else
          errMsg = (SCAN_BATCHES_CFG_BKUP_DIR & " directory cannot be created")
          unity_main.errorstring = errMsg
          unity_main.write_error (LOG_DBG_LEVEL1)
          uniMsg = MLSupport.GGS_Params("frm_backrestore.errMsg1", "%1 directory cannot be created", SCAN_BATCHES_CFG_BKUP_DIR)
          CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
        End If
      Else
        errMsg = (EXT_REFS_CFG_BKUP_DIR & " directory cannot be created")
        unity_main.errorstring = errMsg
        unity_main.write_error (LOG_DBG_LEVEL1)
        uniMsg = MLSupport.GGS_Params("frm_backrestore.errMsg1", "%1 directory cannot be created", EXT_REFS_CFG_BKUP_DIR)
        CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
      End If
    Else
      errMsg = (PRODUCTS_CFG_BKUP_DIR & " directory cannot be created")
      unity_main.errorstring = errMsg
      unity_main.write_error (LOG_DBG_LEVEL1)
      uniMsg = MLSupport.GGS_Params("frm_backrestore.errMsg1", "%1 directory cannot be created", PRODUCTS_CFG_BKUP_DIR)
      CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
    End If
  End If
End Sub

Public Sub backup_models()
  Dim srcDir As String
  Dim destDir As String
  Dim newDir As String
  Dim csf As New clsSearchFiles
  Dim coll As Collection
  Dim ii As Integer
  Dim csfi As clsSearchFilesFInfo
  Dim p As Integer
  Dim n As Integer
  Dim errMsg As String
  Dim uniMsg As String
  
  On Error Resume Next
  srcDir = MODELS_DIR
  destDir = MODELS_BKUP_DIR
  
  ' Backup models directory files
  If (CreatePath(destDir) = False) Then
    errMsg = (destDir & " directory cannot be created")
    unity_main.errorstring = errMsg
    unity_main.write_error (LOG_DBG_LEVEL1)
    uniMsg = MLSupport.GGS_Params("frm_backrestore.errMsg1", "%1 directory cannot be created", destDir)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  Else
    Call unprotect_files(destDir)
    CFile.st_RmFiles (destDir & "*.*")   'get rid of old backups
    Call copy_files(srcDir, destDir, ".*", p)
    n = n + p
  
    Set coll = csf.SearchInPath(srcDir, esfw_only_dirs, True, False, False, -1, True)
  
    For ii = 1 To coll.Count
      Set csfi = coll.Item(ii)
      srcDir = csfi.sPathFileName & "\"
      newDir = destDir & csfi.sFileName & "\"
    
      If (CreatePath(newDir) = False) Then
        errMsg = (newDir & " directory cannot be created")
        unity_main.errorstring = errMsg
        unity_main.write_error (LOG_DBG_LEVEL1)
        uniMsg = MLSupport.GGS_Params("frm_backrestore.errMsg1", "%1 directory cannot be created", newDir)
        CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
      Else
        Call unprotect_files(newDir)
        CFile.st_RmFiles (newDir & "*.*")   'get rid of old backups
        Call copy_files(srcDir, newDir, ".*", p)
        n = n + p
    
        If (Not (csfi.oCollSubFiles Is Nothing)) Then
          Dim subCol As Collection
          Dim jj As Integer
          Dim newSubDir As String
      
          Set subCol = csfi.oCollSubFiles
    
          For jj = 1 To subCol.Count
            Set csfi = subCol.Item(jj)
            srcDir = csfi.sPathFileName & "\"
            newSubDir = newDir & csfi.sFileName & "\"
        
            If (CreatePath(newSubDir) = False) Then
              errMsg = (newSubDir & " directory cannot be created")
              unity_main.errorstring = errMsg
              unity_main.write_error (LOG_DBG_LEVEL1)
              uniMsg = MLSupport.GGS_Params("frm_backrestore.errMsg1", "%1 directory cannot be created", newSubDir)
              CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
            Else
              Call unprotect_files(newSubDir)
              CFile.st_RmFiles (newSubDir & "*.*")   'get rid of old backups
              Call copy_files(srcDir, newSubDir, ".*", p)
              n = n + p
            End If
          Next jj
        End If
    
        DoEvents
      End If
    Next ii
  
    errMsg = ("Number of model files backed up " & Str$(n))
    unity_main.errorstring = errMsg
    unity_main.write_error (LOG_DBG_LEVEL1)
    uniMsg = MLSupport.GGS_Params("frm_backrestore.msg1", "Number of model files backed up %1", Str$(n))
    CWrap.ShowMessageBoxW uniMsg
  End If
End Sub

Public Sub restore_inis()
  Dim sourceDir As String
  Dim targetDir As String
  Dim cnt(6) As Integer
  Dim errMsg As String
  Dim uniMsg As String
  
  If (any_files_exist(CFG_BKUP_DIR, ("*" & CFG_FILE_EXT)) = False) Then
    errMsg = (CFG_BKUP_DIR & " directory and/or files do not exist. No files to restore")
    unity_main.errorstring = errMsg
    unity_main.write_error (LOG_DBG_LEVEL1)
    uniMsg = MLSupport.GGS_Params("frm_backrestore.errMsg2", "%1 directory and/or files do not exist. No files to restore", CFG_BKUP_DIR)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  Else
    ' Restore config directory files
    sourceDir = CFG_BKUP_DIR
    targetDir = CFG_DIR
    CFile.st_RmFiles (targetDir & "*" & CFG_FILE_EXT)  'get rid of possibly corrupt ini
    Call copy_files(sourceDir, targetDir, CFG_FILE_EXT, cnt(1))
  
    CFile.st_RmFiles (targetDir & "*" & USER_INPUTS_CFG_FILE_EXT)  'get rid of possibly corrupt ini
    Call copy_files(sourceDir, targetDir, USER_INPUTS_CFG_FILE_EXT, cnt(2))
  
    CFile.st_RmFiles (targetDir & "*" & PW_FILE_EXT)  'get rid of possibly corrupt ini
    Call copy_files(sourceDir, targetDir, PW_FILE_EXT, cnt(3))
  
    ' Restore any product directory files
    If (any_files_exist(PRODUCTS_CFG_BKUP_DIR, ("*" & CFG_FILE_EXT)) = True) Then
      sourceDir = PRODUCTS_CFG_BKUP_DIR
      targetDir = PRODUCTS_CFG_DIR
      CFile.st_RmFiles (targetDir & "*" & CFG_FILE_EXT)  'get rid of possibly corrupt inis
      Call copy_files(sourceDir, targetDir, CFG_FILE_EXT, cnt(4))
    End If
    
    ' Restore any external reference directory files
    If (any_files_exist(EXT_REFS_CFG_BKUP_DIR, ("*" & CFG_FILE_EXT)) = True) Then
      sourceDir = EXT_REFS_CFG_BKUP_DIR
      targetDir = EXT_REFS_CFG_DIR
      CFile.st_RmFiles (targetDir & "*" & CFG_FILE_EXT)  'get rid of possibly corrupt inis
      Call copy_files(sourceDir, targetDir, CFG_FILE_EXT, cnt(5))
    End If
    
    ' Restore any external reference directory files
    If (any_files_exist(SCAN_BATCHES_CFG_BKUP_DIR, ("*" & SCAN_BATCH_FILE_EXT)) = True) Then
      sourceDir = SCAN_BATCHES_CFG_BKUP_DIR
      targetDir = SCAN_BATCHES_CFG_DIR
      CFile.st_RmFiles (targetDir & "*" & SCAN_BATCH_FILE_EXT)  'get rid of possibly corrupt inis
      Call copy_files(sourceDir, targetDir, SCAN_BATCH_FILE_EXT, cnt(6))
    End If
    
    errMsg = ("Number of configuration files restored = " & Str$(cnt(1) + cnt(2) + cnt(3) + cnt(4) + cnt(5) + cnt(6)))
    unity_main.errorstring = errMsg
    unity_main.write_error (LOG_DBG_LEVEL1)
    uniMsg = MLSupport.GGS_Params("frm_backrestore.msg3", "Number of configuration files restored = %1", Str$(cnt(1) + cnt(2) + cnt(3) + cnt(4) + cnt(5) + cnt(6)))
    CWrap.ShowMessageBoxW uniMsg
  End If
End Sub

Public Sub restore_models()
  Dim srcDir As String
  Dim destDir As String
  Dim newDir As String
  Dim csf As New clsSearchFiles
  Dim coll As Collection
  Dim ii As Integer
  Dim csfi As clsSearchFilesFInfo
  Dim p As Integer
  Dim n As Integer
  Dim errMsg As String
  Dim uniMsg As String
  
  srcDir = MODELS_BKUP_DIR
  destDir = MODELS_DIR
  
  If (any_files_exist(srcDir, "*.*") = False) Then
    errMsg = (srcDir & " directory and/or files do not exist. No files to restore")
    unity_main.errorstring = errMsg
    unity_main.write_error (LOG_DBG_LEVEL1)
    uniMsg = MLSupport.GGS_Params("frm_backrestore.errMsg2", "%1 directory and/or files do not exist. No files to restore", srcDir)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  Else
    ' Restore models directory files
    On Error Resume Next
    CFile.st_RmFiles (destDir & "*.*")  'get rid of possibly corrupt models
    Call copy_files(srcDir, destDir, ".*", p)
    n = n + p
    
    Set coll = csf.SearchInPath(srcDir, esfw_only_dirs, True, False, False, -1, True)
  
    For ii = 1 To coll.Count
      Set csfi = coll.Item(ii)
      srcDir = csfi.sPathFileName & "\"
      newDir = destDir & csfi.sFileName & "\"
    
      If (CreatePath(newDir) = False) Then
        errMsg = (newDir & " directory cannot be created")
        unity_main.errorstring = errMsg
        unity_main.write_error (LOG_DBG_LEVEL1)
        uniMsg = MLSupport.GGS_Params("frm_backrestore.errMsg1", "%1 directory cannot be created", newDir)
        CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
      Else
        Call unprotect_files(newDir)
        CFile.st_RmFiles (newDir & "*.*")   'get rid of old backups
        Call copy_files(srcDir, newDir, ".*", p)
        n = n + p
    
        If (Not (csfi.oCollSubFiles Is Nothing)) Then
          Dim subCol As Collection
          Dim jj As Integer
          Dim newSubDir As String
      
          Set subCol = csfi.oCollSubFiles
    
          For jj = 1 To subCol.Count
            Set csfi = subCol.Item(jj)
            srcDir = csfi.sPathFileName & "\"
            newSubDir = newDir & csfi.sFileName & "\"
      
            If (CreatePath(newSubDir) = False) Then
              errMsg = (newSubDir & " directory cannot be created")
              unity_main.errorstring = errMsg
              unity_main.write_error (LOG_DBG_LEVEL1)
              uniMsg = MLSupport.GGS_Params("frm_backrestore.errMsg1", "%1 directory cannot be created", newSubDir)
              CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
            Else
              Call unprotect_files(newSubDir)
              CFile.st_RmFiles (newSubDir & "*.*")   'get rid of old backups
              Call copy_files(srcDir, newSubDir, ".*", p)
              n = n + p
            End If
          Next jj
        End If
      End If
    
      DoEvents
    Next ii
    
    errMsg = ("Number of model files restored = " & Str$(n))
    unity_main.errorstring = errMsg
    unity_main.write_error (LOG_DBG_LEVEL1)
    uniMsg = MLSupport.GGS_Params("frm_backrestore.msg4", "Number of model files restored = %1", Str$(n))
    CWrap.ShowMessageBoxW uniMsg
  End If
End Sub

Public Sub unprotect_files(dirname As String)
  Dim fpath As String
  Dim i As Long
  Dim cnt  As Long
  Dim zz As Integer

  fpath = dirname
  frm_backrestore.lst_bu.Path = fpath

  ' Unprotect all files
  For i = 0 To lst_bu.ListCount - 1
    zz = CFile.st_GetFileAttr(fpath & lst_bu.List(i))
    
    If (zz <> 0) Then         ' = normal   1 = read only
      On Error GoTo NEXT_FILE
      CFile.st_SetFileAttr (fpath & lst_bu.List(i)), vbNormal
    End If
NEXT_FILE:
  Next i
End Sub

Private Sub cmd_cfgBackup_Click()
  Dim uniMsg As String
  Dim optVal As Integer
  
  unity_main.errorstring = "File Management/Reference Utilities screen Backup Existing Configuration Files button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)

  uniMsg = MLSupport.GSS("frm_backrestore", "msg2", "Are you sure you want to archive existing configuration files? Existing backup/archived files will be deleted.")
  optVal = CWrap.ShowMessageBoxW(uniMsg, vbYesNo)
  
  If (optVal = vbYes) Then
    Call frm_collect.savescansettings(False)
    Call frm_Inst.savemyinsts(False, False)
    Call frm_dynRptCfg.save_cfg(False, False)
    Call backup_inis
  End If
End Sub

Private Sub cmd_exit_Click()
  
  unity_main.errorstring = "File Management/Reference Utilities screen Exit button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Unload frm_backrestore
End Sub

Private Sub cmd_modelBackup_Click()
  Dim optVal As Integer
  Dim errMsg As String
  Dim uniMsg As String
  
  unity_main.errorstring = "File Management/Reference Utilities screen Backup All Model Files button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  If (CreatePath(MODELS_BKUP_DIR) = False) Then
    errMsg = (MODELS_BKUP_DIR & " directory cannot be created")
    unity_main.errorstring = errMsg
    unity_main.write_error (LOG_DBG_LEVEL1)
    uniMsg = MLSupport.GGS_Params("frm_backrestore.errMsg1", "%1 directory cannot be created", MODELS_BKUP_DIR)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  Else
    uniMsg = MLSupport.GSS("frm_backrestore", "msg1", "Are you sure you want to archive existing model files? Existing backup/archived mdoels will be deleted.")
    optVal = CWrap.ShowMessageBoxW(uniMsg, vbYesNo)
    
    If (optVal = vbYes) Then
      backup_models
    End If
  End If
End Sub

Private Sub cmd_modelRestore_Click()
  Dim optVal As Integer
  Dim uniMsg As String
   
  unity_main.errorstring = "File Management/Reference Utilities screen Restore All Model Files button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  uniMsg = MLSupport.GSS("frm_backrestore", "msg4", "Are you sure you want to restore archived model files? Current model files will be deleted.")
  optVal = CWrap.ShowMessageBoxW(uniMsg, vbYesNo)
  
  If (optVal = vbYes) Then
    restore_models
  End If
End Sub

#If SSTAR Then
Private Sub cmd_pptExtRef_Click()

  unity_main.errorstring = "File Management/Reference Utilities screen Collect External Reference Qualification Data button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  frm_collect.build_ref_name_list "external"
  frm_extRefPPT.m_selection = vbCancel
  frm_extRefPPT.Show 1

  If (frm_extRefPPT.m_selection = vbOK) Then
    ' Flag to perform external reference qualification scan
    unity_main.m_extRefPPTScan = True
    unity_main.tmr_all.enabled = True
    
    unity_main.pw_open = False ' will have to reenter password next time in utils
    Unload frm_backrestore
    Unload frmUtils
  End If
End Sub
#End If

#If SSTAR Then
Private Sub cmd_pptRef_Click() '12/25
  
  unity_main.errorstring = "File Management/Reference Utilities screen Collect Internal Reference Qualification Data button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  frm_refPPT.txt_startWvln.Text = unity_main.m_smplStartWvln
  frm_refPPT.txt_endWvln.Text = unity_main.m_smplEndWvln
  frm_refPPT.txt_minWvln.Text = unity_main.m_minWvln
  frm_refPPT.txt_maxWvln.Text = unity_main.m_maxWvln
  frm_refPPT.m_selection = vbCancel

  frm_refPPT.Show 1

  If (frm_refPPT.m_selection = vbOK) Then
    ' Flag to perform internal reference qualification scan
    unity_main.m_intRefPPTScan = True
    unity_main.tmr_all.enabled = True
  
    unity_main.pw_open = False ' will have to reenter password next time in utils
    Unload frm_backrestore
    Unload frmUtils
  End If
End Sub
#End If

Private Sub cmd_cfgRestore_Click()
  Dim optVal As Integer
  Dim uniMsg As String

  unity_main.errorstring = "File Management/Reference Utilities screen Restore Last Saved Configuration Files button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)

  uniMsg = MLSupport.GSS("frm_backrestore", "msg3", "Are you sure you want to restore archived configuration files? Current configuration files will be deleted.")
  optVal = CWrap.ShowMessageBoxW(uniMsg, vbYesNo)
  
  If (optVal = vbYes) Then
    restore_inis
    
    ' Reload configuration files
    frm_Inst.Loadmyinstini
    Call frm_csvCfg.load_cfg(True)
    Call unity_main.load_prod_file("", True)
  End If
End Sub

#If SSTAR Then
Private Sub cmd_intRefCalMgmt_Click()

  unity_main.errorstring = "File Management/Reference Utilities screen Internal Reference Calibration Management button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
 frm_intRefCalMgmt.m_restartFlg = False
  frm_intRefCalMgmt.Show 1
' get handle to new process.
 'ExecCmd (ROOT_DISK & REFRECAL_DIR & REFRECAL_EXEC)

'ExecCmd ("c:\unity\software\IntRefCalMgmtUtil.exec")
 
 
End Sub


#End If

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
  
#If ABBFT Then
  cmd_pptRef.Visible = False
  cmd_pptExtRef.Visible = False
  cmd_intRefCalMgmt.Visible = False
#Else
  ' Hide Internal Reference Calibration Management button if non-calibrated TW system
  If (unity_main.m_allowIntRefCalAccess = False) Then
    cmd_intRefCalMgmt.Visible = False
  End If
  'hide for 2500X as well.
  If (unity_main.m_smplTable = 6) Then
    cmd_intRefCalMgmt.Visible = False
  End If
#End If
End Sub








