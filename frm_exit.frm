VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_exit 
   Caption         =   "InfoStar Exit"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9000
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin HexUniControls.ctlUniButtonImageXP cmd_exit 
      Height          =   700
      Left            =   6240
      TabIndex        =   3
      Top             =   3240
      Width           =   2600
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
      Caption         =   "frm_exit.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_exit.frx":003A
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_exit.frx":005A
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_restart 
      Height          =   700
      Left            =   3180
      TabIndex        =   4
      Top             =   3240
      Width           =   2600
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
      Caption         =   "frm_exit.frx":0076
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_exit.frx":00B6
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_exit.frx":00D6
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   960
      Top             =   3960
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4410
      FormDesignWidth =   9000
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   700
      Left            =   6240
      TabIndex        =   2
      Top             =   2280
      Width           =   2600
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
      Caption         =   "frm_exit.frx":00F2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_exit.frx":011E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_exit.frx":013E
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_shutdown 
      Height          =   700
      Left            =   3180
      TabIndex        =   1
      Top             =   2280
      Width           =   2600
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
      Caption         =   "frm_exit.frx":015A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_exit.frx":018A
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_exit.frx":01AA
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_reboot 
      Height          =   700
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   2600
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
      Caption         =   "frm_exit.frx":01C6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_exit.frx":01F2
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_exit.frx":0212
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   1575
      Left            =   240
      Top             =   240
      Width           =   8415
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_exit.frx":022E
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_exit.frx":02A2
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_exit.frx":02C2
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   1680
      Top             =   3960
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
      Left            =   240
      Top             =   3960
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_exit.frx":02DE
   End
End
Attribute VB_Name = "frm_exit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const EWX_LogOff As Long = 0
Private Const EWX_SHUTDOWN As Long = 1
Private Const EWX_REBOOT As Long = 2
Private Const EWX_FORCE As Long = 4
Private Const EWX_POWEROFF As Long = 8

'The ExitWindowsEx function either logs off, shuts down, or shuts
'down and restarts the system.
Private Declare Function ExitWindowsEx Lib "USER32" _
   (ByVal dwOptions As Long, _
    ByVal dwReserved As Long) As Long

'The GetLastError function returns the calling thread's last-error
'code value. The last-error code is maintained on a per-thread basis.
'Multiple threads do not overwrite each other's last-error code.
Private Declare Function GetLastError Lib "kernel32" () As Long

Private Const mlngWindows95 = 0
Private Const mlngWindowsNT = 1

Public glngWhichWindows32 As Long

'The GetVersion function returns the operating system in use.
Private Declare Function GetVersion Lib "kernel32" () As Long

Private Type LUID
   UsedPart As Long
   IgnoredForNowHigh32BitPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
   TheLuid As LUID
   Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
   PrivilegeCount As Long
   TheLuid As LUID
   Attributes As Long
End Type

'The GetCurrentProcess function returns a pseudohandle for the
'current process.
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

'The OpenProcessToken function opens the access token associated with
'a process.
Private Declare Function OpenProcessToken Lib "advapi32" _
   (ByVal ProcessHandle As Long, _
    ByVal DesiredAccess As Long, _
    TokenHandle As Long) As Long

'The LookupPrivilegeValue function retrieves the locally unique
'identifier (LUID) used on a specified system to locally represent
'the specified privilege name.
Private Declare Function LookupPrivilegeValue Lib "advapi32" _
   Alias "LookupPrivilegeValueA" _
   (ByVal lpSystemName As String, _
    ByVal lpName As String, _
    lpLuid As LUID) As Long

'The AdjustTokenPrivileges function enables or disables privileges
'in the specified access token. Enabling or disabling privileges
'in an access token requires TOKEN_ADJUST_PRIVILEGES access.
Private Declare Function AdjustTokenPrivileges Lib "advapi32" _
   (ByVal TokenHandle As Long, _
    ByVal DisableAllPrivileges As Long, _
    NewState As TOKEN_PRIVILEGES, _
    ByVal BufferLength As Long, _
    PreviousState As TOKEN_PRIVILEGES, _
    ReturnLength As Long) As Long

Private Declare Sub SetLastError Lib "kernel32" _
   (ByVal dwErrCode As Long)

Private Sub AdjustToken()

'********************************************************************
'* This procedure sets the proper privileges to allow a log off or a
'* shut down to occur under Windows NT.
'********************************************************************

   Const TOKEN_ADJUST_PRIVILEGES = &H20
   Const TOKEN_QUERY = &H8
   Const SE_PRIVILEGE_ENABLED = &H2

   Dim hdlProcessHandle As Long
   Dim hdlTokenHandle As Long
   Dim tmpLuid As LUID
   Dim Tkp As TOKEN_PRIVILEGES
   Dim tkpNewButIgnored As TOKEN_PRIVILEGES
   Dim lBufferNeeded As Long

   'Set the error code of the last thread to zero using the
   'SetLast Error function. Do this so that the GetLastError
   'function does not return a value other than zero for no
   'apparent reason.
   SetLastError 0

   'Use the GetCurrentProcess function to set the hdlProcessHandle
   'variable.
   hdlProcessHandle = GetCurrentProcess()

   OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), hdlTokenHandle

   'Get the LUID for shutdown privilege
   LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid

   Tkp.PrivilegeCount = 1    ' One privilege to set
   Tkp.TheLuid = tmpLuid
   Tkp.Attributes = SE_PRIVILEGE_ENABLED

   'Enable the shutdown privilege in the access token of this process
   AdjustTokenPrivileges hdlTokenHandle, _
                         False, _
                         Tkp, _
                         Len(tkpNewButIgnored), _
                         tkpNewButIgnored, _
                         lBufferNeeded
End Sub

Private Sub cmd_exit_Click()
  
  unity_main.errorstring = "InfoStar Exit screen Exit InfoStar button selected"
  unity_main.write_error
  unity_main.errorstring = "InfoStar shutdown"
  unity_main.write_error
  
  Call frm_collect.savescansettings(False)
  Call frm_Inst.savemyinsts(False, False)
  Call frm_dynRptCfg.save_cfg(False, False)

  unity_main.unloadallforms "frm_exit"
  Unload Me
  End
End Sub

Private Sub cmd_reboot_Click()
  
  unity_main.errorstring = "InfoStar Exit screen Reboot button selected"
  unity_main.write_error
  unity_main.errorstring = "System reboot"
  unity_main.write_error
  
  Call frm_collect.savescansettings(False)
  Call frm_Inst.savemyinsts(False, False)
  Call frm_dynRptCfg.save_cfg(False, False)
  
  unity_main.unloadallforms "frm_exit"
  
  If (glngWhichWindows32 = mlngWindowsNT) Then
    AdjustToken
  End If

  ExitWindowsEx (EWX_REBOOT), &HFFFF
End Sub

Private Sub cmd_shutdown_Click()
  
  unity_main.errorstring = "InfoStar Exit screen Shutdown button selected"
  unity_main.write_error
  unity_main.errorstring = "System shutdown"
  unity_main.write_error
  
  Call frm_collect.savescansettings(False)
  Call frm_Inst.savemyinsts(False, False)
  Call frm_dynRptCfg.save_cfg(False, False)
  
  unity_main.unloadallforms "frm_exit"
  
  If (glngWhichWindows32 = mlngWindowsNT) Then
    AdjustToken
  End If

  ExitWindowsEx (EWX_SHUTDOWN Or EWX_FORCE), &HFFFF
End Sub

Private Sub cmd_restart_Click()

  unity_main.errorstring = "InfoStar Exit screen Restart InfoStar button selected"
  unity_main.write_error
  Call frm_collect.savescansettings(False)
  Call frm_Inst.savemyinsts(False, False)
  Call frm_dynRptCfg.save_cfg(False, False)
  
  unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status10", "Ready for New Sample")
  unity_main.Visible = False
  Unload frm_exit

  unity_main.m_smplAutoScan = False
  unity_main.m_smplManualScan = False
  unity_main.m_remoteSmplScan = False
  
  If (unity_main.m_enableRunMode = False) Then
    frm_guilevel.forcemaxgui
  Else
    frm_guipw.loadguipw
    frm_guilevel.pwpassed = False
    frm_guilevel.Visible = True
    frm_guilevel.tmr_pw.enabled = True
  End If
End Sub

Private Sub cmd_cancel_Click()

  Call unity_main.restart_loop(LOG_DBG_LEVEL3, "InfoStar Exit screen Cancel button selected")
  Unload frm_exit
End Sub

Private Sub Form_Deactivate()
  
  Me.ZOrder
End Sub

Private Sub Form_Load()
'********************************************************************
'* When the project starts, check the operating system used by
'* calling the GetVersion function.
'********************************************************************
  Dim lngVersion As Long

  ' Apply language file to form
  MLSupport.ApplyToForm Me

  lngVersion = GetVersion()

  If ((lngVersion And &H80000000) = 0) Then
     glngWhichWindows32 = mlngWindowsNT
  Else
     glngWhichWindows32 = mlngWindows95
  End If
End Sub








