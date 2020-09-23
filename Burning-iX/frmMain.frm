VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Burning-iX v1.0.1ß (beta) 2009 © Salvo Cortesiano."
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11565
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   11565
   Begin VB.TextBox txtInfo 
      Height          =   2040
      Left            =   4335
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   68
      Text            =   "frmMain.frx":23D2
      Top             =   4905
      Width           =   7125
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   10095
      Top             =   165
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Frame7 
      Caption         =   "Commands"
      Height          =   735
      Left            =   4260
      TabIndex        =   58
      Top             =   7005
      Width           =   7260
      Begin VB.PictureBox Picture7 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   60
         ScaleHeight     =   495
         ScaleWidth      =   7155
         TabIndex        =   59
         Top             =   165
         Width           =   7155
         Begin VB.CommandButton cmdExit 
            Caption         =   "Exit"
            Height          =   345
            Left            =   5745
            TabIndex        =   64
            ToolTipText     =   "Close this Program"
            Top             =   105
            Width           =   1335
         End
         Begin VB.CommandButton cmdEject 
            Caption         =   "Eject"
            Enabled         =   0   'False
            Height          =   345
            Left            =   4485
            TabIndex        =   63
            ToolTipText     =   "Eject the current Drive"
            Top             =   105
            Width           =   1125
         End
         Begin VB.CommandButton cmdAbort 
            Caption         =   "Abort"
            Enabled         =   0   'False
            Height          =   345
            Left            =   3315
            TabIndex        =   62
            ToolTipText     =   "Abort current Task"
            Top             =   105
            Width           =   1095
         End
         Begin VB.CommandButton cmdBurning 
            Caption         =   "Burning"
            Enabled         =   0   'False
            Height          =   345
            Left            =   1980
            TabIndex        =   61
            ToolTipText     =   "Start Burning"
            Top             =   105
            Width           =   1260
         End
         Begin VB.CommandButton cmdBurnImage 
            Caption         =   "Burning Image"
            Enabled         =   0   'False
            Height          =   345
            Left            =   45
            TabIndex        =   60
            ToolTipText     =   "Start/Create a ISO file"
            Top             =   105
            Width           =   1860
         End
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Option Writing"
      Height          =   2940
      Left            =   45
      TabIndex        =   40
      Top             =   4800
      Width           =   4185
      Begin VB.PictureBox Picture6 
         BorderStyle     =   0  'None
         Height          =   2670
         Left            =   30
         ScaleHeight     =   2670
         ScaleWidth      =   4125
         TabIndex        =   41
         Top             =   225
         Width           =   4125
         Begin VB.CheckBox CheckOnlyISO 
            Caption         =   "Create only ISO Image compatible"
            Height          =   240
            Left            =   45
            TabIndex        =   57
            Top             =   2430
            Width           =   4035
         End
         Begin VB.ComboBox cmbTypeSupport 
            Height          =   330
            ItemData        =   "frmMain.frx":2D0E
            Left            =   1020
            List            =   "frmMain.frx":2D18
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   1200
            Width           =   3045
         End
         Begin VB.ComboBox cmbErasing 
            Height          =   330
            ItemData        =   "frmMain.frx":2D3D
            Left            =   1530
            List            =   "frmMain.frx":2D47
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   1815
            Width           =   2535
         End
         Begin VB.CheckBox CheckSimulate 
            Caption         =   "Not Write CD/DVD Simulate Writing"
            Height          =   240
            Left            =   45
            TabIndex        =   49
            Top             =   2175
            Width           =   4035
         End
         Begin VB.CheckBox CheckEraseDisk 
            Caption         =   "Erase Disk Before Burning"
            Height          =   240
            Left            =   45
            TabIndex        =   48
            Top             =   1560
            Width           =   4035
         End
         Begin VB.ComboBox cmbSystem 
            Height          =   330
            ItemData        =   "frmMain.frx":2D76
            Left            =   1425
            List            =   "frmMain.frx":2D83
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   825
            Width           =   2640
         End
         Begin VB.ComboBox cmbMode 
            Height          =   330
            ItemData        =   "frmMain.frx":2DB0
            Left            =   705
            List            =   "frmMain.frx":2DBA
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   420
            Width           =   3360
         End
         Begin VB.ComboBox cmbType 
            Height          =   330
            ItemData        =   "frmMain.frx":2DD1
            Left            =   705
            List            =   "frmMain.frx":2DDE
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   30
            Width           =   3360
         End
         Begin VB.Label Label6 
            Caption         =   "Support:"
            Height          =   255
            Left            =   90
            TabIndex        =   53
            Top             =   1245
            Width           =   915
         End
         Begin VB.Label label134 
            Caption         =   "CDRW Erasing:"
            Height          =   270
            Left            =   90
            TabIndex        =   51
            Top             =   1860
            Width           =   1365
         End
         Begin VB.Label Label12 
            Caption         =   "File System:"
            Height          =   270
            Left            =   90
            TabIndex        =   47
            Top             =   870
            Width           =   1365
         End
         Begin VB.Label Label11 
            Caption         =   "Mode:"
            Height          =   270
            Left            =   90
            TabIndex        =   44
            Top             =   465
            Width           =   630
         End
         Begin VB.Label Label10 
            Caption         =   "Type:"
            Height          =   270
            Left            =   90
            TabIndex        =   42
            Top             =   60
            Width           =   630
         End
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Option File Log"
      Height          =   1455
      Left            =   8235
      TabIndex        =   27
      Top             =   3315
      Width           =   3255
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   45
         ScaleHeight     =   1215
         ScaleWidth      =   3165
         TabIndex        =   28
         Top             =   195
         Width           =   3165
         Begin VB.CommandButton cmdBurFile 
            Caption         =   "Browse Folder and Burning"
            Height          =   345
            Left            =   60
            TabIndex        =   67
            ToolTipText     =   "Browse Folder and Burn..."
            Top             =   765
            Width           =   3015
         End
         Begin VB.CheckBox CheckAppend 
            Caption         =   "Append response into LOG"
            Height          =   240
            Left            =   15
            TabIndex        =   30
            Top             =   345
            UseMaskColor    =   -1  'True
            Width           =   3105
         End
         Begin VB.CheckBox CheckLog 
            Caption         =   "Write responce into LOG"
            Height          =   240
            Left            =   15
            TabIndex        =   29
            Top             =   60
            Value           =   1  'Checked
            Width           =   3105
         End
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "                      ISO File"
      Height          =   600
      Left            =   30
      TabIndex        =   11
      Top             =   1485
      Width           =   11490
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   60
         ScaleHeight     =   375
         ScaleWidth      =   11385
         TabIndex        =   12
         Top             =   180
         Width           =   11385
         Begin VB.Timer TBurning 
            Enabled         =   0   'False
            Left            =   -180
            Top             =   360
         End
         Begin VB.CommandButton cmdPath 
            Caption         =   "..."
            Height          =   285
            Left            =   10830
            TabIndex        =   14
            ToolTipText     =   "Select a File ISO to Burning"
            Top             =   30
            Width           =   495
         End
         Begin VB.TextBox txtPath 
            Height          =   255
            Left            =   2385
            TabIndex        =   13
            Top             =   45
            Width           =   8370
         End
         Begin VB.Label lblISOFileSize 
            Caption         =   "Size:"
            Height          =   255
            Left            =   45
            TabIndex        =   65
            Top             =   0
            Width           =   2280
         End
         Begin VB.Label Label9 
            Caption         =   "Time Start:"
            Height          =   255
            Left            =   45
            TabIndex        =   33
            Top             =   195
            Width           =   1260
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "00:00:00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   240
            Left            =   1155
            TabIndex        =   32
            Top             =   195
            Width           =   1185
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "NEROCom Log/Messages"
      Height          =   1455
      Left            =   30
      TabIndex        =   9
      Top             =   3315
      Width           =   8190
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   1200
         Left            =   45
         ScaleHeight     =   1200
         ScaleWidth      =   8115
         TabIndex        =   10
         Top             =   225
         Width           =   8115
         Begin VB.TextBox lst_Messages 
            Height          =   1125
            Left            =   30
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   26
            Top             =   30
            Width           =   8040
         End
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1155
      Left            =   30
      TabIndex        =   3
      Top             =   2115
      Width           =   11490
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   990
         Left            =   60
         ScaleHeight     =   990
         ScaleWidth      =   11385
         TabIndex        =   4
         Top             =   135
         Width           =   11385
         Begin VB.TextBox txtSession 
            Height          =   255
            Left            =   6135
            TabIndex        =   55
            Top             =   375
            Width           =   5205
         End
         Begin VB.TextBox ISOTrackName 
            Height          =   255
            Left            =   6135
            TabIndex        =   19
            Top             =   45
            Width           =   5205
         End
         Begin ComctlLib.ProgressBar fme_Progress 
            Height          =   180
            Left            =   45
            TabIndex        =   5
            Top             =   240
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   318
            _Version        =   327682
            Appearance      =   1
         End
         Begin ComctlLib.ProgressBar pgs_Burn 
            Height          =   180
            Left            =   45
            TabIndex        =   6
            Top             =   735
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   318
            _Version        =   327682
            Appearance      =   1
         End
         Begin ComctlLib.ProgressBar pgs_Buffer 
            Height          =   180
            Left            =   5715
            TabIndex        =   21
            Top             =   735
            Width           =   5070
            _ExtentX        =   8943
            _ExtentY        =   318
            _Version        =   327682
            Appearance      =   1
         End
         Begin VB.Label Label13 
            Caption         =   "Label Session:"
            Height          =   255
            Left            =   4575
            TabIndex        =   56
            Top             =   390
            Width           =   1515
         End
         Begin VB.Label Label7 
            Caption         =   "Label Volume:"
            Height          =   255
            Left            =   4665
            TabIndex        =   31
            Top             =   45
            Width           =   1440
         End
         Begin VB.Label lblPathISOFILE 
            Height          =   195
            Left            =   4230
            TabIndex        =   25
            Top             =   -30
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label lblTCriting 
            Caption         =   "00%"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3795
            TabIndex        =   24
            Top             =   0
            Width           =   570
         End
         Begin VB.Label lblTWriting 
            Caption         =   "00%"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3810
            TabIndex        =   23
            Top             =   495
            Width           =   570
         End
         Begin VB.Label lblPercentuale 
            Caption         =   "00%"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   10875
            TabIndex        =   22
            Top             =   690
            Width           =   450
         End
         Begin VB.Label Label5 
            Caption         =   "Buffer Size:"
            Height          =   255
            Left            =   4350
            TabIndex        =   20
            Top             =   690
            Width           =   1350
         End
         Begin VB.Label Label2 
            Caption         =   "Total Writing:"
            Height          =   195
            Left            =   60
            TabIndex        =   8
            Top             =   495
            Width           =   2955
         End
         Begin VB.Label Label1 
            Caption         =   "Buffer Task:"
            Height          =   195
            Left            =   60
            TabIndex        =   7
            Top             =   0
            Width           =   3210
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Support CD/DVD-DRW                         Write Speed     ISO      YBC     LOG       BEND        BACKUP"
      Height          =   690
      Left            =   30
      TabIndex        =   0
      Top             =   780
      Width           =   11490
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   75
         ScaleHeight     =   435
         ScaleWidth      =   11385
         TabIndex        =   1
         Top             =   225
         Width           =   11385
         Begin VB.OptionButton opt_B 
            Caption         =   "BACKUP"
            Height          =   270
            Index           =   4
            Left            =   10110
            TabIndex        =   54
            ToolTipText     =   "Write all Folders/File in BACKUP"
            Top             =   60
            Width           =   1185
         End
         Begin VB.OptionButton opt_B 
            Caption         =   "ISO"
            Height          =   270
            Index           =   0
            Left            =   5985
            TabIndex        =   34
            ToolTipText     =   "Write all Folders/File in ISO"
            Top             =   60
            Width           =   930
         End
         Begin VB.OptionButton opt_B 
            Caption         =   "YBC"
            Height          =   270
            Index           =   1
            Left            =   6930
            TabIndex        =   18
            ToolTipText     =   "Write all Folders/File in YBC"
            Top             =   60
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.OptionButton opt_B 
            Caption         =   "LOG"
            Height          =   270
            Index           =   2
            Left            =   7830
            TabIndex        =   17
            ToolTipText     =   "Write all Folders/File in LOG"
            Top             =   60
            Width           =   990
         End
         Begin VB.OptionButton opt_B 
            Caption         =   "BEND"
            Height          =   270
            Index           =   3
            Left            =   8820
            TabIndex        =   16
            ToolTipText     =   "Write all Folders/File in BEND"
            Top             =   60
            Width           =   1275
         End
         Begin VB.ComboBox cmbWSpeed 
            Height          =   330
            ItemData        =   "frmMain.frx":2E36
            Left            =   4590
            List            =   "frmMain.frx":2E55
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   30
            Width           =   1140
         End
         Begin VB.ComboBox lst_AvailableDevices 
            Height          =   330
            Left            =   60
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   30
            Width           =   4365
         End
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   11025
      Picture         =   "frmMain.frx":2E7A
      Top             =   210
      Width           =   480
   End
   Begin VB.Label lblVN 
      BackColor       =   &H00FFFFFF&
      Caption         =   "n/a"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   765
      TabIndex        =   66
      Top             =   495
      Width           =   2070
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmMain.frx":3744
      Height          =   615
      Left            =   2895
      TabIndex        =   39
      Top             =   60
      Width           =   8595
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "v1.0.1ß"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   1890
      TabIndex        =   38
      Top             =   105
      Width           =   1005
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Burning-iX"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   450
      Left            =   180
      TabIndex        =   37
      Top             =   15
      Width           =   1935
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "v1.0.1ß"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   300
      Left            =   1875
      TabIndex        =   36
      Top             =   135
      Width           =   1005
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Burning-iX"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   480
      Left            =   135
      TabIndex        =   35
      Top             =   30
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   -15
      Picture         =   "frmMain.frx":3821
      Top             =   120
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000003&
      Height          =   765
      Left            =   -30
      Top             =   -15
      Width           =   11625
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' /*
' Burning-iX v1.0.1ß (beta) 2009 © Salvo Cortesiano.
' See the SDK of NeroCOM for more info
' \*

Option Explicit

' service variables
Dim strLastPath As String
Dim FileLogOfNeroCom As String
Dim MyString As String
Dim BurnError As Boolean
Dim ssLeft As Long
Dim ssTop As Long

' enum to strip the path file
Private Enum Extract
  [Only_Extension] = 0
  [Only_FileName_and_Extension] = 1
  [Only_FileName_no_Extension] = 2
  [Only_Path] = 3
End Enum

' time Burning
Dim Hours As Integer
Dim Minutes As Integer
Dim Seconds As Integer
Dim Days As Integer
Dim AddMinutes As Boolean
Dim AddHours As Boolean
Dim AddDays As Boolean

' variables Nero
Dim Source_Dir As String
Dim FSO As New FileSystemObject
Dim DateFolder As NeroFolder
Dim sFile As NeroFile
Dim rootfolder As NeroFolder

' burn/cretae ISO Image
Dim sISOFile As String

' load NERO references
Public WithEvents Nero As Nero
Attribute Nero.VB_VarHelpID = -1
Public WithEvents Drive As NeroDrive
Attribute Drive.VB_VarHelpID = -1

' drive writable?
Dim IsDriveWriteable As Boolean

' media type
Dim DriveMediaType As String

' variable for holding number of existing sessions on disc when cd info read
Dim NumExistingTracks As Integer

'flag for checking if drive event finished
Dim DriveFinished As Boolean

' list of available drives
Dim Drives As NeroDrives 'INeroDrive2
Dim CancelPressed As Boolean

' main folder to be burnt
Dim fOlder As NeroFolder
Dim ISOTrack As NeroISOTrack
Dim CDStamp As NeroCDStamp

' audio track/s
Dim AudioTracks As NeroAudioTracks
Dim AudioTrack As NeroAudioTrack

Dim gSessionNumber As Integer

' .... Play Sound Resource
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function PlaySoundData Lib "winmm.dll" Alias "PlaySoundA" (lpData As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long

Private Const SND_SYNC = &H0
Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2
Private Const SND_MEMORY = &H4
Private Const SND_ALIAS = &H10000
Private Const SND_FILENAME = &H20000
Private Const SND_RESOURCE = &H40004
Private Const SND_ALIAS_ID = &H110000
Private Const SND_ALIAS_START = 0
Private Const SND_LOOP = &H8
Private Const SND_NOSTOP = &H10
Private Const SND_VALID = &H1F
Private Const SND_NOWAIT = &H2000
Private Const SND_VALIDFLAGS = &H17201F
Private Const SND_RESERVED = &HFF000000
Private Const SND_TYPE_MASK = &H170007

Private Const WAVERR_BASE = 32
Private Const WAVERR_BADFORMAT = (WAVERR_BASE + 0)
Private Const WAVERR_STILLPLAYING = (WAVERR_BASE + 1)
Private Const WAVERR_UNPREPARED = (WAVERR_BASE + 2)
Private Const WAVERR_SYNC = (WAVERR_BASE + 3)
Private Const WAVERR_LASTERROR = (WAVERR_BASE + 3)

Private m_snd() As Byte

Private Sub CheckEraseDisk_Click()
    If CheckEraseDisk.value = 1 Then cmbErasing.Enabled = True Else cmbErasing.Enabled = False
End Sub

Private Sub CheckOnlyISO_Click()
    If LCase$(lst_AvailableDevices.Text) <> "image recorder" Then CheckOnlyISO.value = 0 Else CheckOnlyISO.value = 1
End Sub

Private Sub cmdAbort_Click()
    Dim Response As Long
    Response = MsgBox("Abort May Cause CD To Become Non-Read/Writeable! Abort Anyway?", vbYesNo + vbExclamation)
    If Response = vbYes Then
    Nero.Abort
        CancelPressed = True
        AddMessage "Abort Pressed!"
        cmdAbort.Enabled = False
        TBurning.Enabled = False
        TBurning.Interval = 0
        '/*----------------------------
        Call ResetAll
    End If
End Sub

Private Sub cmdBrowse_Click()
    
End Sub

Private Sub cmdBurFile_Click()
    Dim Source_Dir As String
    Dim X As Boolean
    Dim i As Integer
    Dim sWriteSpeed As Long
    
    On Local Error GoTo Exit_Me
    
    If strLastPath = "" Then strLastPath = App.path + "\"
            Source_Dir = BrowseFolder("Select folder path:", strLastPath)
        If Source_Dir <> "" And Source_Dir <> "Error!" Then
        Source_Dir = Source_Dir + "\"
    Else
        Exit Sub
    End If
    
    sWriteSpeed = cmbWSpeed.List(cmbWSpeed.ListIndex)
    
    cmdAbort.Enabled = True
    cmdBurFile.Enabled = False
    cmdEject.Enabled = False
    cmdExit.Enabled = False
    cmdBurning.Enabled = False
    
    pgs_Burn.value = 0
    pgs_Buffer.value = 0
    fme_Progress.value = 0
    lblPercentuale.Caption = "00%"
    lblTCriting.Caption = "00%"
    lblTWriting.Caption = "00%"
    
    CancelPressed = False
    BurnError = False
    
    lst_Messages.Text = ""
    
    Set fOlder = New NeroFolder
    Set Drive = Drives(lst_AvailableDevices.ListIndex)
    
    Call StartCount

    CancelPressed = False

' erase the Disk
If CheckEraseDisk.value = 1 Then
    AddMessage "Waiting For Erase CD/DVD..."
    DriveFinished = False
    If cmbErasing.Enabled = True And cmbErasing.ListIndex = 0 Then
        AddMessage "Erase mode Quick!"
        Drive.EraseDisc True, NERO_ERASE_MODE_DEFAULT + NERO_ERASE_MODE_DISABLE_EJECT
    ElseIf cmbErasing.Enabled = True And cmbErasing.ListIndex = 1 Then
        AddMessage "Erase mode Complete!"
        Drive.EraseDisc False, NERO_ERASE_MODE_DEFAULT + NERO_ERASE_MODE_DISABLE_EJECT
    End If
End If

If CheckEraseDisk.value = 1 Then
' wait for event done and handled
    While Not DriveFinished
If CancelPressed Then
        GoTo Exit_Me
    End If
        X = DoEvents()
    Wend
End If

'check if multisession data
AddMessage "Checking CD/DVD for existing Data..."
DriveFinished = False
Drive.CDInfo NERO_READ_ISRC

' wait for event done and handled
While Not DriveFinished
If CancelPressed Then
GoTo Exit_Me
End If
X = DoEvents()
Wend

' Not existing session
If NumExistingTracks < 0 Then
' no disk ready
    AddMessage "The CD/DVD not contains any session... Exit!"
GoTo Exit_Me
End If

'if existing session then import the last one
If NumExistingTracks > 0 Then
AddMessage "Reading existing Data from CD/DVD..."

'read in the last session
i = NumExistingTracks - 1
DriveFinished = False

Drive.ImportIsoTrack i, NERO_IMPORT_ISO_ONLY

' wait for event done and handled
While Not DriveFinished
If CancelPressed Then
GoTo Exit_Me
End If
X = DoEvents()
Wend

End If

    Set DateFolder = New NeroFolder
    Set ISOTrack = New NeroISOTrack
    
    AddMessage "Coping Folders and Files... Please wait..."
    DateFolder.Name = txtSession.Text + "-" + Format(Now, "dd") + "-" + Format(Now, "mm") + "-" + Format(Now, "yyyy") + " - " + Format(Now, "hh") + "-" + Format(Now, "nn") + "-" + Format(Now, "ss")
    
    ' Add to folder tree
    fOlder.Folders.Add DateFolder
    
    ' recursively build folder tree
    Call BuildFileFolderTree(DateFolder, FSO.GetFolder(Source_Dir))
    
    If Not FSO.FolderExists(Source_Dir) Then
            MsgBox "Error - Source Folder does not Exist!", vbCritical, App.Title
        GoTo Exit_Me
    End If
    
    If CancelPressed Then GoTo Exit_Me
    
    ISOTrack.Name = ISOTrackName.Text
    ISOTrack.rootfolder = fOlder
    ISOTrack.BurnOptions = NERO_BURN_OPTION_CREATE_ISO_FS + NERO_BURN_OPTION_USE_JOLIET
    
    ' burn folder (check if underrun protection available and use if it is)
    DriveFinished = False
    
    Drive.BurnIsoAudioCD ISOTrackName.Text, txtSession.Text, 0, ISOTrack, Nothing, Nothing, NERO_BURN_FLAG_BUF_UNDERRUN_PROT + _
    NERO_BURN_FLAG_WRITE, cmbWSpeed.List(cmbWSpeed.ListIndex), NERO_MEDIA_DVD_M_R + NERO_MEDIA_DVD_P_R
    
    While Not DriveFinished
        If CancelPressed Then
            GoTo Exit_Me
        End If
    X = DoEvents()
    Wend
    
    If BurnError = False And CancelPressed = False Then
        MsgBox "Burning finished success:" & vbCr & "Total time: " & vbCr & Label8.Caption & "." _
        & vbCr & vbCr & "See the file Log for more info!", vbInformation, App.Title
    End If
    
Call ResetAll
Exit Sub
Exit_Me:
    AddMessage Error$
    AddMessage Nero.LastError
    If Err.Number <> 0 Then
        Call WriteErrorLogs(Err.Number, Err.Description, Nero.LastError & vbCr & "See the Log file for more info!", True, True)
    End If
Call ResetAll
End Sub

Private Sub cmdBurnImage_Click()
    Dim X As Long
    Dim sWriteSpeed As Long
    
    On Local Error GoTo Exit_Me
    
    cmdAbort.Enabled = True
    cmdExit.Enabled = False
    cmdBurnImage.Enabled = False
    cmdBurFile.Enabled = False
    
    lst_Messages.Text = ""
    BurnError = False
    sWriteSpeed = cmbWSpeed.List(cmbWSpeed.ListIndex)
    Set Drive = Drives(lst_AvailableDevices.ListIndex)
    
' init Source_Dir
    Source_Dir = App.path + "\mybackup\bend"
    
    Set fOlder = New NeroFolder
    
' ... to the option selected
    If opt_B(1).value Then
        Source_Dir = Source_Dir + "\ybc"
    ElseIf opt_B(2).value Then
        Source_Dir = Source_Dir + "\log"
    ElseIf opt_B(4).value Then
        Source_Dir = Source_Dir + "\backup"
    ElseIf opt_B(0).value Then
        Source_Dir = Source_Dir + "\ISO"
    End If

    If Not FSO.FolderExists(Source_Dir) Then
            MsgBox "Error - Source Folder does not Exist!", vbCritical, App.Title
        GoTo Exit_Me
    End If

    Call StartCount
    
    Set DateFolder = New NeroFolder
    Set ISOTrack = New NeroISOTrack
    
' Set the Folder
    If opt_B(3).value Then
        DateFolder.Name = "Bend " + Format(Now, "dd") + "-" + Format(Now, "mm") + "-" + Format(Now, "yyyy") + " - " + Format(Now, "hh") + "-" + Format(Now, "nn") + "-" + Format(Now, "ss")
    ElseIf opt_B(1).value Then
        DateFolder.Name = "YBC " + Format(Now, "dd") + "-" + Format(Now, "mm") + "-" + Format(Now, "yyyy") + " - " + Format(Now, "hh") + "-" + Format(Now, "nn") + "-" + Format(Now, "ss")
    ElseIf opt_B(2).value Then
        DateFolder.Name = "Log " + Format(Now, "dd") + "-" + Format(Now, "mm") + "-" + Format(Now, "yyyy") + " - " + Format(Now, "hh") + "-" + Format(Now, "nn") + "-" + Format(Now, "ss")
    ElseIf opt_B(4).value Then
        DateFolder.Name = "Backup " + Format(Now, "dd") + "-" + Format(Now, "mm") + "-" + Format(Now, "yyyy") + " - " + Format(Now, "hh") + "-" + Format(Now, "nn") + "-" + Format(Now, "ss")
    ElseIf opt_B(0).value Then
        DateFolder.Name = "ISO " + Format(Now, "dd") + "-" + Format(Now, "mm") + "-" + Format(Now, "yyyy") + " - " + Format(Now, "hh") + "-" + Format(Now, "nn") + "-" + Format(Now, "ss")
    End If

' Add to folder tree
    fOlder.Folders.Add DateFolder

' recursively build folder tree
    Call BuildFileFolderTree(DateFolder, FSO.GetFolder(Source_Dir))

    ISOTrack.Name = ISOTrackName.Text
    ISOTrack.rootfolder = fOlder
    
    CancelPressed = False
    If CancelPressed Then GoTo Exit_Me

' burn folder to ISO (check if underrun protection available and use if it is)
    DriveFinished = False
        
' start routine ISO
    If CheckOnlyISO.value = 1 Then
        Drive.InitImageRecorder sISOFile, NERO_MEDIA_NONE
        ISOTrack.BurnOptions = NERO_BURN_OPTION_CREATE_ISO_FS + NERO_BURN_OPTION_CREATE_UDF_FS + NERO_BURN_OPTION_USE_JOLIET
        Drive.BurnIsoAudioCD "", "", False, ISOTrack, Nothing, Nothing, NERO_BURN_FLAG_WRITE + NERO_BURN_FLAG_BUF_UNDERRUN_PROT _
        + NERO_BURN_FLAG_CLOSE_SESSION, sWriteSpeed, NERO_MEDIA_NONE

' wait for event done and handled
    While Not DriveFinished
If CancelPressed Then
        GoTo Exit_Me
    End If
        X = DoEvents()
    Wend
End If
    'GetFilePath
    If BurnError = False And CancelPressed = False Then
            MsgBox "Image ISO created success!" & vbCr _
        & vbCr & "FileName: " & GetFilePath(sISOFile, Only_FileName_and_Extension) & vbCr _
        & "Path: " & GetFilePath(sISOFile, Only_Path) & vbCr & "File size: " & GetFileSize(sISOFile), vbInformation, App.Title
    End If
    Call ResetAll
Exit Sub
Exit_Me:
    AddMessage Error$
    AddMessage Nero.LastError
    If Err.Number <> 0 Then
        Call WriteErrorLogs(Err.Number, Err.Description, Nero.LastError & vbCr & "See the Log file for more info!", True, True)
    End If
    Call ResetAll
    Err.Clear
End Sub

Private Sub cmdBurning_Click()
Dim i As Integer
Dim X As Long
Dim j As Long
Dim sWriteSpeed As Long

On Error GoTo Exit_Me:

CancelPressed = False
cmdAbort.Enabled = True
cmdBurning.Enabled = False
cmdEject.Enabled = False
cmdExit.Enabled = False

BurnError = False

lst_Messages.Text = ""
Me.Refresh

sWriteSpeed = cmbWSpeed.List(cmbWSpeed.ListIndex)

Set Drive = Drives(lst_AvailableDevices.ListIndex)

Source_Dir = App.path + "\mybackup\bend"

If opt_B(1).value Then
    Source_Dir = Source_Dir + "\ybc"
ElseIf opt_B(2).value Then
    Source_Dir = Source_Dir + "\log"
ElseIf opt_B(4).value Then
    Source_Dir = Source_Dir + "\backup"
End If

If Not FSO.FolderExists(Source_Dir) Then
        MsgBox "Error - Source Folder does not Exist!", vbCritical, App.Title
    GoTo Exit_Me
End If

Call StartCount

' erase the Disk
If CheckEraseDisk.value = 1 Then
    AddMessage "Waiting For Erase CD/DVD..."
    DriveFinished = False
    If cmbErasing.Enabled = True And cmbErasing.ListIndex = 0 Then
        AddMessage "Erase mode Quick!"
        Drive.EraseDisc True, NERO_ERASE_MODE_DEFAULT + NERO_ERASE_MODE_DISABLE_EJECT
    ElseIf cmbErasing.Enabled = True And cmbErasing.ListIndex = 1 Then
        AddMessage "Erase mode Complete!"
        Drive.EraseDisc False, NERO_ERASE_MODE_DEFAULT + NERO_ERASE_MODE_DISABLE_EJECT
    End If
End If

If CheckEraseDisk.value = 1 Then
' wait for event done and handled
    While Not DriveFinished
If CancelPressed Then
        GoTo Exit_Me
    End If
        X = DoEvents()
    Wend
End If

Set fOlder = New NeroFolder

'check if multisession data
AddMessage "Checking CD/DVD for existing Data..."
DriveFinished = False
Drive.CDInfo NERO_READ_ISRC

' wait for event done and handled
While Not DriveFinished
If CancelPressed Then
GoTo Exit_Me
End If
X = DoEvents()
Wend

' Not existing session
If NumExistingTracks < 0 Then
' no disk ready
    AddMessage "The CD/DVD not contains any session... Exit!"
GoTo Exit_Me
End If

'if existing session then import the last one
If NumExistingTracks > 0 Then
AddMessage "Reading existing Data from CD/DVD..."

'read in the last session
i = NumExistingTracks - 1
DriveFinished = False

Drive.ImportIsoTrack i, NERO_IMPORT_ISO_ONLY

' wait for event done and handled
While Not DriveFinished
If CancelPressed Then
GoTo Exit_Me
End If
X = DoEvents()
Wend

End If

Set DateFolder = New NeroFolder
Set ISOTrack = New NeroISOTrack

' Set the Folder
If opt_B(3).value Then
    DateFolder.Name = "Bend " + Format(Now, "dd") + "-" + Format(Now, "mm") + "-" + Format(Now, "yyyy") + " - " + Format(Now, "hh") + "-" + Format(Now, "nn") + "-" + Format(Now, "ss")
ElseIf opt_B(1).value Then
    DateFolder.Name = "YBC " + Format(Now, "dd") + "-" + Format(Now, "mm") + "-" + Format(Now, "yyyy") + " - " + Format(Now, "hh") + "-" + Format(Now, "nn") + "-" + Format(Now, "ss")
ElseIf opt_B(2).value Then
    DateFolder.Name = "Log " + Format(Now, "dd") + "-" + Format(Now, "mm") + "-" + Format(Now, "yyyy") + " - " + Format(Now, "hh") + "-" + Format(Now, "nn") + "-" + Format(Now, "ss")
ElseIf opt_B(4).value Then
    DateFolder.Name = "Backup " + Format(Now, "dd") + "-" + Format(Now, "mm") + "-" + Format(Now, "yyyy") + " - " + Format(Now, "hh") + "-" + Format(Now, "nn") + "-" + Format(Now, "ss")
ElseIf opt_B(0).value Then
    DateFolder.Name = "ISO " + Format(Now, "dd") + "-" + Format(Now, "mm") + "-" + Format(Now, "yyyy") + " - " + Format(Now, "hh") + "-" + Format(Now, "nn") + "-" + Format(Now, "ss")
End If

' Add to folder tree
fOlder.Folders.Add DateFolder

' recursively build folder tree
Call BuildFileFolderTree(DateFolder, FSO.GetFolder(Source_Dir))

ISOTrack.Name = ISOTrackName.Text
ISOTrack.rootfolder = fOlder

CancelPressed = False

If CancelPressed Then GoTo Exit_Me

' burn folder (check if underrun protection available and use if it is)
DriveFinished = False

' option file system
Select Case cmbSystem.ListIndex
    Case 0
        ISOTrack.BurnOptions = NERO_BURN_OPTION_CREATE_ISO_FS + NERO_BURN_OPTION_USE_JOLIET
    Case 1
        ISOTrack.BurnOptions = NERO_BURN_OPTION_CREATE_ISO_FS + NERO_BURN_OPTION_CREATE_UDF_FS + NERO_BURN_OPTION_USE_JOLIET
    Case 2
        ISOTrack.BurnOptions = NERO_BURN_OPTION_CREATE_ISO_FS + NERO_BURN_OPTION_USE_MODE2
End Select

' option type
'/* TODO
' Now not implemented
Select Case cmbType.ListIndex
    Case 0
        
    Case 1
        
    Case 2
        'Drive.BurnIsoAudioCD ISOTrackName.Text, txtSession.Text, True, ISOTrack, Nothing, Nothing, nerolib.NERO_BURN_FLAG_DAO _
        , sWriteSpeed, NERO_MEDIA_DVD_M_R + NERO_MEDIA_DVD_P_R
End Select

' .... burn folder (check if underrun protection available and use if it is)
If Drive.Capabilities And NERO_CAP_BUF_UNDERRUN_PROT Then
    If cmbTypeSupport.ListIndex = 0 Then
        If CheckSimulate.value = 1 Then
            Drive.BurnIsoAudioCD "", "", 0, ISOTrack, Nothing, Nothing, NERO_BURN_FLAG_SIMULATE + NERO_BURN_FLAG_BUF_UNDERRUN_PROT + NERO_BURN_FLAG_WRITE, sWriteSpeed, NERO_MEDIA_DVD_M_R + NERO_MEDIA_DVD_P_R
        Else
            Drive.BurnIsoAudioCD ISOTrackName.Text, txtSession.Text, 0, ISOTrack, Nothing, Nothing, NERO_BURN_FLAG_WRITE + _
            NERO_BURN_FLAG_BUF_UNDERRUN_PROT + NERO_BURN_FLAG_CLOSE_SESSION, sWriteSpeed, nerolib.NERO_MEDIA_TYPE.NERO_MEDIA_DVD_M_R _
            + nerolib.NERO_MEDIA_TYPE.NERO_MEDIA_DVD_P_R + nerolib.NERO_ERASE_MODE_DISABLE_EJECT
        End If
    ElseIf cmbTypeSupport.ListIndex = 1 Then
        If CheckSimulate.value = 1 Then
            Drive.BurnIsoAudioCD "", "", 0, ISOTrack, Nothing, Nothing, NERO_BURN_FLAG_SIMULATE + NERO_BURN_FLAG_WRITE, sWriteSpeed, nerolib.NERO_MEDIA_CDRW
        Else
            Drive.BurnIsoAudioCD ISOTrackName.Text, txtSession.Text, 0, ISOTrack, Nothing, Nothing, NERO_BURN_FLAG_WRITE + _
            NERO_BURN_FLAG_CLOSE_SESSION, sWriteSpeed, nerolib.NERO_MEDIA_CDRW
        End If
    End If
Else
    If cmbTypeSupport.ListIndex = 0 Then
        If CheckSimulate.value = 1 Then
            Drive.BurnIsoAudioCD "", "", 0, ISOTrack, Nothing, Nothing, NERO_BURN_FLAG_SIMULATE + NERO_BURN_FLAG_WRITE, sWriteSpeed, NERO_MEDIA_DVD_M_R + NERO_MEDIA_DVD_P_R
        Else
            Drive.BurnIsoAudioCD ISOTrackName.Text, txtSession.Text, 0, ISOTrack, Nothing, Nothing, NERO_BURN_FLAG_WRITE + _
            NERO_BURN_FLAG_CLOSE_SESSION, sWriteSpeed, nerolib.NERO_MEDIA_TYPE.NERO_MEDIA_DVD_M_R + _
            nerolib.NERO_MEDIA_TYPE.NERO_MEDIA_DVD_P_R
        End If
    ElseIf cmbTypeSupport.ListIndex = 1 Then
        If CheckSimulate.value = 1 Then
            Drive.BurnIsoAudioCD "", "", 0, ISOTrack, Nothing, Nothing, NERO_BURN_FLAG_SIMULATE + NERO_BURN_FLAG_WRITE, sWriteSpeed, nerolib.NERO_MEDIA_CDRW
        Else
            Drive.BurnIsoAudioCD ISOTrackName.Text, txtSession.Text, 0, ISOTrack, Nothing, Nothing, NERO_BURN_FLAG_WRITE + _
            NERO_BURN_FLAG_CLOSE_SESSION, sWriteSpeed, nerolib.NERO_MEDIA_CDRW
        End If
    End If
End If

While Not DriveFinished
If CancelPressed Then
    GoTo Exit_Me
End If
    X = DoEvents()
Wend
    If BurnError = False And CancelPressed = False Then
        MsgBox "Burning finished success:" & vbCr & "Total time: " & vbCr & Label8.Caption & "." _
        & vbCr & vbCr & "See the file Log for more info!", vbInformation, App.Title
    End If
    Call ResetAll
Exit Sub

Exit_Me:
    AddMessage Error$
    AddMessage Nero.LastError
    If Err.Number <> 0 Then
        Call WriteErrorLogs(Err.Number, Err.Description, Nero.LastError & vbCr & "See the Log file for more info!", True, True)
    End If
    Call ResetAll
End Sub
Private Sub cmdEject_Click()
    Set Drive = Drives(lst_AvailableDevices.ListIndex)
    Drive.EjectCD
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub


Private Sub cmdPath_Click()
    Dim ISOF As String
    Dim IDir As String
    Dim X As Long
    Dim sWriteSpeed As Integer
    
    On Error GoTo Exit_Me:
    
    With cDialog
        IDir = INI.GetKeyValue("SETTING", "LAST PATH SELECTED")
        If IDir = "" Then .InitDir = App.path Else IDir = GetFilePath(INI.GetKeyValue("SETTING", "LAST PATH SELECTED"), Only_Path)
        .CancelError = True
        .Filter = "Supported ISO Files (*.iso)|*.iso|All Supported Image (*.nrg;*.cue;*.iso)|*.nrg;*.cue;*.iso"
        .DialogTitle = "Choose ISO image!"
        .DefaultExt = ".iso"
        .FilterIndex = 1
        .InitDir = IDir
        .ShowOpen
        If .FileName = "" Then Exit Sub
        ISOF = .FileName
    End With
    
    txtPath.Text = ISOF
    lblISOFileSize.Caption = "Size: " & GetFileSize(cDialog.FileName)
    ISOTrackName.Text = UCase$(GetFilePath(ISOF, Only_FileName_no_Extension))
    txtSession.Text = UCase$(GetFilePath(ISOF, Only_FileName_no_Extension))
    ISOTrackName.Text = UCase$(Replace(ISOTrackName.Text, " ", "_"))
    txtSession.Text = UCase$(Replace(txtSession.Text, " ", "_"))
    ' save to INI last file ISO selected
    INI.DeleteKey "SETTING", "LAST PATH SELECTED"
    INI.CreateKeyValue "SETTING", "LAST PATH SELECTED", txtPath.Text
    strLastPath = GetFilePath(txtPath.Text, Only_Path) + "\"
    
    ' -------------------------------------------------------------- START BURN ISO FILE
    
    Set sFile = New NeroFile
    Set ISOTrack = New NeroISOTrack
    Set fOlder = New NeroFolder
    
    CancelPressed = False
    cmdAbort.Enabled = True
    cmdBurning.Enabled = False
    cmdEject.Enabled = False
    cmdExit.Enabled = False

    BurnError = False

    lst_Messages.Text = ""
    Me.Refresh

    sWriteSpeed = cmbWSpeed.List(cmbWSpeed.ListIndex)

    Set Drive = Drives(lst_AvailableDevices.ListIndex)
    
    DoEvents
    ISOTrack.Name = txtSession.Text
    ISOTrack.rootfolder = fOlder
    
    sFile.Name = GetFilePath(txtPath.Text, Only_FileName_and_Extension)
    sFile.SourceFilePath = txtPath.Text
    fOlder.Files.Add sFile
    
    DriveFinished = False
    
    Call StartCount
    
    ' erase the Disk
If CheckEraseDisk.value = 1 Then
    AddMessage "Waiting For Erase CD/DVD..."
    DriveFinished = False
    If cmbErasing.Enabled = True And cmbErasing.ListIndex = 0 Then
        AddMessage "Erase mode Quick!"
        Drive.EraseDisc True, NERO_ERASE_MODE_DEFAULT + NERO_ERASE_MODE_DISABLE_EJECT
    ElseIf cmbErasing.Enabled = True And cmbErasing.ListIndex = 1 Then
        AddMessage "Erase mode Complete!"
        Drive.EraseDisc False, NERO_ERASE_MODE_DEFAULT + NERO_ERASE_MODE_DISABLE_EJECT
    End If
End If

If CheckEraseDisk.value = 1 Then
' wait for event done and handled
    While Not DriveFinished
If CancelPressed Then
        GoTo Exit_Me
    End If
        X = DoEvents()
    Wend
End If
    
    DriveFinished = False
    CancelPressed = False
    
    ISOTrack.BurnOptions = NERO_BURN_OPTION_CREATE_ISO_FS + NERO_BURN_OPTION_CREATE_UDF_FS + NERO_BURN_OPTION_USE_JOLIET
    
    Drive.BurnImage sFile.SourceFilePath, nerolib.NERO_BURN_FLAG_WRITE + nerolib.NERO_MEDIA_CDRW _
    + nerolib.NERO_MEDIA_DVD_M + nerolib.NERO_MEDIA_DVD_P + NERO_BURN_FLAG_DETECT_NON_EMPTY_CDRW + _
    NERO_BURN_FLAG_BUF_UNDERRUN_PROT + NERO_BURN_FLAG_CLOSE_SESSION, sWriteSpeed
    
    While Not DriveFinished
If CancelPressed Then
    GoTo Exit_Me
End If
    X = DoEvents()
Wend
    
    If BurnError = False And CancelPressed = False Then
        MsgBox "Burning finished success:" & vbCr & "Total time: " & vbCr & Label8.Caption & "." _
        & vbCr & vbCr & "See the file Log for more info!", vbInformation, App.Title
    End If
    Call ResetAll
Exit Sub
Exit_Me:
    AddMessage Error$
    AddMessage Nero.LastError
    If Err.Number <> 0 Then
        Call WriteErrorLogs(Err.Number, Err.Description, Nero.LastError & vbCr & "See the Log file for more info!", True, True)
    End If
    Call ResetAll
End Sub

' Abort = False
Private Sub Drive_OnAborted(Abort As Boolean)
    Abort = False
End Sub

' .... Add Working
Private Sub Drive_OnAddLogLine(TextType As nerolib.NERO_TEXT_TYPE, Text As String)
    SplitText Text
End Sub


Private Sub Drive_OnDoneBAOCloseHandle(ByVal StatusCode As nerolib.NERO_BURN_ERROR)
    SplitText StatusCode
End Sub

Private Sub Drive_OnDoneBAOCreateHandle(ByVal StatusCode As nerolib.NERO_BURN_ERROR)
    SplitText StatusCode
End Sub


Private Sub Drive_OnDoneBAOWriteToFile(ByVal StatusCode As nerolib.NERO_BURN_ERROR, ByVal lNumberOfBytesWritten As Long)
    On Local Error Resume Next
    AddMessage "WriteToFile " & StatusCode
    AddMessage "Number of bytes Written " & lNumberOfBytesWritten
End Sub

' Event for burn complete prints results to message list
Private Sub Drive_OnDoneBurn(StatusCode As nerolib.NERO_BURN_ERROR)
    SplitText Nero.ErrorLog
    SplitText Nero.LastError
    If StatusCode <> nerolib.NERO_BURN_OK Then
            AddMessage "Burn not finished successfully: " & StatusCode
        Call PlaySoundResource(101)
            MsgBox "Burn not finished successfully:" & vbCr & StatusCode _
        & vbCr & vbCr & "See the file Log for more info!", vbCritical, App.Title
        BurnError = True
        Call ResetAll
    Else
        AddMessage "Burn finished successfully: " & StatusCode
        Call PlaySoundResource(102)
        BurnError = False
        Call ResetAll
    End If
    DriveFinished = True
    Me.Caption = "Burning-iX v1.0.1ß (beta) 2009 © Salvo Cortesiano."
End Sub

' Event for read cd info done
Private Sub Drive_OnDoneCDInfo(ByVal pCDInfo As nerolib.INeroCDInfo)
    'set number of existing sessions
    On Local Error GoTo NoTracks:
    NumExistingTracks = pCDInfo.Tracks.Count
    IsDriveWriteable = pCDInfo.IsWriteable
    DriveMediaType = pCDInfo.MediaType

    'set done flag
    DriveFinished = True
    Exit Sub
NoTracks:
    NumExistingTracks = 0
    DriveFinished = True
End Sub


Private Sub Drive_OnDoneErase(Ok As Boolean)
On Local Error GoTo backError
    If Ok Then
        AddMessage "Disc Erase Successful!"
    Else
        AddMessage "Disc Erase Failed!"
    End If
    DriveFinished = True
Exit Sub
backError:
        'set done flag
        DriveFinished = True
    Err.Clear
End Sub

' get the track size in bytes?
Private Sub Drive_OnDoneEstimateTrackSize(ByVal bOk As Boolean, ByVal BlockSize As Long)
    On Local Error Resume Next
    If bOk Then
        AddMessage "Track Size: " & BlockSize
    Else
        AddMessage "Error to determine the Track Size!"
    End If
    'set done flag
    'DriveFinished = True
End Sub

' mmm i tink to remove this? ...
Private Sub drive_OnDoneImport(Ok As Boolean, fOlder As nerolib.INeroFolder, CDStamp As nerolib.INeroCDStamp)
    On Local Error Resume Next
    If Ok Then
        Set rootfolder = fOlder
        AddMessage "Previous Session Imported!"
    Else
        AddMessage "Error Importing Session!"
    End If
    ' set done flag
    DriveFinished = True
End Sub

' Importing of data done event
Private Sub Drive_OnDoneImport2(ByVal bOk As Boolean, ByVal pFolder As nerolib.INeroFolder, ByVal pCDStamp As nerolib.INeroCDStamp, ByVal pImportInfo As nerolib.INeroImportDataTrackInfo, ByVal importResult As nerolib.NERO_IMPORT_DATA_TRACK_RESULT)
    Dim i As Integer
        If bOk Then
            Set fOlder = pFolder
        Else
                MsgBox "Error Reading In Data!", vbCritical, App.Title
            AddMessage "Error Reading In Data!"
        End If
    ' set done flag
    DriveFinished = True
    
    ' I'm truble :)
    
    ' Dim i As Integer
    ' Dim tempfile As NeroFile
    
    ' If Ok Then
        ' For i = 0 To (pFolder.Files.Count - 1) Step 1
            ' tempfile = New NeroFile
            ' tempfile.Name = pFolder.Files.Item(i).Name
            ' tempfile.EntryTime = pFolder.Files.Item(i).EntryTime
            ' tempfile.SourceFilePath = pFolder.Files.Item(i).SourceFilePath
            ' rootfolder.Files.Add(tempfile)
            ' tempfile = Nothing
        ' Next i
    ' End If

    ' AddMessage "Previous Session Imported!"
    ' DriveFinished = True
End Sub

' Dispaly finish info...
Private Sub Drive_OnDoneWaitForMedia(Success As Boolean)
    AddMessage "Done waiting for media=" & Success
End Sub


Private Sub Drive_OnDriveStatusChanged(ByVal driveStatus As nerolib.NERO_DRIVESTATUS_RESULT)
    On Local Error Resume Next
    AddMessage "Drive " & driveStatus
End Sub

Private Sub Drive_OnMajorPhase(phase As nerolib.NERO_MAJOR_PHASE)
    SplitText phase
End Sub

' Display the Work...
Private Sub Drive_OnProgress(ProgressInPercent As Long, Abort As Boolean)
    Abort = False
    pgs_Burn.value = ProgressInPercent
    lblTWriting.Caption = Format(ProgressInPercent, "00") & "%"
    If Me.WindowState = vbMinimized Then Me.Caption = "Burning " & Format(ProgressInPercent, "00") _
    & "%" Else Me.Caption = "Burning-iX v1.0.1ß (beta) 2009 © Salvo Cortesiano."
    On Local Error Resume Next
    pgs_Buffer.value = ProgressInPercent / 2
    lblPercentuale.Caption = Format(ProgressInPercent / 2, "00") & "%"
End Sub


Private Sub Drive_OnRoboPrintLabel(pbSuccess As Boolean)
    If pbSuccess Then
        AddMessage "Print Label success..."
    Else
        AddMessage "Error to Print Label!"
    End If
    'set done flag
    'DriveFinished = True
End Sub

' Add the Line...
Private Sub Drive_OnSetPhase(Text As String)
    SplitText Text
End Sub


' Display the SubTask Work...
Private Sub Drive_OnSubTaskProgress(ProgressInPercent As Long, Abort As Boolean)
    Abort = False
    On Local Error Resume Next
    fme_Progress.value = ProgressInPercent
    lblTCriting.Caption = Format(fme_Progress.value, "00") & "%"
    ' until crash :)
    If Format(fme_Progress.value, "00") >= 100 Then fme_Progress.value = 80
    pgs_Buffer.value = ProgressInPercent / 2
    lblPercentuale.Caption = Format(ProgressInPercent / 2, "00") & "%"
End Sub


Private Sub Drive_OnWriteDAE(ignore As Long, Data As Variant)
     AddMessage "WriteDAE: " & Data
End Sub

Private Sub AddMessage(ByVal Message As String)
    On Local Error Resume Next
    '/* USE TEXTBOX
    lst_Messages.Text = lst_Messages.Text + Message + Chr$(13) + Chr$(10)
    If CheckLog.value = 1 Then WriteLog lst_Messages.Text + Message + Chr$(13) + Chr$(10)
    lst_Messages.SelStart = Len(lst_Messages.Text)
    '/* OR USE LISTBOX ;)
    'lst_Messages.AddItem Message
    '    If lst_Messages.ListCount <> 0 Then
    '        lst_Messages.ListIndex = lst_Messages.ListCount - 1
    '    lst_Messages.Refresh
    'End If
End Sub

' Function for removing extra spaces and lines from messages
Private Sub SplitText(ByVal Data As String)
    Dim Temp As String
    Dim i As Integer
    Temp = ""
        For i = 1 To Len(Data)
            If Mid$(Data, i, 1) = Chr$(13) Then
                'lst_Messages.AddItem Trim$(Temp)
                AddMessage Trim$(Temp)
                Temp = ""
            ElseIf Mid$(Data, i, 1) <> Chr$(10) Then
                Temp = Temp + Mid$(Data, i, 1)
            End If
        Next
    If Temp <> "" Then AddMessage Trim$(Temp)
End Sub

Private Sub BuildFileFolderTree(ByRef nroFolderToUse As NeroFolder, ByRef folCurrent As fOlder)
    Dim folTMP As fOlder
    Dim filTMP As File
    Dim nroFolTmp As NeroFolder
    Dim nroFilTmp As NeroFile

    'Add all files in the current directory
    For Each filTMP In folCurrent.Files
        Set nroFilTmp = New NeroFile
        nroFilTmp.Name = filTMP.Name
        nroFilTmp.SourceFilePath = filTMP.path
        nroFolderToUse.Files.Add nroFilTmp
    Next
    'Write the sub folders
    For Each folTMP In folCurrent.SubFolders
        Set nroFolTmp = New NeroFolder
        nroFolTmp.Name = folTMP.Name
        nroFolderToUse.Folders.Add nroFolTmp
        Call BuildFileFolderTree(nroFolTmp, folTMP)
    Next
End Sub

Private Sub Form_Initialize()
    Dim myIndex As Long
    Dim Major_High As Integer
    Dim Major_Low As Integer
    Dim Minor_High As Integer
    Dim Minor_Low As Integer
    Dim ValidVersion As Boolean
    Dim ns As NeroSpeeds
    Dim k As Long
    Dim strBuffer As String
    
    On Local Error GoTo Init_Error
    
    ' init Nero
    Set Nero = New Nero
    
    'Check valid version
    ValidVersion = True
    Nero.APIVersion Major_High, Major_Low, Minor_High, Minor_Low
    If Major_High < 6 Then
        ValidVersion = False
    ElseIf Major_High = 6 And Major_Low < 3 Then
        ValidVersion = False
    ElseIf Major_High = 6 And Major_Low = 3 And Minor_High < 1 Then
        ValidVersion = False
    ElseIf Major_High = 6 And Major_Low = 3 And Minor_High = 1 And Minor_Low < 6 Then
        ValidVersion = False
    End If
    
    ' valid version of Nero?
    If Not ValidVersion Then
            MsgBox "Nero Version 6.3.1.6 Or Greater Required!", vbCritical, App.Title
        End
    End If
    
    ' get Drive Nero version
    AddMessage "Init Nero:"
    AddMessage "Nero Version: " & "v." & Major_High & "." & Major_Low & "." & Minor_High & Minor_Low
    
    lblVN.Caption = "v." & Major_High & "." & Major_Low & "." & Minor_High & Minor_Low
    
    ' count available Drives
    Set Drives = Nero.GetDrives(NERO_MEDIA_CDR)
    lst_AvailableDevices.Clear
    For myIndex = 0 To Drives.Count - 1
        If Drives(myIndex).DevType = NERO_SCSI_DEVTYPE_WORM And _
            InStr(LCase$(Drives(myIndex).DeviceName), "image recorder") = 0 Then
            lst_AvailableDevices.AddItem Drives(myIndex).DeviceName, myIndex
        Else
            lst_AvailableDevices.AddItem Drives(myIndex).DeviceName, myIndex
        End If
    
    ' now retrive additional info
    Set Drive = Drives(myIndex)
    If Drives(myIndex).BufUnderrunProtName <> "" Then
            AddMessage "/*"
        AddMessage "Drive: " & myIndex & "-" & Drives(myIndex).DeviceName & " ** Device Ready (" & CStr(Drive.DeviceReady) & ")"
        
        'get read speed
        Set ns = Drive.AvailableSpeeds(NERO_ACCESSTYPE_READ, NERO_MEDIA_CDR + NERO_MEDIA_DVD_ANY)
            AddMessage "Base Read Speed: " & ns.BaseSpeedKBs & " Kb/s"
        For k = 0 To ns.Count - 1
            strBuffer = strBuffer & CStr(ns(k)) & "-"
        Next
            strBuffer = strBuffer & " Kb/s"
            AddMessage "Available Read Speeds: " & strBuffer
        
        'get write speed
        strBuffer = ""
            Set ns = Drive.AvailableSpeeds(NERO_ACCESSTYPE_WRITE, NERO_MEDIA_CDR + NERO_MEDIA_DVD_ANY)
            AddMessage "Base Write Speed: " & ns.BaseSpeedKBs & " Kb/s"
        For k = 0 To ns.Count - 1
            strBuffer = strBuffer & CStr(ns(k)) & "-"
        Next
            strBuffer = strBuffer & " Kb/s"
        ' stamp info
        AddMessage "Available Write Speeds: " & strBuffer
        AddMessage "Buffer Underrun Protection Name: " & Drive.BufUnderrunProtName
        AddMessage "Device Ready: " & CStr(Drive.DeviceReady)
        AddMessage "Drive buffer size: " & CStr(Drive.DriveBufferSize) & " Kb"
        AddMessage "*\"
    Else
        AddMessage "Drive: " & myIndex & "-" & Drives(myIndex).DeviceName & " ** Device Ready (" & CStr(Drive.DeviceReady) & ")"
    End If
    Next myIndex
    Set ns = Nothing
    
    ' use first Drive selected?
    If lst_AvailableDevices.ListCount > 0 Then
            lst_AvailableDevices.ListIndex = 0
        Set Drive = Drives(lst_AvailableDevices.ListIndex)
    End If

Exit Sub
Init_Error:
        Call WriteErrorLogs(Err.Number, Err.Description, "FormMain {Form: Initialize}", True, True)
    Err.Clear
End Sub

Private Sub Form_Load()
    
    On Error GoTo ErrorHandler
    
    ' double instance exit
    If App.PrevInstance Then
            AppActivate App.Title
                SendKeys "+", True
                readyToClose = True
            Unload Me
        End
    End If
    
    ' create the Folders backup if not exist
    Call CreateMyFoldersBackup
    
    ' .... Reset INI File Path
    INI.ResetINIFilePath
    
    ' .... Read the File INI
    If Dir$(App.path & "\" & App.EXEName & "_setting.ini") <> "" Then
        If INI.GetKeyValue("SETTING", "S_TOP") <> "" Then ssTop = INI.GetKeyValue("SETTING", "S_TOP")
        If INI.GetKeyValue("SETTING", "S_LEFT") <> "" Then ssLeft = INI.GetKeyValue("SETTING", "S_LEFT")
        ' .... Position MainForm
        If Len(ssLeft) = 0 Then
        ' .... Center Form
            ssTop = (Screen.Height - frmMain.Height) \ 2
            ssLeft = (Screen.Width - frmMain.Width) \ 2
            frmMain.Move ssLeft, ssTop
        Else
            frmMain.Move ssLeft, ssTop
        End If
        
        ' .... write type support
        If INI.GetKeyValue("SETTING", "cmbTypeSupport") <> "" Then _
        cmbTypeSupport.ListIndex = INI.GetKeyValue("SETTING", "cmbTypeSupport") Else cmbTypeSupport.ListIndex = 0
        
        ' speed writing
        If INI.GetKeyValue("SETTING", "cmbWSpeed") <> "" Then _
        cmbWSpeed.ListIndex = INI.GetKeyValue("SETTING", "cmbWSpeed") Else cmbWSpeed.ListIndex = 0
        
        ' write type support
        If INI.GetKeyValue("TYPE", "cmbType") <> "" Then _
        cmbType.ListIndex = INI.GetKeyValue("TYPE", "cmbType") Else cmbType.ListIndex = 0
        If INI.GetKeyValue("MODE", "cmbMode") <> "" Then _
        cmbMode.ListIndex = INI.GetKeyValue("MODE", "cmbMode") Else cmbMode.ListIndex = 0
        If INI.GetKeyValue("FILESYSTEM", "cmbSystem") <> "" Then _
        cmbSystem.ListIndex = INI.GetKeyValue("FILESYSTEM", "cmbSystem") Else cmbSystem.ListIndex = 0
        
        ' .... oter option
        If INI.GetKeyValue("SETTING", "CheckLog") <> "" Then CheckLog.value = INI.GetKeyValue("SETTING", "CheckLog")
        If INI.GetKeyValue("SETTING", "CheckEraseDisk") <> "" Then CheckEraseDisk.value = INI.GetKeyValue("SETTING", "CheckEraseDisk")
        If INI.GetKeyValue("SETTING", "CheckSimulate") <> "" Then CheckSimulate.value = INI.GetKeyValue("SETTING", "CheckSimulate")
        If INI.GetKeyValue("SETTING", "CheckAppend") <> "" Then CheckAppend.value = INI.GetKeyValue("SETTING", "CheckAppend")
        If INI.GetKeyValue("SETTING", "CheckOnlyISO") <> "" Then CheckOnlyISO.value = INI.GetKeyValue("SETTING", "CheckOnlyISO")
        
        ' mode writing speed
        If INI.GetKeyValue("FILESYSTEM", "cmbWSpeed") <> "" Then _
        cmbWSpeed.ListIndex = INI.GetKeyValue("FILESYSTEM", "cmbWSpeed") Else cmbWSpeed.ListIndex = 0
        
        ' mode erasing
        If INI.GetKeyValue("SETTING", "cmbErasing") <> "" Then cmbErasing.ListIndex = INI.GetKeyValue("SETTING", "cmbErasing") Else cmbErasing.ListIndex = 0
        
        ' assume true or false
        cmbErasing.Enabled = CheckEraseDisk.value
        
        ' burn type
        If INI.GetKeyValue("SETTING", "TYPE BURN") <> "" Then
            opt_B(INI.GetKeyValue("SETTING", "TYPE BURN")).value = True
        
        Select Case INI.GetKeyValue("SETTING", "TYPE BURN")
            Case 0
            ' default path
            If INI.GetKeyValue("SETTING", "LAST PATH SELECTED") <> "" Then
                txtPath.Text = INI.GetKeyValue("SETTING", "LAST PATH SELECTED")
                strLastPath = GetFilePath(txtPath.Text, Only_Path) + "\"
                lblISOFileSize.Caption = "Size: " & GetFileSize(txtPath.Text)
                txtSession.Text = "ISO"
                ISOTrackName.Text = "ISO-" & Format(Now, "dd-mm-yy")
            Else
                txtSession.Text = "ISO"
                ISOTrackName.Text = "ISO-" & Format(Now, "dd-mm-yy")
            End If
            cmdBurning.ToolTipText = "Start Burning Folder/File ISO"
            Case 1
                txtSession.Text = "YBC"
                ISOTrackName.Text = "YBC-" & Format(Now, "dd-mm-yy")
                cmdBurning.ToolTipText = "Start Burning Folder/File YBC"
            Case 2
                txtSession.Text = "LOG"
                ISOTrackName.Text = "LOG-" & Format(Now, "dd-mm-yy")
                cmdBurning.ToolTipText = "Start Burning Folder/File LOGO"
            Case 3
                txtSession.Text = "BEND"
                ISOTrackName.Text = "BEND-" & Format(Now, "dd-mm-yy")
                cmdBurning.ToolTipText = "Start Burning Folder/File BEND"
            Case 4
                txtSession.Text = "BACKUP"
                ISOTrackName.Text = "BACKUP-" & Format(Now, "dd-mm-yy")
                cmdBurning.ToolTipText = "Start Burning Folder/File BACKUP"
    End Select
        Else
            opt_B(1).value = True
        End If
    Else
        ' .... Center Form
        ssTop = (Screen.Height - frmMain.Height) \ 2
        ssLeft = (Screen.Width - frmMain.Width) \ 2
        frmMain.Move ssLeft, ssTop
        
    End If
    
    ' get last path
    If INI.GetKeyValue("SETTING", "LAST PATH SELECTED") <> "" Then
        strLastPath = GetFilePath(INI.GetKeyValue("SETTING", "LAST PATH SELECTED"), Only_Path) + "\"
    Else
        strLastPath = App.path + "\"
    End If
    
Exit Sub
ErrorHandler:
    Call WriteErrorLogs(Err.Number, Err.Description, "FormMain {Form: Load}", True, True)
Err.Clear
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("Are you sure to Close this Application?", vbYesNo + vbInformation + _
        vbDefaultButton1, App.Title) = vbYes Then
        readyToClose = True
        ' .... Release the Deugger
        SetUnhandledExceptionFilter ByVal 0&
        ' .... Release the Library
        Call FreeLibrary(m_hMod)
        ' .... UnHook the SO
        If Not InIDE() Then SetErrorMode SEM_NOGPFAULTERRORBOX
        ' .... Release Lyb NEROCom
        Set Drive = Nothing
        Set Nero = Nothing
        ' .... Save Setting to File *.INI
        SaveSettingINI
        ' .... Unload Class INI
        Set INI = Nothing
    Else
        readyToClose = False
    End If
    Cancel = Not readyToClose
End Sub

Private Sub Form_Terminate()
    On Error Resume Next
    Set frmMain = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub ISOTrackName_KeyUp(KeyCode As Integer, Shift As Integer)
    On Local Error Resume Next
    If Len(ISOTrackName.Text) > 16 Then
        MsgBox "Too much char!" & vbCr & "The name truncated to 16 chars!", vbExclamation, App.Title
        ISOTrackName.Text = Mid(ISOTrackName.Text, 1, 16)
        ISOTrackName.SelStart = Len(ISOTrackName.Text)
        ISOTrackName.SetFocus
        Exit Sub
    End If
End Sub

Private Sub ISOTrackName_LostFocus()
    On Local Error Resume Next
    If ISOTrackName.Text <> "" Then ISOTrackName.Text = UCase$(Replace(ISOTrackName.Text, " ", "_"))
End Sub


Private Sub lst_AvailableDevices_Click()
    Dim myIndex As Long
    On Local Error GoTo ErrorHandler
    
    ' capture the Drive selected
    'Set Drives = Nero.GetDrives(NERO_MEDIA_CDR)
    Set Drive = Drives(lst_AvailableDevices.ListIndex)
    
    If LCase$(lst_AvailableDevices.Text) = "image recorder" Then
            cmdBurning.Enabled = False
            cmdBurFile.Enabled = False
            cmdEject.Enabled = False
            CheckEraseDisk.value = 0
            CheckEraseDisk.Enabled = False
            CheckOnlyISO.Enabled = True
            CheckOnlyISO.value = 1
            cmdBurnImage.Enabled = True
            CheckSimulate.Enabled = False
        Exit Sub
    Else
        CheckOnlyISO.value = 0
        CheckOnlyISO.Enabled = False
        cmdBurnImage.Enabled = False
        CheckSimulate.Enabled = True
        ' verify the drive
    If Drives(lst_AvailableDevices.ListIndex).BufUnderrunProtName = "" Then
        If LCase$(lst_AvailableDevices.Text) <> "image recorder" Then
            Beep
            cmdBurning.Enabled = False
            cmdBurFile.Enabled = False
            cmdEject.Enabled = False
            CheckSimulate.Enabled = False
        End If
    
    ElseIf Drives(lst_AvailableDevices.ListIndex).BufUnderrunProtName <> "" Then
        If LCase$(lst_AvailableDevices.Text) <> "image recorder" Then
            cmdBurning.Enabled = True
            cmdBurFile.Enabled = True
            cmdEject.Enabled = True
            CheckSimulate.Enabled = True
        End If
    End If
End If
    
    CheckEraseDisk.Enabled = True
        
Exit Sub
ErrorHandler:
        cmdBurning.Enabled = False
        cmdEject.Enabled = False
    Err.Clear
End Sub


Private Sub Nero_OnCopyQualityLoss(Response As nerolib.NERO_RESPONSE)
    AddMessage "Quality Loss: " & Response
End Sub

Private Sub Nero_OnDlgMessageBox(ByVal pDlgMessageBox As nerolib.INeroDlgMessageBox, Response As nerolib.NERO_RESPONSE)
    AddMessage Response
End Sub


Private Sub Nero_OnDriveStatusChanged(ByVal hostID As Long, ByVal targetID As Long, ByVal driveStatus As nerolib.NERO_DRIVECHANGE_RESULT)
    AddMessage hostID
    AddMessage targetID
    AddMessage driveStatus
End Sub
Private Sub Nero_OnFileSelImage(FileName As String)
    Dim IDir As String
    On Local Error GoTo Error_ISO
    With cDialog
        IDir = INI.GetKeyValue("SETTING", "LAST PATH SELECTED")
        If IDir = "" Then IDir = App.path Else IDir = GetFilePath(INI.GetKeyValue("SETTING", "LAST PATH SELECTED"), Only_Path)
        .CancelError = True
        .Filter = "Supported ISO Files (*.iso)|*.iso|Supported Nero Image (*.nrg)|*.nrg|Image ISO (*.cue)|*.cue|Image ISO (*.bin)|*.bin"
        .DialogTitle = "Save ISO image as:"
        .InitDir = IDir
        .FilterIndex = 2 ' = *.nrg Nero Image ;)
        .DefaultExt = ".nrg"
        .ShowSave
        If .FileName = "" Then Exit Sub
    FileName = .FileName
    End With
    sISOFile = FileName
Exit Sub
Error_ISO:
        Call WriteErrorLogs(Err.Number, Err.Description, Nero.LastError, True, True)
    Err.Clear
End Sub

' Dispaly the Fatal Error...
Private Sub Nero_OnMegaFatal()
    AddMessage "A fatal error has occurred."
    BurnError = True
    Me.Caption = "Burning-iX v1.0.1ß (beta) 2009 © Salvo Cortesiano."
End Sub


' The CD/DVD is Not empty...
Private Sub Nero_OnNonEmptyCDRW(Response As nerolib.NERO_RESPONSE)
    AddMessage "The CD-RW/DVD is not empty!"
    Response = NERO_RETURN_EXIT
End Sub


Private Sub Nero_OnNoTrackFound()
    AddMessage "No Track found!"
End Sub


Private Sub Nero_OnOverburn(Response As nerolib.NERO_RESPONSE)
    AddMessage "OverBurn " & Response
End Sub


Private Sub Nero_OnOverburn2(ByVal pOverburnInfo As nerolib.INeroOverburnInfo, Response As nerolib.NERO_RESPONSE)
    AddMessage "OverBurn2 " & Response
    AddMessage "Total Blocks on CD " & pOverburnInfo.TotalBlocksOnCD
    AddMessage "Total Capacity " & pOverburnInfo.TotalCapacity
End Sub


' Info to Restart the SO
Private Sub Nero_OnRestart()
    AddMessage "The System is being restarted."
End Sub


Private Sub Nero_OnRoboMoveUserMessage(ByVal messageType As nerolib.ROBOUSERMESSAGETYPE, ByVal bstrMessage As String, Response As nerolib.NERO_RESPONSE)
    AddMessage Response
End Sub
Private Sub Nero_OnSettingsRestart(Response As nerolib.NERO_RESPONSE)
    AddMessage Response
End Sub

Private Sub Nero_OnTempSpace(ByVal bstrCurrentDir As String, ByVal pi64FreeSpace As nerolib.IInt64, ByVal pi64SpaceNeeded As nerolib.IInt64, pbstrNewTempDir As String)
    bstrCurrentDir = App.path
    pbstrNewTempDir = App.path
End Sub


' Wait for the Empty CD/DVD...
Private Sub Nero_OnWaitCD(WaitCD As nerolib.NERO_WAITCD_TYPE, WaitCDLocalizedText As String)
    SplitText WaitCDLocalizedText
End Sub


' Release the CD/DVD...
Private Sub Nero_OnWaitCDDone()
    AddMessage "Done waiting for CD."
End Sub


' Wait for Request Media Type CD/DVD...
Private Sub Nero_OnWaitCDMediaInfo(LastDetectedMedia As nerolib.NERO_MEDIA_TYPE, LastDetectedMediaName As String, RequestedMedia As nerolib.NERO_MEDIA_TYPE, RequestedMediaName As String)
    AddMessage "Waiting for a particular media type: " + RequestedMediaName
End Sub


' Wait for the CD/DVD...
Private Sub Nero_OnWaitCDReminder()
    AddMessage "Still waiting for CD..."
End Sub



Private Sub Delay(ByVal Time As Single, Optional ByVal ForceWait As Boolean = False)
    Dim Start
    Dim X
    Dim SleepVal As Long
    Start = Timer
        While Start + Time > Timer
    If Start > Timer Then
        Start = Timer
    End If
        If Not ForceWait Then
            X = DoEvents()
        End If
    Wend
End Sub

Private Function MakeDirectory(szDirectory As String) As Boolean
Dim strFolder As String
Dim szRslt As String
On Error GoTo IllegalFolderName
If Right$(szDirectory, 1) <> "\" Then szDirectory = szDirectory & "\"
strFolder = szDirectory
szRslt = Dir(strFolder, 63)
While szRslt = ""
    DoEvents
    szRslt = Dir(strFolder, 63)
    strFolder = Left$(strFolder, Len(strFolder) - 1)
    If strFolder = "" Then GoTo IllegalFolderName
Wend
If Right$(strFolder, 1) <> "\" Then strFolder = strFolder & "\"
While strFolder <> szDirectory
    strFolder = Left$(szDirectory, Len(strFolder) + 1)
    If Right$(strFolder, 1) = "\" Then MkDir strFolder
Wend
MakeDirectory = True
Exit Function
IllegalFolderName:
    Err.Clear
End Function

Private Function DirExist(Source_Dir As String) As Boolean
    On Local Error GoTo FolderError
    If (GetAttr(Source_Dir) And vbDirectory) = vbDirectory Then
        DirExist = True
    Else
        DirExist = False
    End If
Exit Function
FolderError:
        DirExist = False
    Err.Clear
End Function


Private Function BrowseFolder(ByVal strTitle As String, Optional strPath As String = "") As String
    Dim fOlder As String
    On Local Error GoTo ErrorHandler
    If strPath = "" Then
        strPath = App.path + "\"
    Else
        If Right$(strPath, 1) <> "\" Then strPath = strPath + "\"
    End If
    fOlder = BrowseForFolder(Me.hWnd, strTitle, strPath)
    If fOlder <> "" Then BrowseFolder = fOlder Else BrowseFolder = ""
Exit Function
ErrorHandler:
        BrowseFolder = "Error!"
    Err.Clear
End Function

Private Function PlaySoundResource(ByVal SndID As Long) As Long
   Const flags = SND_MEMORY Or SND_ASYNC Or SND_NODEFAULT
   On Error GoTo ErrorHandler
   DoEvents
   m_snd = LoadResData(SndID, "WAVE")
   PlaySoundResource = PlaySoundData(m_snd(0), 0, flags)
Exit Function
ErrorHandler:
    Err.Clear
End Function

Private Function GetFilePath(ByVal FileName As String, strExtract As Extract) As String
    Select Case strExtract
        'Extract only extension of File
    Case 0
         GetFilePath = Mid$(FileName, InStrRev(FileName, ".", , vbTextCompare) + 1)
        'Extract only Filename and Extension
    Case 1
        GetFilePath = Mid$(FileName, InStrRev(FileName, "\") + 1, Len(FileName))
        'Extract only FileName
   Case 2
        GetFilePath = StripString(Mid$(FileName, InStrRev(FileName, "\", , vbTextCompare) + 1))
        'Extract only Path
   Case 3
        GetFilePath = Mid$(FileName, 1, InStrRev(FileName, "\", , vbTextCompare) - 1)
   End Select
End Function

Private Function StripString(ByVal sString As String) As String
    Dim i As Integer
    Dim stmp As String
    On Error Resume Next
    stmp = Mid(sString, i + 1, Len(sString))
    For i = 1 To Len(stmp)
      If Mid(stmp, i, 1) = "." Then
        Exit For
    Else
        MyString = Mid(sString, i + 2, Len(sString))
    End If
Next
     StripString = Left(stmp, i - 1)
End Function

Private Sub ResetAll()
    If LCase$(lst_AvailableDevices.Text) = "image recorder" Then
        cmdAbort.Enabled = False
        cmdBurning.Enabled = False
        cmdBurFile.Enabled = False
        cmdBurnImage.Enabled = True
        cmdExit.Enabled = True
        cmdEject.Enabled = False
    Else
        cmdAbort.Enabled = False
        cmdBurning.Enabled = True
        cmdBurFile.Enabled = True
        cmdBurnImage.Enabled = False
        cmdExit.Enabled = True
        cmdEject.Enabled = True
    End If
        TBurning.Enabled = False
        TBurning.Interval = 0
        pgs_Burn.value = 0
        pgs_Buffer.value = 0
        fme_Progress.value = 0
        lblPercentuale.Caption = "00%"
        lblTCriting.Caption = "00%"
        lblTWriting.Caption = "00%"
End Sub

Private Function WriteLog(strTxt As String)
    Dim FF As Variant
    On Local Error GoTo ErrorLog
    FF = FreeFile
    FileLogOfNeroCom = App.path + "\NeroCOM_Log_" + Format(Now, "dd") + "-" + Format(Now, "mm") + "-" + Format(Now, "yyyy") + ".txt"
    If Dir$(FileLogOfNeroCom) = "" Then
        Open FileLogOfNeroCom For Output As #FF
            Print #FF, Tab(5); "Log NeroCOM Generate from [" & App.EXEName & "]..."
            Print #FF, Tab(5); Format(Now, "Long Date") & "/" & Time
            Print #FF, Tab(5); "----------------------------------------------------------------------------"
            Print #FF, Tab(5); ""
            Print #FF, Tab(5); ""
            Print #FF, Tab(5); "*/___ LOG NEROCOM STARTED..."
            Print #FF, Tab(5); ""
            Print #FF, Tab(5); ""
            Print #FF, Tab(5); Time & "} " & strTxt
        Close #FF
    Else
        If CheckAppend.value = 1 Then Open FileLogOfNeroCom For Append As #FF Else Open FileLogOfNeroCom For Output As #FF
            Print #FF, Tab(5); Time & "} " & strTxt
            Print #FF, Tab(5); ""
            Print #FF, Tab(5); ""
        Close #FF
    End If
Exit Function
ErrorLog:
    Err.Clear
End Function

Private Sub Restart(WhatLabel As Label)
    WhatLabel.Caption = "00:00:00": Seconds = 0: Minutes = 0: Hours = 0: Days = 0
End Sub

Private Sub StopWatch(WhatLabel As Label)
Dim addSeconds As Long
On Error GoTo ErrorHandler
If Seconds = 60 Then
    AddMinutes = True
    addSeconds = 0
    Seconds = 0
Else
    Seconds = Seconds + 1
End If
If AddMinutes = True Then
    If Minutes = 60 Then
    AddHours = True
    Minutes = 0
Else
        Minutes = Minutes + 1
        'lblMinutes.Caption = lblMinutes.Caption + 1
    End If
    AddMinutes = False
End If
If AddHours = True Then
    If Hours = 24 Then
        AddDays = True
        Hours = 0
        Days = Days + 1
    Else
        Hours = Hours + 1
    End If
    AddHours = False
End If

If AddDays = True Then
    If Days = 999 Then
    Days = 0
Else
        Days = Days + 1
    End If
    AddDays = False
End If
' my old code :)
'WhatLabel.Caption = Format(Days, "000") & ":" & Format(Hours, "00") & ":" & Format(Minutes, "00") & ":" & Format(Seconds, "00")
WhatLabel.Caption = Format(Hours, "00") & ":" & Format(Minutes, "00") & ":" & Format(Seconds, "00")
Exit Sub
ErrorHandler:
    Err.Clear
    Call Restart(Label8)
    TBurning.Enabled = False
    TBurning.Interval = 0
End Sub

Private Sub opt_B_Click(index As Integer)
    Select Case index
        Case 0
            ' default path
            If INI.GetKeyValue("SETTING", "LAST PATH SELECTED") <> "" Then
                txtPath.Text = INI.GetKeyValue("SETTING", "LAST PATH SELECTED")
                strLastPath = GetFilePath(txtPath.Text, Only_Path) + "\"
                lblISOFileSize.Caption = "Size: " & GetFileSize(txtPath.Text)
                txtSession.Text = "ISO"
                ISOTrackName.Text = "ISO-" & Format(Now, "dd-mm-yy")
            Else
                txtSession.Text = "ISO"
                ISOTrackName.Text = "ISO-" & Format(Now, "dd-mm-yy")
            End If
            cmdBurning.ToolTipText = "Start Burning Folder/File ISO"
        Case 1
            txtSession.Text = "YBC"
            ISOTrackName.Text = "YBC-" & Format(Now, "dd-mm-yy")
            cmdBurning.ToolTipText = "Start Burning Folder/File YBC"
        Case 2
            txtSession.Text = "LOG"
            ISOTrackName.Text = "LOG-" & Format(Now, "dd-mm-yy")
            cmdBurning.ToolTipText = "Start Burning Folder/File LOGO"
        Case 3
            txtSession.Text = "BEND"
            ISOTrackName.Text = "BEND-" & Format(Now, "dd-mm-yy")
            cmdBurning.ToolTipText = "Start Burning Folder/File BEND"
        Case 4
            txtSession.Text = "BACKUP"
            ISOTrackName.Text = "BACKUP-" & Format(Now, "dd-mm-yy")
            cmdBurning.ToolTipText = "Start Burning Folder/File BACKUP"
    End Select
    
    If txtSession.Text <> "" Then txtSession.Text = UCase$(Replace(txtSession.Text, " ", "_"))
    If ISOTrackName.Text <> "" Then ISOTrackName.Text = UCase$(Replace(ISOTrackName.Text, " ", "_"))
    
    If Len(ISOTrackName.Text) > 16 Then
        MsgBox "Too much char!" & vbCr & "The name truncated to 16 chars!", vbExclamation, App.Title
        ISOTrackName.Text = Mid(ISOTrackName.Text, 1, 16)
        ISOTrackName.SelStart = Len(ISOTrackName.Text)
        ISOTrackName.SetFocus
    ElseIf Len(txtSession.Text) > 16 Then
        MsgBox "Too much char!" & vbCr & "The name truncated to 16 chars!", vbExclamation, App.Title
        txtSession.Text = Mid$(txtSession.Text, 1, 16)
        txtSession.SelStart = Len(txtSession.Text)
        txtSession.SetFocus
    End If
End Sub


Private Sub TBurning_Timer()
        Call StopWatch(Label8)
    DoEvents
End Sub



Private Sub StartCount()
    Call Restart(Label8)
    TBurning.Interval = 900
    TBurning.Enabled = True
End Sub

Private Sub SaveSettingINI()
    Dim i As Integer
    On Local Error Resume Next
    
    ' option burning
    For i = 0 To 4
        If opt_B(i).value = True Then
                INI.DeleteKey "SETTING", "TYPE BURN"
                INI.CreateKeyValue "SETTING", "TYPE BURN", opt_B(i).index
            ' until next
            Exit For
        End If
    Next i
    
    ' oter option
    INI.DeleteKey "SETTING", "CheckLog"
    INI.CreateKeyValue "SETTING", "CheckLog", CheckLog.value
    INI.DeleteKey "SETTING", "CheckEraseDisk"
    INI.CreateKeyValue "SETTING", "CheckEraseDisk", CheckEraseDisk.value
    INI.DeleteKey "SETTING", "CheckSimulate"
    INI.CreateKeyValue "SETTING", "CheckSimulate", CheckSimulate.value
    INI.DeleteKey "SETTING", "CheckAppend"
    INI.CreateKeyValue "SETTING", "CheckAppend", CheckAppend.value
    
    ' ISO Image option
    INI.DeleteKey "SETTING", "CheckOnlyISO"
    INI.CreateKeyValue "SETTING", "CheckOnlyISO", CheckOnlyISO.value
    
    ' write type support
    INI.DeleteKey "TYPE", "cmbType"
    INI.CreateKeyValue "TYPE", "cmbType", cmbType.ListIndex
    INI.DeleteKey "MODE", "cmbMode"
    INI.CreateKeyValue "MODE", "cmbMode", cmbMode.ListIndex
    INI.DeleteKey "FILESYSTEM", "cmbSystem"
    INI.CreateKeyValue "FILESYSTEM", "cmbSystem", cmbSystem.ListIndex
    
    ' speed writing
    INI.DeleteKey "FILESYSTEM", "cmbWSpeed"
    INI.CreateKeyValue "FILESYSTEM", "cmbWSpeed", cmbWSpeed.ListIndex
    
    ' mode Writing
    INI.DeleteKey "SETTING", "cmbTypeSupport"
    INI.CreateKeyValue "SETTING", "cmbTypeSupport", cmbTypeSupport.ListIndex
    
    ' mode Erasing
    INI.DeleteKey "SETTING", "cmbErasing"
    INI.CreateKeyValue "SETTING", "cmbErasing", cmbErasing.ListIndex
    
    ' mode writing speed
    INI.DeleteKey "SETTING", "cmbWSpeed"
    INI.CreateKeyValue "SETTING", "cmbWSpeed", cmbWSpeed.ListIndex
    
    ' position form
    If Me.WindowState <> vbMinimized Then
        INI.DeleteKey "SETTING", "S_LEFT"
        INI.CreateKeyValue "SETTING", "S_LEFT", frmMain.Left
        INI.DeleteKey "SETTING", "S_TOP"
        INI.CreateKeyValue "SETTING", "S_TOP", frmMain.Top
    End If
    
    ' last file ISO selected
    INI.DeleteKey "SETTING", "LAST PATH SELECTED"
    INI.CreateKeyValue "SETTING", "LAST PATH SELECTED", txtPath.Text
    
End Sub

Private Sub txtSession_KeyUp(KeyCode As Integer, Shift As Integer)
    On Local Error Resume Next
    If Len(txtSession.Text) > 16 Then
        MsgBox "Too much char!" & vbCr & "The name truncated to 16 chars!", vbExclamation, App.Title
        txtSession.Text = Mid$(txtSession.Text, 1, 16)
        txtSession.SelStart = Len(txtSession.Text)
        txtSession.SetFocus
        Exit Sub
    End If
End Sub


Private Sub txtSession_LostFocus()
    On Local Error Resume Next
    If txtSession.Text <> "" Then txtSession.Text = UCase$(Replace(txtSession.Text, " ", "_"))
End Sub



Private Sub CreateMyFoldersBackup()
    ' set to be whatever folders you need to backup
    ' create my folders backup if not exist :)
    If MakeDirectory(App.path + "\mybackup\bend\ybc") Then: ' until display the warning msg
    If MakeDirectory(App.path + "\mybackup\bend\log") Then: '  ""    ""      ""      ""
    If MakeDirectory(App.path + "\mybackup\bend\backup") Then: '  ""    ""      ""      ""
    If MakeDirectory(App.path + "\mybackup\bend\ISO") Then: '  ""    ""      ""      ""
End Sub
