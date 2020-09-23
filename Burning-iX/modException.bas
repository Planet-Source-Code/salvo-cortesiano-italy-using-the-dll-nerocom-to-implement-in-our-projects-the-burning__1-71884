Attribute VB_Name = "modException"
Option Explicit

' .... Class INI
Public INI As New clsINI

' ... Init control's XP or Vista
Private Type tagInitCommonControlsEx
    lngSize As Long
    lngICC As Long
End Type

Private Declare Function InitCommonControlsEx Lib "Comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public m_hMod As Long

'ITALIAN SORRY :(
' Di conseguenza possiamo risolvere questo problema semplicemente ignorandolo.
' L'unico problema in questo modo è che l'applicazione continua a inviare messaggi al sistema e danno origine
' alla nota finestra che invita a trasmettere le informazioni del Microsoft sul problema:
Public Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Public Const SEM_FAILCRITICALERRORS = &H1
Public Const SEM_NOGPFAULTERRORBOX = &H2
Public Const SEM_NOOPENFILEERRORBOX = &H8000&

' ... Exception Handler (Call the Stack)
Public Const MySEH_ERROR = 12345&

Public Declare Sub RaiseException Lib "kernel32" (ByVal dwExceptionCode As Long, ByVal dwExceptionFlags As Long, ByVal nNumberOfArguments As Long, lpArguments As Long)
Public Declare Function SetUnhandledExceptionFilter Lib "kernel32" (ByVal lpTopLevelExceptionFilter As Long) As Long

Private Const EXCEPTION_CONTINUE_EXECUTION = -1
Private Const EXCEPTION_CONTINUE_SEARCH = 0
Private Const EXCEPTION_EXECUTE_HANDLER = 1

Public Declare Sub DebugBreak Lib "kernel32" ()
Private m_bInIDE As Boolean

Private Sub InitControlsCtx()
 On Local Error GoTo ErrorHandler
      Dim iccex As tagInitCommonControlsEx
      With iccex
          .lngSize = LenB(iccex)
          .lngICC = ICC_USEREX_CLASSES
      End With
      InitCommonControlsEx iccex
Exit Sub
ErrorHandler:
    Err.Clear
End Sub

Public Function MyExceptionHandler(lpEP As Long) As Long
   Dim lRes As VbMsgBoxResult
   lRes = MsgBox("Exception Handler!" & vbCrLf & "Ignore, Close, or Call the Debugger?", _
   vbAbortRetryIgnore Or vbCritical, App.Title & "Exception Handler")
   Select Case lRes
      Case vbIgnore
         If InIDE Then
            Stop
            MyExceptionHandler = EXCEPTION_CONTINUE_EXECUTION
            On Error GoTo 0
            Err.Raise MySEH_ERROR, "MyExceptionHandler", "EXCEPTION ???"
         Else
            MyExceptionHandler = EXCEPTION_CONTINUE_SEARCH
            Err.Raise MySEH_ERROR, "MyExceptionHandler", "EXCEPTION ???"
         End If
       Case vbAbort
         MyExceptionHandler = EXCEPTION_EXECUTE_HANDLER
         Err.Raise MySEH_ERROR, "MyExceptionHandler", "EXCEPTION ???"
       Case vbRetry
         MyExceptionHandler = EXCEPTION_CONTINUE_SEARCH
         Err.Raise MySEH_ERROR, "MyExceptionHandler", "EXCEPTION ???"
   End Select
End Function

Public Sub Main()
    On Local Error GoTo ErrorHandler
    ' .... Exception Handler' = (Call the stack)
    SetUnhandledExceptionFilter AddressOf MyExceptionHandler
    ' .... Subclass the SO
    SetErrorMode SEM_NOGPFAULTERRORBOX
    ' .... Load the Library
    m_hMod = LoadLibrary("shell32.dll")
    ' .... Init the Controls
    InitControlsCtx
    ' .... Show the Form
    Load frmMain
    frmMain.Show
Exit Sub
ErrorHandler:
    Call WriteErrorLogs(Err.Number, Err.Description, "ModMain {Sub: Main}", True, True)
        Err.Clear
    End
End Sub

Public Sub WriteErrorLogs(strErrNumber As String, strErrDescription As String, Optional strErrSource As String = "Unknow", _
                        Optional visError As Boolean = True, Optional errAppend As Boolean = True)
    Dim FileNum As Integer
    Dim sFN As String
    On Error GoTo ErrorHandler
    FileNum = FreeFile
    sFN = App.path & "\" & App.EXEName & "\_errs.log"
    If Dir$(sFN) = "" Then
        Open sFN For Output As FileNum
        Print #FileNum, Tab(5); "Log Error Generate from [" & App.EXEName & "]..."
        Print #FileNum, Tab(5); Format(Now, "Long Date") & "/" & Time
        Print #FileNum, Tab(5); "----------------------------------------------------------------------------"
        Print #FileNum, Tab(5); ""
        Print #FileNum, Tab(5); ""
        Print #FileNum, Tab(5); "*/___ LOG STARTED..."
        Print #FileNum, Tab(5); ""
        Close FileNum
    End If
    If errAppend Then Open sFN For Append As FileNum Else Open sFN For Output As FileNum
        Print #FileNum, Tab(5); Format(Now, "Long Date") & "/" & Time
        Print #FileNum, Tab(5); "Error #" & CStr(strErrNumber)
        Print #FileNum, Tab(5); "Description: " & CStr(strErrDescription)
        Print #FileNum, Tab(5); "Source: " & CStr(strErrSource)
        Print #FileNum, Tab(5); ""
        Print #FileNum, Tab(5); ""
        Close FileNum
        If visError Then
            MsgBox "Error #" & CStr(strErrNumber) & "." & vbCrLf & "Description: " & CStr(strErrDescription) _
            & vbCrLf & "Source: " & CStr(strErrSource) & vbCrLf & vbCrLf & "For more info, see the Log file!", vbCritical, App.Title
        End If
    Exit Sub
ErrorHandler:
    ' until display the Dialog-box
        'MsgBox "Unexpected Error #" & Err.Number & "!" & vbCrLf & Err.Description, vbCritical, App.Title
    Err.Clear
End Sub

Public Property Get InIDE() As Boolean
   Debug.Assert (pIsInIDE)
   InIDE = m_bInIDE
End Property

Public Property Get pIsInIDE() As Boolean
   m_bInIDE = True
   pIsInIDE = True
End Property
