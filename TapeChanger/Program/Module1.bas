Attribute VB_Name = "Module1"
Option Explicit

' Versions:
' VER_PLATFORM_WIN32_WINDOWS(1)
'   W95 = 4.0
'   W98 = 4.1
'   WME = 4.9
' VER_PLATFORM_WIN32_NT(2)
'   WNT Serv = 4.0
'   W2k Prof = 5.0
'   WXP Prof = 5.1

Public Const OS_UNKNOWN = -1
Public Const OS_WIN95 = 0
Public Const OS_WIN98 = 1
Public Const OS_WINNT35 = 2
Public Const OS_WINNT4 = 3
Public Const OS_WIN2K = 4

'Version structure
Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Public Type PROCESS_INFORMATION
      hProcess As Long
      hThread As Long
      dwProcessID As Long
      dwThreadID As Long
End Type

   

Public Type STARTUPINFO
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


'dwPlatformId defines
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2
Public Const SW_SHOWNORMAL = 1
Public Const GW_HWNDNEXT = 2
Public Const NORMAL_PRIORITY_CLASS = &H20&
Public Const INFINITE = -1&

Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, _
  ByVal wCmd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
  (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
  (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" _
  (ByVal hwnd As Long, lpdwprocessid As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function CreateProcessA Lib "kernel32" (ByVal _
      lpApplicationName As String, ByVal lpCommandLine As String, ByVal _
      lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
      ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
      ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
      lpStartupInfo As STARTUPINFO, lpProcessInformation As _
      PROCESS_INFORMATION) As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

Public Declare Function GetVersionEx Lib "kernel32" _
    Alias "GetVersionExA" _
    (lpVersionInformation As OSVERSIONINFO) As Long

Public Declare Function LoadLibraryEx Lib "kernel32" _
    Alias "LoadLibraryExA" _
    (ByVal lpLibFileName As String, ByVal hFile As Long, _
     ByVal dwFlags As Long) As Long

Public Declare Function FreeLibrary Lib "kernel32" _
    (ByVal hLibModule As Long) As Long

Public Declare Function LoadLibrary Lib "kernel32" _
    Alias "LoadLibraryA" _
    (ByVal lpLibFileName As String) As Long

Public Declare Function GetProcAddress Lib "kernel32" _
    (ByVal hModule As Long, ByVal lpProcName As String) As Long

'*******************************************************************
'** TOC format
'*******************************************************************
Type TOC_TRACK
    Rsvd1 As Byte
    ADR As Byte
    Track As Byte
    Rsvd2 As Byte
    Addr(3) As Byte
End Type

Type TOC
    TocLen(1) As Byte
    FirstTrack As Byte
    LastTrack As Byte
    TocTrack(99) As TOC_TRACK
End Type

Type Changer_element_status
    header(7) As Byte
    ElementStatusPages(1023) As Byte
End Type

Type Changer_Inquiry
    Info(94) As Byte
End Type

Global bRet As Boolean
Global nRet As Long
Global i As Integer
Global j As Integer
Global cnt As Integer
Global Inquiry As SRB_HAInquiry
Global DevType As SRB_GetDevType
Global ExecIO As SRB_ExecuteIO
Global DataBuffer As TOC
Global Databuffer1 As Changer_element_status
Global Databuffer2 As Changer_Inquiry
Global JUKE_HA As Long
Global JUKE_ID As Long
Global ChangerHostAdapter As Long
Global ChangerSCSIID As Long
Global Adapter As Long
Global Device As Long
Global TapeNowLoaded As Integer
Global CleanHeadTimeout As Integer

'*******************************************************************
'** ASPI command definitions
'*******************************************************************
Public Const SC_HA_INQUIRY = &H0        'Host adapter inquiry
Public Const SC_GET_DEV_TYPE = &H1      'Get device type
Public Const SC_EXEC_SCSI_CMD = &H2     'Execute SCSI command
Public Const SC_ABORT_SRB = &H3         'Abort an SRB
Public Const SC_RESET_DEV = &H4         'SCSI bus device reset
Public Const SC_SET_HA_PARMS = &H5      'Set HA parameters
Public Const SC_GET_DISK_INFO = &H6     'Get Disk
Public Const SC_RESCAN_SCSI_BUS = &H7   'Rebuild SCSI device map
Public Const SC_GETSET_TIMEOUTS = &H8   'Get/Set target timeouts

'*******************************************************************
'** SRB Status
'*******************************************************************
Public Const SS_PENDING = &H0           'SRB being processed
Public Const SS_COMP = &H1              'SRB completed without error
Public Const SS_ABORTED = &H2           'SRB aborted                    */
Public Const SS_ABORT_FAIL = &H3        'Unable to abort SRB
Public Const SS_ERR = &H4               'SRB completed with error
Public Const SS_INVALID_CMD = &H80      'Invalid ASPI command
Public Const SS_INVALID_HA = &H81       'Invalid host adapter number
Public Const SS_NO_DEVICE = &H82        'SCSI device not installed
Public Const SS_INVALID_SRB = &HE0      'Invalid parameter set in SRB
Public Const SS_OLD_MANAGER = &HE1      'ASPI manager doesn't support windows
Public Const SS_BUFFER_ALIGN = &HE1     'Buffer not aligned (SS_OLD_MANAGER in Win32)
Public Const SS_ILLEGAL_MODE = &HE2     'Unsupported Windows mode
Public Const SS_NO_ASPI = &HE3          'No ASPI managers
Public Const SS_FAILED_INIT = &HE4      'ASPI for windows failed init
Public Const SS_ASPI_IS_BUSY = &HE5     'No resources available to execute command
Public Const SS_BUFFER_TOO_BIG = &HE6   'Buffer size too big to handle
Public Const SS_MISMATCH_FILES = &HE7   'The DLLs/EXEs of ASPI don't version check
Public Const SS_NO_ADAPTERS = &HE8      'No host adapters located
Public Const SS_SHORT_RESOURCES = &HE9  'Couldn't allocate resources  needed to init
Public Const SS_ASPI_IS_SHUTDOWN = &HEA 'Call came to ASPI after PROCESS_DETACH
Public Const SS_BAD_INSTALL = &HEB      'The DLL or other components are installed wrong

'*******************************************************************
'** ASPI Command Packets
'*******************************************************************
'** SRB - COMMAND HEADER COMMON
Public Type SRB
    SRB_Cmd As Byte             '00h/00 ASPI command code
    SRB_Status As Byte          '01h/01 ASPI command status byte
    SRB_HaID As Byte            '02h/02 ASPI host adapter number
    SRB_Flags As Byte           '03h/03 ASPI request flags
    SRB_Hdr_Rsvd As Long        '04h/04 Reserved, must = 0
End Type

'** SRB - HOST ADAPTER INQUIRIY - SC_HA_INQUIRY (0)
Public Type SRB_HAInquiry
    SRB_Cmd As Byte             '00h/00 ASPI command code == SC_HA_INQUIRY
    SRB_Status As Byte          '01h/01 ASPI command status byte
    SRB_HaID As Byte            '02h/02 ASPI host adapter number
    SRB_Flags As Byte           '03h/03 ASPI request flags
    SRB_Hdr_Rsvd As Long        '04h/04 Reserved, must = 0
    HA_Count As Byte            '08h/08 Number of host adapters present
    HA_Id As Byte               '09h/09 SCSI ID of host adapter
    HA_MgrId As String * 16     '0ah/10 String describing the manager
    HA_Ident As String * 16     '1ah/26 String describing the host adapter
    HA_Unique(15) As Byte       '2ah/42 Host Adapter Unique parameters
    HA_Rsvd As Integer          '3ah/58 Reserved, must = 0
    HA_Pad(19) As Byte          '3eh/62 padding
End Type

'** SRB - GET DEVICE TYPE - SC_GET_DEV_TYPE (1)
Public Type SRB_GetDevType
    SRB_Cmd As Byte             '00h/00 ASPI command code == SC_HA_INQUIRY
    SRB_Status As Byte          '01h/01 ASPI command status byte
    SRB_HaID As Byte            '02h/02 ASPI host adapter number
    SRB_Flags As Byte           '03h/03 ASPI request flags
    SRB_Hdr_Rsvd As Long        '04h/04 Reserved, must = 0
    SRB_Target As Byte          '08h/08 Target's SCSI ID
    SRB_Lun As Byte             '09h/09 Target's LUN number
    DEV_DeviceType As Byte      '0ah/10 Target's peripheral device type
    DEV_Rsvd1 As Byte           '0bh/11 Reserved, must = 0
    DEV_Pad(67) As Byte         '0ch/12 padding
End Type

'** SRB - EXECUTE SCSI COMMAND - SC_EXEC_SCSI_CMD (2)
Public Type SRB_ExecuteIO
    SRB_Cmd As Byte             '00h/00 ASPI command code == SC_HA_INQUIRY
    SRB_Status As Byte          '01h/01 ASPI command status byte
    SRB_HaID As Byte            '02h/02 ASPI host adapter number
    SRB_Flags As Byte           '03h/03 ASPI request flags
    SRB_Hdr_Rsvd As Long        '04h/04 Reserved, must = 0
    SRB_Target As Byte          '08h/08 Target's SCSI ID
    SRB_Lun As Byte             '09h/09 Target's LUN number
    SRB_Rsvd1 As Integer        '0ah/10 Reserved for alignment
    SRB_BufLen As Long          '0ch/12 Data Allocation Length
    SRB_BufPointer As Long      '10h/16 Data Buffer Pointer
    SRB_SenseLen As Byte        '14h/20 Sense Allocation Length
    SRB_CDBLen As Byte          '15h/21 CDB Length
    SRB_HaStat As Byte          '16h/22 Host Adapter Status
    SRB_TargStat As Byte        '17h/23 Target Status
    SRB_PostProc As Long        '18h/24 Post routine
    SRB_Rsvd2(19) As Byte       '1ch/28 Reserved, must = 0
    SRB_CDBByte(15) As Byte     '30h/48 SCSI CDB
    SRB_SenseData(15) As Byte   '40h/64 Request Sense buffer
End Type

'*******************************************************************
'** PERIPHERAL DEVICE TYPE DEFINITIONS
'*******************************************************************
Public Const DTYPE_DASD = 0         'Disk Device
Public Const DTYPE_SEQD = 1         'Tape Device
Public Const DTYPE_PRNT = 2         'Printer
Public Const DTYPE_PROC = 3         'Processor
Public Const DTYPE_WORM = 4         'Write-once read-multiple
Public Const DTYPE_CROM = 5         'CD-ROM device
Public Const DTYPE_CDROM = 5        'CD-ROM device
Public Const DTYPE_SCAN = 6         'Scanner device
Public Const DTYPE_OPTI = 7         'Optical memory device
Public Const DTYPE_JUKE = 8         'Medium Changer device
Public Const DTYPE_COMM = 9         'Communications device
Public Const DTYPE_RESL = &HA       'Reserved (low)
Public Const DTYPE_RESH = &H1E      'Reserved (high)
Public Const DTYPE_UNKNOWN = &H1F   'Unknown or no device type

'*******************************************************************
'** Misc constants used by SCSI I/O commands
'*******************************************************************
Public Const SENSE_LEN = 14         'Default sense buffer length.
Public Const SRB_DIR_IN = &H8       'Transfer from SCSI target to host.
Public Const SRB_DIR_OUT = &H10     'Transfer from host to SCSI target.
Public Const SRB_POSTING = &H1      'Enable ASPI posting.
Public Const SRB_EVENT_NOTIFY = &H40    'Enable ASPI event notification.
Public Const SRB_ENABLE_RESIDUAL = &H4  'Enable residual byte count reporting.

'*******************************************************************
'** Host Adapter Status Values
'*******************************************************************
Public Const HASTAT_OK = &H0            'Host adapter did not detect an error.
Public Const HASTAT_TIMEOUT = &H9       'Timed out while SRB was waiting to be processed.
Public Const HASTAT_CMD_TIMEOUT = &HB   'While processing the SRB, adapter timed out.
Public Const HASTAT_MSG_REJECT = &HD    'While processing SRB, the adapter received a MESSAGE REJECT.
Public Const HASTAT_BUS_RESET = &HE     'A bus reset was detected.
Public Const HASTAT_PARITY_ERROR = &HF  'A parity error was detected.
Public Const HASTAT_REQ_SENSE_FAIL = &H10 'The adapter failed in issuing REQUEST SENSE.
Public Const HASTAT_SEL_TO = &H11       'Selection Timeout.
Public Const HASTAT_DO_DU = &H12        'Data overrun / data underrun.
Public Const HASTAT_BUS_FREE = &H13     'Unexpected bus free.
Public Const HASTAT_PHASE_ERR = &H14    'Target bus phase sequence failure.

'*******************************************************************
'** Target Status Values
'*******************************************************************
Public Const STATUS_GOOD = &H0          'Status Good.
Public Const STATUS_CHKCOND = &H2       'Check Condition.
Public Const STATUS_CONDMET = &H4       'Condition Met.
Public Const STATUS_BUSY = &H8          'Busy.
Public Const STATUS_INTERM = &H10       'Intermediate.
Public Const STATUS_INTCDMET = &H14     'Intermediate-condition met.
Public Const STATUS_RESCONF = &H18      'Reservation conflict.
Public Const STATUS_CMD_TERM = &H22     'Command Terminated.
Public Const STATUS_QFULL = &H28        'Queue full.

'*******************************************************************
'** Sense Codes
'*******************************************************************
Public Const SENSE_CURRENT = &H70       'Sense data is from current command.
Public Const SENSE_DEFFERED = &H71      'Sense data is from a previous command.

'*******************************************************************
'** Sense Key Values
'*******************************************************************
Public Const KEY_NOSENSE = &H0          'No Sense.
Public Const KEY_RECERROR = &H1         'Recovered Error.
Public Const KEY_NOTREADY = &H2         'Not Ready.
Public Const KEY_MEDIUMERROR = &H3      'Medium Error.
Public Const KEY_HARDERROR = &H4        'Hardware Error.
Public Const KEY_ILLGLREQ = &H5         'Illegal Request.
Public Const KEY_UNITATT = &H6          'Unit Attention.
Public Const KEY_DATAPROT = &H7         'Data Protection.
Public Const KEY_BLANKCHK = &H8         'Blank Check.
Public Const KEY_VENDSPEC = &H9         'Vendor Specific.
Public Const KEY_COPYABORT = &HA        'Copy Aborted.
Public Const KEY_ABORTCMD = &HB         'Aborted Command.
Public Const KEY_EQUAL = &HC            'Equal (Search).
Public Const KEY_VOLOVRFLW = &HD        'Volume Overflow.
Public Const KEY_MISCOMP = &HE          'Miscompare (Search).
Public Const KEY_RSVD = &HF             'Reserved.

'*******************************************************************
'** SCSI Commands for all Device Types
'*******************************************************************
Public Const SCSI_TST_U_RDY = &H0       'Test Unit Ready (Mandatory)
Public Const SCSI_REQ_SENSE = &H3       'Request Sense (Mandatory)
Public Const SCSI_READ = &H8            'Read (Mandatory)
Public Const SCSI_WRITE = &HA           'Write (Mandatory)
Public Const SCSI_INQUIRY = &H12        'Inquiry (Mandatory)
Public Const SCSI_MODE_SEL6 = &H15      'Mode Select 6-byte (Device Specific)
Public Const SCSI_MODE_SEN6 = &H1A      'Mode Sense 6-byte (Device Specific)
Public Const SCSI_MODE_SEL10 = &H55     'Mode Select 10-byte (Device Specific)
Public Const SCSI_MODE_SEN10 = &H5A     'Mode Sense 10-byte (Device Specific)

'*******************************************************************
'** ASPI DLL Declarations
'*******************************************************************
Public Declare Function GetASPI32SupportInfoEx Lib "ASPIshim" _
    () As Long

Public Declare Function SendASPI32InquiryEx Lib "ASPIshim" _
    Alias "SendASPI32CommandEx" (hSRB As SRB_HAInquiry) As Long

Public Declare Function SendASPI32DevTypeEx Lib "ASPIshim" _
    Alias "SendASPI32CommandEx" (hSRB As SRB_GetDevType) As Long

Public Declare Function SendASPI32ExecIOEx Lib "ASPIshim" _
    Alias "SendASPI32CommandEx" (hSRB As SRB_ExecuteIO) As Long
Public Function AspiCheck() As Boolean
    Dim hLoad As Long
    
    'load the error messages to parse...
    hLoad = LoadLibrary("WNASPI32.DLL")

    'check for ASPI driver
    If (GetProcAddress(hLoad, "GetASPI32SupportInfo") <> 0 And _
        GetProcAddress(hLoad, "SendASPI32Command") <> 0) Then
        AspiCheck = True
    End If

    If (hLoad <> 0) Then Call FreeLibrary(hLoad)

End Function
Public Function AspiGetNumAdapters() As Integer
    Dim nRet As Long
    Dim sts As Integer
    Dim cnt As Integer

    'query ASPI for info on transport
    nRet = GetASPI32SupportInfoEx()
    sts = (nRet / 256)
    cnt = nRet And &HF

    If (sts = SS_COMP) Then AspiGetNumAdapters = cnt
End Function

Public Function convDecToBin(ByVal curNumber As Currency) As String
  On Error GoTo convDecToBin_end
  Dim strBin As String
  Dim i As Long

  For i = 64 To 0 Step -1

    If Int(curNumber / (2 ^ i)) = 1 Then

      strBin = strBin & "1"
      curNumber = curNumber - (2 ^ i)

    Else

      If strBin <> "" Then
        strBin = strBin & "0"
      End If

    End If

  Next

  convDecToBin = strBin

convDecToBin_end:
  If Err <> 0 Or strBin = "" Then convDecToBin = "-E-"
  Exit Function
End Function

Public Function ExecCmd(cmdline$)
 Dim proc As PROCESS_INFORMATION
 Dim start As STARTUPINFO
 
      ' Initialize the STARTUPINFO structure:
      start.cb = Len(start)

      ' Start the shelled application:
      nRet = CreateProcessA(vbNullString, cmdline$, 0&, 0&, 1&, _
         NORMAL_PRIORITY_CLASS, 0&, vbNullString, start, proc)

      ' Wait for the shelled application to finish:
         nRet = WaitForSingleObject(proc.hProcess, INFINITE)
         Call GetExitCodeProcess(proc.hProcess, nRet)
         Call CloseHandle(proc.hThread)
         Call CloseHandle(proc.hProcess)
         ExecCmd = nRet
End Function

