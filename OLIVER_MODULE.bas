Attribute VB_Name = "OLIVER_MODULE"
'Function For Parallel Port Inside VBIO.dll located at System32
Public Declare Sub Anjan Lib "vbio.dll" ()
Public Declare Function INP Lib "vbio.dll" Alias "Inp" (ByVal Port&) As Integer
Public Declare Function Inpw Lib "vbio.dll" (ByVal Port&) As Long
Public Declare Sub Out Lib "vbio.dll" (ByVal Port&, ByVal byt%)
Public Declare Sub Outw Lib "vbio.dll" (ByVal Port&, ByVal wrd&)
Public Declare Function GetLptBaseAddr Lib "vbio.dll" (ByVal lpt&) As Integer
Public Declare Function GetComBaseAddr Lib "vbio.dll" (ByVal com&) As Integer

'Function For USB Port Inside VISA32.dll located at System32
Private Declare Function viOpenDefaultRM Lib "VISA32.DLL" Alias "#141" (sesn As Long) As Long
Private Declare Function viFindRsrc Lib "VISA32.DLL" Alias "#129" (ByVal sesn As Long, ByVal expr As String, vi As Long, retCount As Long, ByVal desc As String) As Long
Private Declare Function viFindNext Lib "VISA32.DLL" Alias "#130" (ByVal vi As Long, ByVal desc As String) As Long
Private Declare Function viOpen Lib "VISA32.DLL" Alias "#131" (ByVal sesn As Long, ByVal viDesc As String, ByVal mode As Long, ByVal timeout As Long, vi As Long) As Long
Private Declare Function viClose Lib "VISA32.DLL" Alias "#132" (ByVal vi As Long) As Long
Private Declare Function viVPrintf Lib "VISA32.DLL" Alias "#270" (ByVal vi As Long, ByVal writeFmt As String, params As Any) As Long
Private Declare Function viRead Lib "VISA32.DLL" Alias "#256" (ByVal vi As Long, ByVal Buffer As String, ByVal count As Long, retCount As Long) As Long


'Function For USB Port Inside DIRECT_IO.DLL located at System32
Public Declare Function download Lib "DIRECT_IO" () As Long
Public Declare Function USB_Init Lib "DIRECT_IO" (ByVal Descriptor As String) As Long
Public Declare Sub USB_Close Lib "DIRECT_IO" (ByVal USB_Handle As Long)
Public Declare Function USB_Send Lib "DIRECT_IO" (ByVal USB_Handle As Long, ByRef data As Byte, ByVal numbytes As Long) As Long
Public Declare Function USB_Read Lib "DIRECT_IO" (ByVal USB_Handle As Long) As Byte
Public Declare Function USB_Get_String_Descriptor Lib "DIRECT_IO" (ByVal USB_Handle As Long)
Public Declare Function ReadData Lib "DIRECT_IO" (ByVal i As Byte) As Byte
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const SWP_NOMOVE = 2
      Public Const SWP_NOSIZE = 1
      Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
      Public Const HWND_TOPMOST = -1
      Public Const HWND_NOTOPMOST = -2
 Declare Function SetWindowPos Lib "user32" _
            (ByVal hwnd As Long, _
            ByVal hWndInsertAfter As Long, _
            ByVal x As Long, _
            ByVal y As Long, _
            ByVal cx As Long, _
            ByVal cy As Long, _
            ByVal wFlags As Long) As Long


Global handle As Long

Public Const uA = 1000000
Public Const mA = 1000
Public Const mV = 1000
Public Const NONE = 0
Public Const V = 1
Public Const mS = 1000
Public Const uS = 1000000
Public Const nS = 1000000000


Global DevInst As Long
Global CalInst As Long
Global CalInst2 As Long
Global CalInst3 As Long
Global CalInst4 As Long
Global CalInst5 As Long
Global CalInst6 As Long
Global CalInst7 As Long
Global CalInst8 As Long
Global CalInst9 As Long
Global CalInst10 As Long
Global CalInst11 As Long
Global CalInst12 As Long

Global y As String * 25
Global z As Long
Global col As Long
Global mydata As Single
Global ADDRESSid As Integer
Global lp As Integer
Global lp2 As Integer
Global Myprog As Long
Global PbarStart As Long
Global PbarStop As Long
Global PbarVAL As Long
Global DA As Variant
Const Giga = 1000000000
Const Mega = 1000000
Const Kilo = 1000
Const Milli = 0.001
Const Micro = 0.00001

Global PowerLP As Single
Global PowerLEVEL As Single
Global My128MHz As Single
Global MyPMeter As Single
Global ATT300HZ As Single
Global ATT3KHZ As Single
Global Atten As Single
Global RefLevel As Single

Global FormDet As Integer
Global CalErrLog As Boolean
Global TestExit As Boolean
Global GENSETexit As Boolean
Global Test17_verify As Boolean
Global OpticalPower As Single
Global OpticPow1 As Single
Global sHuntval As Double
'Global SheetName As String

Global MyZeroOffset As Integer
Global SheetName As String
Global CommercialCheck As Integer
Global CommercialLp As Integer

Global MyTempValue
Global MyTempValue2
Global DMMvalue
Global PMaccuracy


Public YScope As Long

' Copyright 1992-2002 Agilent Technologies, Inc.  All Rights Reserved.
'
' This file defines constants, record types, and entry points
' for the Agilent Standard Instrument Control Library.  You need to
' add this file to each Visual BASIC project that uses the
' Agilent Standard Instrument Control Library.

' Name of SICL DLL

Const conSiclDll$ = "SICL32.DLL"

' Support levels:
Global Const I_SICL_REVISION = 40       ' Agilent SICL Revision 4.0
Global Const I_SICL_LEVEL = 3           ' Support Level

' Byte Ordering constants
Global Const I_ORDER_LE = True
Global Const I_ORDER_BE = False

' Session types
Global Const I_SESS_INTF = 1
Global Const I_SESS_DEV = 2
Global Const I_SESS_CMDR = 3

' Interface Types
Global Const I_INTF_NONE = 0
Global Const I_INTF_GPIB = 1
Global Const I_INTF_VXI = 2
Global Const I_INTF_RS232 = 3
Global Const I_INTF_GPIO = 4
' 5 is reserved -- don't use
Global Const I_INTF_USRDEF = 6
' 7 is reserved -- don't use
Global Const I_INTF_MSIB = 8
Global Const I_INTF_LAN = 9

' iread termination conditions
Global Const I_TERM_MAXCNT = 1
Global Const I_TERM_CHR = 2
Global Const I_TERM_END = 4
Global Const I_TERM_NON_BLOCKED = 8

' ixtrig which values.
Global Const I_TRIG_STD = &H1&
Global Const I_TRIG_ALL = &HFFFFFFFF
Global Const I_TRIG_TTL0 = &H1000&
Global Const I_TRIG_TTL1 = &H2000&
Global Const I_TRIG_TTL2 = &H4000&
Global Const I_TRIG_TTL3 = &H8000&
Global Const I_TRIG_TTL4 = &H10000
Global Const I_TRIG_TTL5 = &H20000
Global Const I_TRIG_TTL6 = &H40000
Global Const I_TRIG_TTL7 = &H80000
Global Const I_TRIG_ECL0 = &H100000
Global Const I_TRIG_ECL1 = &H200000
Global Const I_TRIG_ECL2 = &H400000
Global Const I_TRIG_ECL3 = &H800000
Global Const I_TRIG_EXT0 = &H1000000
Global Const I_TRIG_EXT1 = &H2000000
Global Const I_TRIG_EXT2 = &H4000000
Global Const I_TRIG_EXT3 = &H8000000
Global Const I_TRIG_CLK0 = &H10000000
Global Const I_TRIG_CLK1 = &H20000000
Global Const I_TRIG_CLK2 = &H40000000
Global Const I_TRIG_CLK10 = &H80000000
Global Const I_TRIG_CLK100 = &H800&
Global Const I_TRIG_SERIAL_DTR = &H400&
Global Const I_TRIG_SERIAL_RTS = &H200&
Global Const I_TRIG_GPIO_CTL0 = &H100&
Global Const I_TRIG_GPIO_CTL1 = &H80&

' ihint values
Global Const I_HINT_DONTCARE = 0
Global Const I_HINT_USEDMA = 1
Global Const I_HINT_USEPOLL = 2
Global Const I_HINT_USEINTR = 3
Global Const I_HINT_SYSTEM = 4
Global Const I_HINT_IO = 5

' isetintr values.  1-15 are interface independant.
Global Const I_INTR_OFF = 0
Global Const I_INTR_INTFACT = 1
Global Const I_INTR_INTFDEACT = 2
Global Const I_INTR_TRIG = 3
Global Const I_INTR_STB = 4
Global Const I_INTR_DEVCLR = 5

' VXI Interrupts
Global Const I_INTR_VXI_SIGNAL = 16
Global Const I_INTR_VXI_SYSRESET = 17
Global Const I_INTR_VXI_VME = 18
Global Const I_INTR_VXI_LLOCK = 19
Global Const I_INTR_VXI_UKNSIG = 20
Global Const I_INTR_VXI_VMESYSFAIL = 21
Global Const I_INTR_VME_IRQ1 = 22
Global Const I_INTR_VME_IRQ2 = 23
Global Const I_INTR_VME_IRQ3 = 24
Global Const I_INTR_VME_IRQ4 = 25
Global Const I_INTR_VME_IRQ5 = 26
Global Const I_INTR_VME_IRQ6 = 27
Global Const I_INTR_VME_IRQ7 = 28
Global Const I_INTR_ANY_SIG = 29

' GP-IB Interrupts
Global Const I_INTR_GPIB_IFC = 16
Global Const I_INTR_GPIB_PPOLLCONFIG = 17
Global Const I_INTR_GPIB_REMLOC = 18
Global Const I_INTR_GPIB_GET = 20
Global Const I_INTR_GPIB_TLAC = 21

' RS-232 Interrupts
Global Const I_INTR_SERIAL_DAV = 16
Global Const I_INTR_SERIAL_MSL = 17
Global Const I_INTR_SERIAL_BREAK = 18
Global Const I_INTR_SERIAL_ERROR = 19
Global Const I_INTR_SERIAL_TEMT = 20
Global Const I_INTR_SERIAL_MCL = 21

' GP-IO Interrupts
Global Const I_INTR_GPIO_EIR = 16
Global Const I_INTR_GPIO_RDY = 17

' MSIB Interrupts
Global Const I_INTR_MSIB_END_RECEIVED = 22
Global Const I_INTR_MSIB_LINK_BROKEN = 23

' 32 maximum isetintr values
Global Const I_INTR_MAX = 32

' ivxibusstatus values
Global Const I_VXI_BUS_TRIGGER = 0
Global Const I_VXI_BUS_LADDR = 1
Global Const I_VXI_BUS_SERVANT_AREA = 2
Global Const I_VXI_BUS_NORMOP = 3
Global Const I_VXI_BUS_CMDR_LADDR = 4
Global Const I_VXI_BUS_MAN_ID = 5
Global Const I_VXI_BUS_MODEL_ID = 6
Global Const I_VXI_BUS_PROTOCOL = 7
Global Const I_VXI_BUS_XPROT = 8
Global Const I_VXI_BUS_SHM_SIZE = 9
Global Const I_VXI_BUS_SHM_ADDR_SPACE = 10
Global Const I_VXI_SHM_PAGE = 11
Global Const I_VXI_BUS_VXIMXI = 12
Global Const I_VXI_BUS_TRIGSUPP = 13

' igpibbusstatus values
Global Const I_GPIB_BUS_REM = 1
Global Const I_GPIB_BUS_SRQ = 2
Global Const I_GPIB_BUS_NDAC = 3
Global Const I_GPIB_BUS_SYSCTLR = 4
Global Const I_GPIB_BUS_ACTCTLR = 5
Global Const I_GPIB_BUS_TALKER = 6
Global Const I_GPIB_BUS_LISTENER = 7
Global Const I_GPIB_BUS_ADDR = 8
Global Const I_GPIB_BUS_LINES = 9

Global Const I_GPIB_T1DELAY_MIN = 350
Global Const I_GPIB_T1DELAY_MAX = 2400
   
' values for igpioctrl and igpiostat
Global Const I_GPIO_AUX = 1
Global Const I_GPIO_CTRL = 2
Global Const I_GPIO_DATA = 3
Global Const I_GPIO_INFO = 4
Global Const I_GPIO_SET_PCTL = 5
Global Const I_GPIO_STAT = 6
Global Const I_GPIO_READ_EOI = 7
Global Const I_GPIO_TEST_ONLY = 8
Global Const I_GPIO_POLARITY = 9
Global Const I_GPIO_READ_CLK = 10
Global Const I_GPIO_PCTL_DELAY = 11

Global Const I_GPIO_CTRL_CTL0 = &H1
Global Const I_GPIO_CTRL_CTL1 = &H2

Global Const I_GPIO_STAT_STI0 = &H1
Global Const I_GPIO_STAT_STI1 = &H2
Global Const I_GPIO_EIR = &H4
Global Const I_GPIO_PSTS = &H8
Global Const I_GPIO_CHK_PSTS = &H10
Global Const I_GPIO_AUTO_HDSK = &H20
Global Const I_GPIO_ENH_MODE = &H40
Global Const I_GPIO_READY = &H80
Global Const I_GPIO_EOI_NONE = &H10000

' RS-232 values
Global Const I_SERIAL_BAUD = 1
Global Const I_SERIAL_PARITY = 2
Global Const I_SERIAL_STOP = 3
Global Const I_SERIAL_WIDTH = 4
Global Const I_SERIAL_FLOW_CTRL = 5
Global Const I_SERIAL_MSL = 6
Global Const I_SERIAL_STAT = 7
Global Const I_SERIAL_RESET = 9
Global Const I_SERIAL_READ_EOI = 10
Global Const I_SERIAL_WRITE_EOI = 11
Global Const I_SERIAL_DUPLEX = 12
Global Const I_SERIAL_READ_BUFSZ = 13
Global Const I_SERIAL_READ_DAV = 14

' RS-232 duplex modes
Global Const I_SERIAL_DUPLEX_HALF = 1
Global Const I_SERIAL_DUPLEX_FULL = 2

' RS-232 UART status
Global Const I_SERIAL_DAV = &H1
Global Const I_SERIAL_OVERFLOW = &H2
Global Const I_SERIAL_PARERR = &H4
Global Const I_SERIAL_FRAMING = &H8
Global Const I_SERIAL_BREAK = &H10
Global Const I_SERIAL_TEMT = &H20

' RS-232 flow control
Global Const I_SERIAL_FLOW_NONE = 0
Global Const I_SERIAL_FLOW_XON = 1
Global Const I_SERIAL_FLOW_RTS_CTS = 2
Global Const I_SERIAL_FLOW_DTR_DSR = 3

' RS-232 modem status lines
Global Const I_SERIAL_DCD = &H1
Global Const I_SERIAL_DSR = &H2
Global Const I_SERIAL_CTS = &H4
Global Const I_SERIAL_RI = &H8
Global Const I_SERIAL_D_DCD = &H10
Global Const I_SERIAL_D_DSR = &H20
Global Const I_SERIAL_D_CTS = &H40
Global Const I_SERIAL_D_TERI = &H80

' RS-232 modem control lines
Global Const I_SERIAL_RTS = &H1000
Global Const I_SERIAL_DTR = &H2000

' RS-232 parity values
Global Const I_SERIAL_PAR_NONE = 0
Global Const I_SERIAL_PAR_EVEN = 1
Global Const I_SERIAL_PAR_ODD = 2
Global Const I_SERIAL_PAR_MARK = 3
Global Const I_SERIAL_PAR_SPACE = 4
Global Const I_SERIAL_PAR_IGNORE = 5

' RS-232 stop-bit values
Global Const I_SERIAL_STOP_1 = 1
Global Const I_SERIAL_STOP_2 = 2

' RS-232 character width
Global Const I_SERIAL_CHAR_5 = 5
Global Const I_SERIAL_CHAR_6 = 6
Global Const I_SERIAL_CHAR_7 = 7
Global Const I_SERIAL_CHAR_8 = 8

' EOI support (used with the I_SERIAL_*_EOI command)
Global Const I_SERIAL_EOI_CHR = &H100
Global Const I_SERIAL_EOI_NONE = &H200
Global Const I_SERIAL_EOI_BIT8 = &H400


' MSIB error types (for imsibseterror)
Global Const I_MSIB_PERMANENTERR = 0
Global Const I_MSIB_TRANSIENTERR = 1

' MSIB commands (for imsibcmd)
Global Const I_MSIB_CMD_NULL = &H0
Global Const I_MSIB_CMD_END = &H1
Global Const I_MSIB_CMD_SEND_CAPABILITY = &H2
Global Const I_MSIB_CMD_RETURN_TO_LOCAL = &H6
Global Const I_MSIB_CMD_LOCK_LINK = &H7
Global Const I_MSIB_CMD_UNLOCK_LINK = &H8
Global Const I_MSIB_CMD_LIGHT_ACTIVE = &H9
Global Const I_MSIB_CMD_UNLIGHT_ACTIVE = &HA
Global Const I_MSIB_CMD_ERROR_OCCURRED = &HB
Global Const I_MSIB_CMD_ERRORS_CLEARED = &HC
Global Const I_MSIB_CMD_SEND_STATUS = &H10
Global Const I_MSIB_CMD_SEND_ERRORS = &H11
Global Const I_MSIB_CMD_SEND_MODULE_ID = &H12
Global Const I_MSIB_CMD_SEND_MANUFACTURER = &H13
Global Const I_MSIB_CMD_SEND_TIME = &H14
Global Const I_MSIB_CMD_LINK_REMOTE = &H15
Global Const I_MSIB_CMD_LINK_LOCAL = &H16
Global Const I_MSIB_CMD_SEND_MODEL_NUMBER = &H17
Global Const I_MSIB_CMD_SEND_SERIAL_NUMBER = &H18
Global Const I_MSIB_CMD_SEND_FIRMWARE_REV = &H19
Global Const I_MSIB_CMD_STATUS = &H600
Global Const I_MSIB_CMD_SET_IEEE_ADDRESS = &H700

' imap mapspace values
Global Const I_MAP_A16 = &H0
Global Const I_MAP_A24 = &H1
Global Const I_MAP_A32 = &H2
Global Const I_MAP_VXIDEV = &H3
Global Const I_MAP_EXTEND = &H4
Global Const I_MAP_INTFREG = &H5
Global Const I_MAP_SHARED = &H6

' Following is for icmd; uses Radisys define
Global Const DOCMD_VALIDATE_MAPPING = &H40000005

' Error Codes
' NOTE that User Error Codes 32501-32630 are reserved
' for Agilent SICL.
Const SICL_ERR_BASE = 32501

Global Const I_ERR_NOERROR = 0
Global Const I_ERR_SYNTAX = SICL_ERR_BASE
Global Const I_ERR_SYMNAME = 1 + SICL_ERR_BASE
Global Const I_ERR_BADADDR = 2 + SICL_ERR_BASE
Global Const I_ERR_BADID = 3 + SICL_ERR_BASE
Global Const I_ERR_PARAM = 4 + SICL_ERR_BASE
Global Const I_ERR_NOCONN = 5 + SICL_ERR_BASE
Global Const I_ERR_NOPERM = 6 + SICL_ERR_BASE
Global Const I_ERR_NOTSUPP = 7 + SICL_ERR_BASE
Global Const I_ERR_NORSRC = 8 + SICL_ERR_BASE
Global Const I_ERR_NOINTF = 9 + SICL_ERR_BASE
Global Const I_ERR_LOCKED = 10 + SICL_ERR_BASE
Global Const I_ERR_NOLOCK = 11 + SICL_ERR_BASE
Global Const I_ERR_BADFMT = 12 + SICL_ERR_BASE
Global Const I_ERR_DATA = 13 + SICL_ERR_BASE
Global Const I_ERR_TIMEOUT = 14 + SICL_ERR_BASE
Global Const I_ERR_OVERFLOW = 15 + SICL_ERR_BASE
Global Const I_ERR_IO = 16 + SICL_ERR_BASE
Global Const I_ERR_OS = 17 + SICL_ERR_BASE
Global Const I_ERR_BADMAP = 18 + SICL_ERR_BASE
Global Const I_ERR_NODEV = 19 + SICL_ERR_BASE
Global Const I_ERR_INVLADDR = 20 + SICL_ERR_BASE
Global Const I_ERR_NOTIMPL = 21 + SICL_ERR_BASE
Global Const I_ERR_ABORTED = 22 + SICL_ERR_BASE
Global Const I_ERR_BADCONFIG = 23 + SICL_ERR_BASE
Global Const I_ERR_NOCMDR = 24 + SICL_ERR_BASE
Global Const I_ERR_VERSION = 25 + SICL_ERR_BASE
Global Const I_ERR_NESTEDIO = 26 + SICL_ERR_BASE
Global Const I_ERR_BUSY = 27 + SICL_ERR_BASE
Global Const I_ERR_CONNEXISTS = 28 + SICL_ERR_BASE
Global Const I_ERR_BUSERR = 29 + SICL_ERR_BASE
Global Const I_ERR_BUSERR_RETRY = 30 + SICL_ERR_BASE
Global Const I_ERR_INTERNAL = 127 + SICL_ERR_BASE
Global Const I_ERR_INTERRUPT = 128 + SICL_ERR_BASE
Global Const I_ERR_UNKNOWNERR = 129 + SICL_ERR_BASE
Global Const SICL_ERR_LAST = I_ERR_UNKNOWNERR

Global Const I_READ_BUF_SZ = 4096
Global Const I_WRITE_BUF_SZ = 128

Global Const I_BUF_READ = &H1
Global Const I_BUF_WRITE = &H2
Global Const I_BUF_DISCARD_READ = &H4
Global Const I_BUF_DISCARD_WRITE = &H8
Global Const I_BUF_WRITE_END = &H10

' Data Types used by SICL
Type lu_info
  logical_unit As Long
  symname As String * 32
  cardname As String * 32
  filler As Long
  intfType As Long
  location As Long
  busaddr As Long
  hwarg(0 To 15) As String * 20
  visaname As String * 32
  filler2(0 To 3) As Long
End Type

Type vxiinfo
  laddr As Integer
  name As String * 16
  manuf_name As String * 16
  model_name As String * 16
  man_id As Long
  model As Long
  devclass As Long
  selftest As Integer
  cage_num As Integer
  slot As Integer
  protocol As Long
  x_protocol As Long
  servant_area As Long
  addrspace As Long
  memSize As Long
  memstart As Long
  slot0_laddr As Integer
  cmdr_laddr As Integer
  int_handler(0 To 7)  As Integer
  interrupter(0 To 7) As Integer
  fill(0 To 9) As Integer
End Type

' Version Information
'Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function vb_iversion Lib "vbsicl32.dll" (specversion As Integer, implversion As Integer) As Integer
Declare Function vb_idrvrversion Lib "vbsicl32.dll" (ByVal id As Integer, specversion As Integer, implversion As Integer) As Integer

' Open/Close
Declare Function vb_iopen Lib "vbsicl32.dll" (ByVal addr As String) As Integer
Declare Function vb_iclose Lib "vbsicl32.dll" (ByVal id As Integer) As Integer
Declare Function vb_igetintfsess Lib "vbsicl32.dll" (ByVal id As Integer) As Integer

' Write/Read

Declare Function vb_iwrite Lib "vbsicl32.dll" (ByVal which As Integer, ByVal id As Integer, ByVal buf As Variant, ByVal datalen As Long, ByVal endi As Integer, actual As Long) As Integer
Declare Function vb_iread Lib "vbsicl32.dll" (ByVal which As Integer, ByVal id As Integer, buf As Variant, ByVal bufSize As Long, reason As Integer, actual As Long) As Integer
Declare Function vb_itermchr Lib "vbsicl32.dll" (ByVal id As Integer, ByVal tchr As Integer) As Integer
Declare Function vb_igettermchr Lib "vbsicl32.dll" (ByVal id As Integer, tchr As Integer) As Integer

' Formatted I/O
Declare Function vb_iscan Lib "vbsicl32.dll" (ByVal which As Integer, ByVal id As Integer, ByVal s As String, ByVal fmt As String, ByRef va1 As Variant, ByRef va2 As Variant, ByRef va3 As Variant, ByRef va4 As Variant, ByRef va5 As Variant, ByRef va6 As Variant, ByRef va7 As Variant, ByRef va8 As Variant, ByRef va9 As Variant, ByRef va10 As Variant) As Integer
Declare Function vb_iprint Lib "vbsicl32.dll" (ByVal which As Integer, ByVal id As Integer, ByVal s As String, ByVal fmt As String, ByRef ap() As Variant) As Integer

Declare Function vb_ivprintf Lib "vbsicl32.dll" (ByVal id As Integer, ByVal fmt As String, ByVal ap As Variant, ByVal lenBstr As Long) As Integer
Declare Function vb_ivscanf Lib "vbsicl32.dll" (ByVal id As Integer, ByVal fmt As String, ByRef ap As Variant, ByVal lenBstr As Long) As Integer
Declare Function vb_iflush Lib "vbsicl32.dll" (ByVal id As Integer, ByVal mask As Integer) As Integer
Declare Function vb_isetbuf Lib "vbsicl32.dll" (ByVal id As Integer, ByVal mask As Integer, ByVal size As Integer) As Integer

' Device/Interface Control
Declare Function vb_iclear Lib "vbsicl32.dll" (ByVal id As Integer) As Integer
Declare Function vb_ilocal Lib "vbsicl32.dll" (ByVal id As Integer) As Integer
Declare Function vb_iremote Lib "vbsicl32.dll" (ByVal id As Integer) As Integer
Declare Function vb_ireadstb Lib "vbsicl32.dll" (ByVal id As Integer, ByRef stb As Integer) As Integer
Declare Function vb_itrigger Lib "vbsicl32.dll" (ByVal id As Integer) As Integer
Declare Function vb_ixtrig Lib "vbsicl32.dll" (ByVal id As Integer, ByVal which As Long) As Integer
Declare Function vb_ihint Lib "vbsicl32.dll" (ByVal id As Integer, ByVal hint As Integer) As Integer

' Commander Sessions
Declare Function vb_isetstb Lib "vbsicl32.dll" (ByVal id As Integer, ByVal stb As Byte) As Integer

' Locking
Declare Function vb_ilock Lib "vbsicl32.dll" (ByVal id As Integer) As Integer
Declare Function vb_iunlock Lib "vbsicl32.dll" (ByVal id As Integer) As Integer
Declare Function vb_isetlockwait Lib "vbsicl32.dll" (ByVal id As Integer, ByVal flag As Integer) As Integer
Declare Function vb_igetlockwait Lib "vbsicl32.dll" (ByVal id As Integer, flag As Integer) As Integer

' Timeouts
Declare Function vb_itimeout Lib "vbsicl32.dll" (ByVal id As Integer, ByVal tval As Long) As Integer
Declare Function vb_igettimeout Lib "vbsicl32.dll" (ByVal id As Integer, tval As Long) As Integer

' Misc routines
Declare Function vb_igetaddr Lib "vbsicl32.dll" (ByVal id As Integer, ByVal addr As String) As Integer
Declare Function vb_igetintftype Lib "vbsicl32.dll" (ByVal id As Integer, pdata As Integer) As Integer
Declare Function vb_igetsesstype Lib "vbsicl32.dll" (ByVal id As Integer, pdata As Integer) As Integer
Declare Function vb_igetdevaddr Lib "vbsicl32.dll" (ByVal id As Integer, prim As Integer, sec As Integer) As Integer
Declare Function vb_igetlu Lib "vbsicl32.dll" (ByVal id As Integer, lu As Integer) As Integer
Declare Function vb_iswap Lib "vbsicl32.dll" (ByRef addr As Variant, ByVal length As Long, ByVal datasize As Integer) As Integer
Declare Function vb_igetlulist Lib "vbsicl32.dll" (list() As Integer) As Integer
Declare Function vb_igetluinfo Lib "vbsicl32.dll" (ByVal lu As Integer, result As lu_info) As Integer
Declare Function vb_igetgatewaytype Lib "vbsicl32.dll" (ByVal id As Integer, pdata As Integer) As Integer


' Error Handling
Declare Function vb_igeterrno Lib "vbsicl32.dll" () As Integer
Declare Function vb_iseterrno Lib "vbsicl32.dll" (ByVal id As Integer, ByVal xint As Integer) As Integer
Declare Function vb_igeterrstr Lib "vbsicl32.dll" (ByVal errcode As Integer, ByVal myerrstr As String) As Integer
Declare Function vb_icauseerr Lib "vbsicl32.dll" (ByVal id As Integer, ByVal errcode As Integer, ByVal flag As Integer) As Integer
Declare Function vbsetsiclerrbase Lib "vbsicl32.dll" (ByVal errbase As Integer) As Integer

' RS-232 specific routines
Declare Function vb_iserialmclctrl Lib "vbsicl32.dll" (ByVal id As Integer, ByVal sline As Integer, ByVal state As Integer) As Integer
Declare Function vb_iserialmclstat Lib "vbsicl32.dll" (ByVal id As Integer, ByVal sline As Integer, state As Integer) As Integer
Declare Function vb_iserialctrl Lib "vbsicl32.dll" (ByVal id As Integer, ByVal request As Integer, ByVal setting As Long) As Integer
Declare Function vb_iserialstat Lib "vbsicl32.dll" (ByVal id As Integer, ByVal request As Integer, result As Long) As Integer
Declare Function vb_iserialbreak Lib "vbsicl32.dll" (ByVal id As Integer) As Integer

' VXI Specific routines
Declare Function vb_ivxibusstatus Lib "vbsicl32.dll" (ByVal id As Integer, ByVal request As Integer, result As Long) As Integer
Declare Function vb_ivxiwaitnormop Lib "vbsicl32.dll" (ByVal id As Integer) As Integer
Declare Function vb_ivxitrigon Lib "vbsicl32.dll" (ByVal id As Integer, ByVal which As Long) As Integer
Declare Function vb_ivxitrigoff Lib "vbsicl32.dll" (ByVal id As Integer, ByVal which As Long) As Integer
Declare Function vb_ivxitrigroute Lib "vbsicl32.dll" (ByVal id As Integer, ByVal in_which As Long, ByVal out_which As Long) As Integer
Declare Function vb_ivxigettrigroute Lib "vbsicl32.dll" (ByVal id As Integer, ByVal which As Long, route As Long) As Integer
Declare Function vb_ivxiws Lib "vbsicl32.dll" (ByVal id As Integer, ByVal wscmd As Integer, wsresp As Integer, rpe As Integer) As Integer
Declare Function vb_ivxiservants Lib "vbsicl32.dll" (ByVal id As Integer, ByVal maxnum As Integer, list() As Integer) As Integer
Declare Function vb_ivxirminfo Lib "vbsicl32.dll" (ByVal id As Integer, ByVal laddr As Integer, ByRef info As vxiinfo) As Integer

' GP-IB Specific Details
Declare Function vb_igpibbusstatus Lib "vbsicl32.dll" (ByVal id As Integer, ByVal request As Integer, result As Integer) As Integer
Declare Function vb_igpibppoll Lib "vbsicl32.dll" (ByVal id As Integer, result As Integer) As Integer
Declare Function vb_igpibppollconfig Lib "vbsicl32.dll" (ByVal id As Integer, ByVal cval As Integer) As Integer
Declare Function vb_igpibppollresp Lib "vbsicl32.dll" (ByVal id As Integer, ByVal sval As Integer) As Integer
Declare Function vb_igpibpassctl Lib "vbsicl32.dll" (ByVal id As Integer, ByVal busaddr As Integer) As Integer
Declare Function vb_igpibrenctl Lib "vbsicl32.dll" (ByVal id As Integer, ByVal ren As Integer) As Integer
Declare Function vb_igpibatnctl Lib "vbsicl32.dll" (ByVal id As Integer, ByVal atnval As Integer) As Integer
Declare Function vb_igpibsendcmd Lib "vbsicl32.dll" (ByVal id As Integer, ByVal buf As String, ByVal length As Integer) As Integer
Declare Function vb_igpibllo Lib "vbsicl32.dll" (ByVal id As Integer) As Integer
Declare Function vb_igpibbusaddr Lib "vbsicl32.dll" (ByVal id As Integer, ByVal busaddr As Integer) As Integer
Declare Function vb_igpibgett1delay Lib "vbsicl32.dll" (ByVal id As Integer, delay As Integer) As Integer
Declare Function vb_igpibsett1delay Lib "vbsicl32.dll" (ByVal id As Integer, ByVal delay As Integer) As Integer
Declare Function vb_igpibpulseifc Lib "vbsicl32.dll" (ByVal id As Integer) As Integer

' GPIO Specific routines
Declare Function vb_igpioctrl Lib "vbsicl32.dll" (ByVal id As Integer, ByVal request As Integer, ByVal setting As Long) As Integer
Declare Function vb_igpiostat Lib "vbsicl32.dll" (ByVal id As Integer, ByVal request As Integer, ByRef result As Long) As Integer
Declare Function vb_igpiosetwidth Lib "vbsicl32.dll" (ByVal id As Integer, ByVal dwidth As Integer) As Integer
Declare Function vb_igpiogetwidth Lib "vbsicl32.dll" (ByVal id As Integer, ByRef dwidth As Integer) As Integer

' LAN Specific functions
Declare Function vb_ilantimeout Lib "vbsicl32.dll" (ByVal id As Integer, ByVal tval As Long) As Integer
Declare Function vb_ilangettimeout Lib "vbsicl32.dll" (ByVal id As Integer, tval As Long) As Integer

' Map routines
Declare Function vb_imap Lib "vbsicl32.dll" (ByVal id As Integer, ByVal mapSpace As Integer, ByVal pagestart As Integer, ByVal pagecnt As Integer, ByVal suggested As Long) As Long
Declare Function vb_iunmap Lib "vbsicl32.dll" (ByVal id As Integer, ByVal addr As Long, ByVal mapSpace As Integer, ByVal pagestart As Integer, ByVal pagecnt As Integer) As Integer
Declare Function vb_imapx Lib "vbsicl32.dll" (ByVal id As Integer, ByVal mapSpace As Integer, ByVal pagestart As Integer, ByVal pagecnt As Integer) As Long
Declare Function vb_iunmapx Lib "vbsicl32.dll" (ByVal id As Integer, ByVal addr As Long, ByVal mapSpace As Integer, ByVal pagestart As Integer, ByVal pagecnt As Integer) As Integer
Declare Function vb_imapinfo Lib "vbsicl32.dll" (ByVal id As Integer, ByVal mapSpace As Integer, numwindows As Integer, winsize As Integer) As Integer

' peekx/pokex/blockmovex routines
Declare Function vb_ipokex8 Lib "vbsicl32.dll" (ByVal id As Integer, ByVal handle As Long, ByVal Offset As Long, ByVal value As Byte) As Integer
Declare Function vb_ipokex16 Lib "vbsicl32.dll" (ByVal id As Integer, ByVal handle As Long, ByVal Offset As Long, ByVal value As Integer) As Integer
Declare Function vb_ipokex32 Lib "vbsicl32.dll" (ByVal id As Integer, ByVal handle As Long, ByVal Offset As Long, ByVal value As Long) As Integer
Declare Function vb_ipeekx8 Lib "vbsicl32.dll" (ByVal id As Integer, ByVal handle As Long, ByVal Offset As Long, value As Byte) As Integer
Declare Function vb_ipeekx16 Lib "vbsicl32.dll" (ByVal id As Integer, ByVal handle As Long, ByVal Offset As Long, value As Integer) As Integer
Declare Function vb_ipeekx32 Lib "vbsicl32.dll" (ByVal id As Integer, ByVal handle As Long, ByVal Offset As Long, value As Long) As Integer
Declare Function vb_iblockmovex Lib "vbsicl32.dll" (ByVal id As Integer, ByVal srcHandle As Long, ByRef srcOffset As Variant, ByVal srcWidth As Integer, ByVal srcIncrement As Integer, ByVal destHandle As Long, ByRef destOffset As Variant, ByVal destWidth As Integer, ByVal destIncrement As Integer, ByVal count As Long, ByVal swap As Integer) As Integer

' Block copy and fifo routines
Declare Function vb_ibblockcopy Lib "vbsicl32.dll" (ByVal id As Integer, ByVal src As Long, ByVal dest As Long, ByVal CNT As Long) As Integer
Declare Function vb_iwblockcopy Lib "vbsicl32.dll" (ByVal id As Integer, ByVal src As Long, ByVal dest As Long, ByVal CNT As Long, ByVal swap As Integer) As Integer
Declare Function vb_ilblockcopy Lib "vbsicl32.dll" (ByVal id As Integer, ByVal src As Long, ByVal dest As Long, ByVal CNT As Long, ByVal swap As Integer) As Integer
Declare Function vb_ibpushfifo Lib "vbsicl32.dll" (ByVal id As Integer, ByVal src As Long, ByVal fifo As Long, ByVal CNT As Long) As Integer
Declare Function vb_iwpushfifo Lib "vbsicl32.dll" (ByVal id As Integer, ByVal src As Long, ByVal fifo As Long, ByVal CNT As Long, ByVal swap As Integer) As Integer
Declare Function vb_ilpushfifo Lib "vbsicl32.dll" (ByVal id As Integer, ByVal src As Long, ByVal fifo As Long, ByVal CNT As Long, ByVal swap As Integer) As Integer
Declare Function vb_ibpopfifo Lib "vbsicl32.dll" (ByVal id As Integer, ByVal fifo As Long, ByVal dest As Long, ByVal CNT As Long) As Integer
Declare Function vb_iwpopfifo Lib "vbsicl32.dll" (ByVal id As Integer, ByVal fifo As Long, ByVal dest As Long, ByVal CNT As Long, ByVal swap As Integer) As Integer
Declare Function vb_ilpopfifo Lib "vbsicl32.dll" (ByVal id As Integer, ByVal fifo As Long, ByVal dest As Long, ByVal CNT As Long, ByVal swap As Integer) As Integer
Declare Function vb_icmd Lib "vbsicl32.dll" (ByVal id As Integer, ByVal cmd As Long, ByVal datalen As Integer, ByVal datawidth As Integer, ByRef pdata As Long) As Integer

' Windows 3.1 Cleanup routines
Declare Function vb__siclcleanup Lib "vbsicl32.dll" () As Integer

' Windows 3.1 yield control routine
Declare Function vb__setsiclyield Lib "vbsicl32.dll" (ByVal yield_option As Integer) As Integer

' Peek/Poke routines
Declare Sub vb_ibpoke Lib "vbsicl32.dll" (ByVal addr As Long, ByVal value As Byte)
Declare Sub vb_iwpoke Lib "vbsicl32.dll" (ByVal addr As Long, ByVal value As Integer)
Declare Sub vb_ilpoke Lib "vbsicl32.dll" (ByVal addr As Long, ByVal value As Long)
Declare Function vb_ibpeek Lib "vbsicl32.dll" (ByVal addr As Long) As Byte
Declare Function vb_iwpeek Lib "vbsicl32.dll" (ByVal addr As Long) As Integer
Declare Function vb_ilpeek Lib "vbsicl32.dll" (ByVal addr As Long) As Long

Function iversion(specversion As Integer, implversion As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iversion(specversion, implversion)
    iversion = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function idrvrversion(id1 As Integer, specversion As Integer, implversion As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_idrvrversion(id1, specversion, implversion)
    idrvrversion = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function iopen(siclAddr As String) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iopen(siclAddr)
    iopen = id

    ' If we get 0 back, there was an error, try to report it
    If id = 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            'Err.Raise (thisErrno) 'Raise the error
            If thisErrno = 32503 Then
            MsgBox ("Equipment Address Error." & Chr(10) & _
            "Make sure that the equipment is properly connected " & Chr(10) & _
            "and the Address is properly set."), vbCritical
            CalErrLog = True
            Form1.Show
            End If
        End If
    End If

End Function



Function iclose(ByVal id1 As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iclose(id1)
    iclose = id
    
    ' If return value was not 0, we had an error
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function igetintfsess(ByVal id1 As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igetintfsess(id1)
    igetintfsess = id

    ' If we get 0 back, there was an error, try to report it
    If id = 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function iwrite(ByVal id1 As Integer, ByVal buf As Variant, ByVal datalen As Long, ByVal endi As Integer, actual As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    tmp = VarType(buf)

    ' Call the function in the SICL DLL and check for errors
    id = vb_iwrite(1, id1, buf, datalen, endi, actual)
    
    iwrite = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ifwrite(ByVal id1 As Integer, ByVal buf As Variant, ByVal datalen As Long, ByVal endi As Integer, actual As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    tmp = VarType(buf)

    ' Call the function in the SICL DLL and check for errors
    id = vb_iwrite(2, id1, buf, datalen, endi, actual)
    
    ifwrite = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function iread(ByVal id1 As Integer, ByRef buf As Variant, ByVal bufSize As Long, reason As Integer, actual As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer
        
    tmp = VarType(buf)

    ' Call the function in the SICL DLL and check for errors
    id = vb_iread(1, id1, buf, bufSize, reason, actual)
    
    iread = id

    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ifread(ByVal id1 As Integer, ByRef buf As Variant, ByVal bufSize As Long, reason As Integer, actual As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer
        
    tmp = VarType(buf)

    ' Call the function in the SICL DLL and check for errors
    id = vb_iread(2, id1, buf, bufSize, reason, actual)
    
    ifread = id

    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function itermchr(ByVal id1 As Integer, ByVal tchr As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_itermchr(id1, tchr)
    itermchr = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function igettermchr(ByVal id1 As Integer, tchr As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igettermchr(id1, tchr)
    igettermchr = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function
Public Sub ActLock(Lock1 As Integer, Lock2 As Integer, lock3 As Integer, lock3dcv As Integer, Lock4 As Integer, Lock5 As Integer, lock5lin As Integer, Lock6 As Integer, _
Lock7 As Integer, Lock8 As Integer, Lock9 As Integer, Lock10 As Integer, Lock11 As Integer, Lock12 As Integer, Lock13 As Integer, _
Lock14 As Integer, Lockall As Integer) '''remove the lock3 for Rear, lock15
If Lock1 = 1 Then Form1.Test1.Enabled = True
If Lock2 = 1 Then Form1.Test2.Enabled = True
If lock3 = 1 Then Form1.Test3.Enabled = True
If lock3dcv = 1 Then Form1.Test3DCV.Enabled = True

If Lock4 = 1 Then Form1.Test4.Enabled = True
If Lock5 = 1 Then Form1.Test5.Enabled = True
If lock5lin = 1 Then Form1.Test5LIN.Enabled = True
If Lock6 = 1 Then Form1.Test6.Enabled = True
If Lock7 = 1 Then Form1.Test7.Enabled = True
If Lock8 = 1 Then Form1.Test8.Enabled = True
If Lock9 = 1 Then Form1.Test9(1).Enabled = True
If Lock10 = 1 Then Form1.Test10(1).Enabled = True
If Lock11 = 1 Then Form1.Test11(1).Enabled = True
If Lockall = 1 Then Form1.TestAll.Enabled = True
If Lock12 = 1 Then Form1.Test12.Enabled = True
If Lock13 = 1 Then Form1.Test2DCV.Enabled = True
If Lock14 = 1 Then Form1.Test2OHM.Enabled = True
''If Lock15 = 1 Then Form1.Test3DCV.Enabled = True

If Lock1 = 0 Then Form1.Test1.Enabled = False
If Lock2 = 0 Then Form1.Test2.Enabled = False
If lock3 = 0 Then Form1.Test3.Enabled = False
If lock3dcv = 0 Then Form1.Test3DCV.Enabled = False
If Lock4 = 0 Then Form1.Test4.Enabled = False
If Lock5 = 0 Then Form1.Test5.Enabled = False
If lock5lin = 0 Then Form1.Test5LIN.Enabled = False
If Lock6 = 0 Then Form1.Test6.Enabled = False
If Lock7 = 0 Then Form1.Test7.Enabled = False
If Lock8 = 0 Then Form1.Test8.Enabled = False
If Lock9 = 0 Then Form1.Test9(1).Enabled = False
If Lock10 = 0 Then Form1.Test10(1).Enabled = False
If Lock11 = 0 Then Form1.Test11(1).Enabled = False
If Lockall = 0 Then Form1.TestAll.Enabled = False
If Lock12 = 0 Then Form1.Test12.Enabled = False
If Lock13 = 0 Then Form1.Test2DCV.Enabled = False
If Lock14 = 0 Then Form1.Test2OHM.Enabled = False
''If Lock15 = 0 Then Form1.Test3DCV.Enabled = False

End Sub



Function ivprintf(ByVal id1 As Integer, ByVal fmt As String, Optional ByVal ap As Variant) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer
    Dim howLong As Long

    ' Force no error to be condition
    Call vb_icauseerr(id1, 0, 0)

    
    If VarType(ap) = 8 Then
       howLong = Len(ap)
    Else
       howLong = 0
    End If
      
    ' Call the function in the SICL DLL and check for errors
    If IsMissing(ap) Then
       id = vb_ivprintf(ByVal id1, ByVal fmt, ByVal sEmpty, ByVal howLong)
    ElseIf IsEmpty(ap) Then
       id = vb_ivprintf(ByVal id1, ByVal fmt, ByVal sEmpty, ByVal howLong)
    Else
       id = vb_ivprintf(ByVal id1, ByVal fmt, ByVal ap, ByVal howLong)
    End If
    
    ivprintf = id

    thisErrno = vb_igeterrno()
    
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            'Err.Raise (thisErrno) 'Raise the error
            If thisErrno = 32503 Then
            MsgBox ("Equipment Address Error." & Chr(10) & _
            "Make sure that the equipment is properly connected " & Chr(10) & _
            "and the Address is properly set."), vbCritical
            CalErrLog = True
            Form1.Show
            End If
            
            If thisErrno = 32510 Then
            MsgBox ("Equipment Address Error." & Chr(10) & _
            "Make sure that correct equipment is connected " & Chr(10) & _
            "and the Address is properly set."), vbCritical
            CalErrLog = True
            End If
        End If
    
'    If thisErrno <> 0 Then
'        Err.Clear    ' set default values in the error object
'        ' set the error string and raise the error
'        tmp = vb_igeterrstr(thisErrno, myerrstr)
'        Err.Description = myerrstr
'        Err.Raise (thisErrno) 'Raise the error
'    End If





End Function


Function ivscanf(ByVal id1 As Integer, ByVal fmt As String, ByRef myVal As Variant) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim howLong As Long
    Dim myerrstr As String * 60
    Dim tmp As Integer
    Dim returnStr As String
        
    ' Put anything in the local string to make non-null
    returnStr = "aa"
                   
    ' Force no error to be condition
    Call vb_icauseerr(id1, 0, 0)

    tmp = VarType(myVal)

    If VarType(myVal) = 8 Then
       howLong = Len(myVal)
    Else
       howLong = 0
    End If

    ' Call the function in the SICL DLL and check for errors
    If tmp = 8 Then
       id = vb_ivscanf(id1, fmt, returnStr, howLong)

       'Place scanf value into myVal
       myVal = returnStr
    Else
       id = vb_ivscanf(id1, fmt, myVal, howLong)
    End If

        
    ivscanf = id
        
    thisErrno = vb_igeterrno()
    If thisErrno <> 0 Then
        myerrstr = ""
        Err.Clear    ' set default values in the error object
        ' set the error string and raise the error
        tmp = vb_igeterrstr(thisErrno, myerrstr)
        Err.Description = myerrstr
        MsgBox myerrstr
        Err.Raise (thisErrno) 'Raise the error
    End If

End Function


Function iflush(ByVal id1 As Integer, ByVal mask As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iflush(id1, mask)
    iflush = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function isetbuf(ByVal id1 As Integer, ByVal mask As Integer, ByVal size As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_isetbuf(id1, mask, size)
    isetbuf = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function iclear(ByVal id1 As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iclear(id1)
    iclear = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ilocal(ByVal id1 As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ilocal(id1)
    ilocal = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function iremote(ByVal id1 As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iremote(id1)
    iremote = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function ireadstb(ByVal id1 As Integer, ByRef stb As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ireadstb(id1, stb)
    ireadstb = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function itrigger(ByVal id1 As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_itrigger(id1)
    itrigger = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function ixtrig(ByVal id1 As Integer, ByVal which As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ixtrig(id1, which)
    ixtrig = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function ihint(ByVal id1 As Integer, ByVal hint As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ihint(id1, hint)
    ihint = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function isetstb(ByVal id1 As Integer, ByVal stb As Byte) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_isetstb(id1, stb)
    isetstb = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function ilock(ByVal id1 As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ilock(id1)
    ilock = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function iunlock(ByVal id1 As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iunlock(id1)
    iunlock = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function isetlockwait(ByVal id1 As Integer, ByVal flag As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_isetlockwait(id1, flag)
    isetlockwait = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function igetlockwait(ByVal id1 As Integer, flag As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igetlockwait(id1, flag)
    igetlockwait = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function itimeout(ByVal id1 As Integer, ByVal tval As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_itimeout(id1, tval)
    itimeout = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function igettimeout(ByVal id1 As Integer, tval As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igettimeout(id1, tval)
    igettimeout = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function igetaddr(ByVal id1 As Integer, ByRef addr As String) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igetaddr(id1, addr)
    igetaddr = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function igetintftype(ByVal id1 As Integer, pdata As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igetintftype(id1, pdata)
    igetintftype = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function igetsesstype(ByVal id1 As Integer, pdata As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igetsesstype(id1, pdata)
    igetsesstype = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function igetdevaddr(ByVal id1 As Integer, prim As Integer, sec As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igetdevaddr(id1, prim, sec)
    igetdevaddr = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function igetlu(ByVal id1 As Integer, lu As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igetlu(id1, lu)
    igetlu = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function ibeswap(addr As Variant, ByVal length As Long, ByVal datasize As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iswap(addr, length, datasize)
    ibeswap = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function ileswap(addr As Variant, ByVal length As Long, ByVal datasize As Integer) As Integer
   ' We are already LE, so no swapping necesary...
   ileswap = I_ERR_NOERROR
End Function


Function iswap(addr As Variant, ByVal length As Long, ByVal datasize As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iswap(addr, length, datasize)
    iswap = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function igetlulist(list() As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igetlulist(list)
    igetlulist = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function igetluinfo(ByVal lu As Integer, result As lu_info) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer
    Dim tempLu As lu_info

    tempLu.hwarg(0) = "abc0"
    tempLu.hwarg(1) = "efg1"
    tempLu.hwarg(2) = "ijk2"

    ' Call the function in the SICL DLL and check for errors
    id = vb_igetluinfo(lu, tempLu)
    igetluinfo = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    Else
       'No error, so copy data to struct
       result = tempLu
    End If

End Function


Function igetgatewaytype(ByVal id1 As Integer, pdata As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igetgatewaytype(id1, pdata)
    igetgatewaytype = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function iserialmclstat(ByVal id1 As Integer, ByVal sline As Integer, state As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iserialmclstat(id1, sline, state)
    iserialmclstat = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function iserialmclctrl(ByVal id1 As Integer, ByVal sline As Integer, ByVal state As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iserialmclctrl(id1, sline, state)
    iserialmclctrl = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function iserialctrl(ByVal id1 As Integer, ByVal request As Integer, ByVal setting As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iserialctrl(id1, request, setting)
    iserialctrl = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function iserialstat(ByVal id1 As Integer, ByVal request As Integer, result As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iserialstat(id1, request, result)
    iserialstat = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function iserialbreak(ByVal id1 As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iserialbreak(id)
    iserialbreak = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ivxibusstatus(ByVal id1 As Integer, ByVal request As Integer, result As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ivxibusstatus(id1, request, result)
    ivxibusstatus = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ivxiwaitnormop(ByVal id1 As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ivxiwaitnormop(id)
    ivxiwaitnormop = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ivxitrigon(ByVal id1 As Integer, ByVal which As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ivxitrigon(id1, which)
    ivxitrigon = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ivxitrigoff(ByVal id1 As Integer, ByVal which As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ivxitrigoff(id1, which)
    ivxitrigoff = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ivxitrigroute(ByVal id1 As Integer, ByVal in_which As Long, ByVal out_which As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ivxitrigroute(id1, in_which, out_which)
    ivxitrigroute = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ivxigettrigroute(ByVal id1 As Integer, ByVal which As Long, route As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ivxigettrigroute(id1, which, route)
    ivxigettrigroute = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ivxiws(ByVal id1 As Integer, ByVal wscmd As Integer, wsresp As Integer, rpe As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ivxiws(id1, wscmd, wsresp, rpe)
    ivxiws = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ivxiservants(ByVal id1 As Integer, ByVal maxnum As Integer, list() As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ivxiservants(id1, maxnum, list)
    ivxiservants = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ivxirminfo(ByVal id1 As Integer, ByVal laddr As Integer, ByRef info As vxiinfo) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ivxirminfo(id1, laddr, info)
    ivxirminfo = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function igpibbusstatus(ByVal id1 As Integer, ByVal request As Integer, result As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpibbusstatus(id1, request, result)
    igpibbusstatus = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function igpibppoll(ByVal id1 As Integer, result As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpibppoll(id1, result)
    igpibppoll = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function igpibppollconfig(ByVal id1 As Integer, ByVal cval As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpibppollconfig(id1, cval)
    igpibppollconfig = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function igpibppollresp(ByVal id1 As Integer, ByVal sval As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpibppollresp(id1, sval)
    igpibppollresp = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function igpibpassctl(ByVal id1 As Integer, ByVal busaddr As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpibpassctl(id1, busaddr)
    igpibpassctl = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function igpibrenctl(ByVal id1 As Integer, ByVal ren As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpibrenctl(id1, ren)
    igpibrenctl = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function igpibatnctl(ByVal id1 As Integer, ByVal atnval As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpibatnctl(id1, atnval)
    igpibatnctl = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function igpibsendcmd(ByVal id1 As Integer, ByVal buf As String, ByVal length As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpibsendcmd(id1, buf, length)
    igpibsendcmd = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function igpibllo(ByVal id1 As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpibllo(id)
    igpibllo = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function igpibbusaddr(ByVal id1 As Integer, ByVal busaddr As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpibbusaddr(id1, busaddr)
    igpibbusaddr = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function igpibgett1delay(ByVal id1 As Integer, delay As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpibgett1delay(id1, delay)
    igpibgett1delay = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function igpibsett1delay(ByVal id1 As Integer, ByVal delay As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpibsett1delay(id1, delay)
    igpibsett1delay = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function igpibpulseifc(ByVal id1 As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpibpulseifc(id)
    igpibpulseifc = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function igpioctrl(ByVal id1 As Integer, ByVal request As Integer, ByVal setting As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpioctrl(id1, request, setting)
    igpioctrl = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function igpiostat(ByVal id1 As Integer, ByVal request As Integer, ByRef result As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpiostat(id1, request, result)
    igpiostat = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function igpiosetwidth(ByVal id1 As Integer, ByVal width As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpiosetwidth(id1, width)
    igpiosetwidth = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function igpiogetwidth(ByVal id1 As Integer, ByRef width As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpiogetwidth(id1, width)
    igpiogetwidth = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ilantimeout(ByVal id1 As Integer, ByVal tval As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ilantimeout(id1, tval)
    ilantimeout = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ilangettimeout(ByVal id1 As Integer, tval As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ilangettimeout(id1, tval)
    ilangettimeout = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function imap(ByVal id1 As Integer, ByVal mapSpace As Integer, ByVal pagestart As Integer, ByVal pagecnt As Integer, ByVal suggested As Long) As Long
    Dim id As Long
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Force no error to be condition
    Call vb_icauseerr(id1, 0, 0)

    ' Call the function in the SICL DLL and check for errors
    id = vb_imap(id1, mapSpace, pagestart, pagecnt, suggested)
    imap = id
    thisErrno = vb_igeterrno()
    If thisErrno <> 0 Then
        Err.Clear    ' set default values in the error object
        ' set the error string and raise the error
        tmp = vb_igeterrstr(thisErrno, myerrstr)
        Err.Description = myerrstr
        Err.Raise (thisErrno) 'Raise the error
    End If

End Function

Function imapx(ByVal id1 As Integer, ByVal mapSpace As Integer, ByVal pagestart As Integer, ByVal pagecnt As Integer) As Long
    Dim retVal As Long
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Force no error to be condition
    Call vb_icauseerr(id1, 0, 0)

    ' Call the function in the SICL DLL and check for errors
    retVal = vb_imapx(id1, mapSpace, pagestart, pagecnt)
    imapx = retVal
    thisErrno = vb_igeterrno()
    If thisErrno <> 0 Then
        Err.Clear    ' set default values in the error object
        ' set the error string and raise the error
        tmp = vb_igeterrstr(thisErrno, myerrstr)
        Err.Description = myerrstr
        Err.Raise (thisErrno) 'Raise the error
    End If

End Function

Function iunmap(ByVal id1 As Integer, ByVal addr As Long, ByVal mapSpace As Integer, ByVal pagestart As Integer, ByVal pagecnt As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iunmap(id1, addr, mapSpace, pagestart, pagecnt)
    iunmap = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function iunmapx(ByVal id1 As Integer, ByVal addr As Long, ByVal mapSpace As Integer, ByVal pagestart As Integer, ByVal pagecnt As Integer) As Integer
    Dim retVal As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    retVal = vb_iunmapx(id1, addr, mapSpace, pagestart, pagecnt)
    iunmapx = retVal
    If retVal <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function imapinfo(ByVal id1 As Integer, ByVal mapSpace As Integer, numwindows As Integer, winsize As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_imapinfo(id1, mapSpace, numwindows, winsize)
    imapinfo = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function ipokex8(ByVal siclId As Integer, ByVal mapHandle As Long, ByVal Offset As Long, ByVal value As Byte) As Integer
    Dim retVal As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    retVal = vb_ipokex8(siclId, mapHandle, Offset, value)
    ipokex8 = retVal
    If retVal <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If
End Function

Function ipeekx8(ByVal siclId As Integer, ByVal mapHandle As Long, ByVal Offset As Long, value As Byte) As Integer
    Dim retVal As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    retVal = vb_ipeekx8(siclId, mapHandle, Offset, value)
    ipeekx8 = retVal
    If retVal <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If
End Function

Function ipokex16(ByVal siclId As Integer, ByVal mapHandle As Long, ByVal Offset As Long, ByVal value As Integer) As Integer
    Dim retVal As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    retVal = vb_ipokex16(siclId, mapHandle, Offset, value)
    ipokex16 = retVal
    If retVal <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If
End Function

Function ipeekx16(ByVal siclId As Integer, ByVal mapHandle As Long, ByVal Offset As Long, value As Integer) As Integer
    Dim retVal As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    retVal = vb_ipeekx16(siclId, mapHandle, Offset, value)
    ipeekx16 = retVal
    If retVal <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If
End Function

Function ipokex32(ByVal siclId As Integer, ByVal mapHandle As Long, ByVal Offset As Long, ByVal value As Long) As Integer
    Dim retVal As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    retVal = vb_ipokex32(siclId, mapHandle, Offset, value)
    ipokex32 = retVal
    If retVal <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If
End Function

Function ipeekx32(ByVal siclId As Integer, ByVal mapHandle As Long, ByVal Offset As Long, value As Long) As Integer
    Dim retVal As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    retVal = vb_ipeekx32(siclId, mapHandle, Offset, value)
    ipeekx32 = retVal
    If retVal <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If
End Function

Function iblockmovex(ByVal siclId As Integer, ByVal srcHandle As Long, ByRef srcOffset As Variant, ByVal srcWidth As Integer, ByVal srcIncrement As Integer, ByVal destHandle As Long, ByRef destOffset As Variant, ByVal destWidth As Integer, ByVal destIncrement As Integer, ByVal count As Long, ByVal swap As Integer) As Integer
    Dim retVal As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    retVal = vb_iblockmovex(siclId, srcHandle, srcOffset, srcWidth, srcIncrement, destHandle, destOffset, destWidth, destIncrement, count, swap)
    iblockmovex = retVal
    If retVal <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If
End Function

Function ibblockcopy(ByVal id1 As Integer, ByVal src As Long, ByVal dest As Long, ByVal CNT As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ibblockcopy(id1, src, dest, CNT)
    ibblockcopy = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function iwblockcopy(ByVal id1 As Integer, ByVal src As Long, ByVal dest As Long, ByVal CNT As Long, ByVal swap As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iwblockcopy(id1, src, dest, CNT, swap)
    iwblockcopy = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ilblockcopy(ByVal id1 As Integer, ByVal src As Long, ByVal dest As Long, ByVal CNT As Long, ByVal swap As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ilblockcopy(id1, src, dest, CNT, swap)
    ilblockcopy = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ibpushfifo(ByVal id1 As Integer, ByVal src As Long, ByVal fifo As Long, ByVal CNT As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ibpushfifo(id1, src, fifo, CNT)
    ibpushfifo = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function iwpushfifo(ByVal id1 As Integer, ByVal src As Long, ByVal fifo As Long, ByVal CNT As Long, ByVal swap As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iwpushfifo(id1, src, fifo, CNT, swap)
    iwpushfifo = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ilpushfifo(ByVal id1 As Integer, ByVal src As Long, ByVal fifo As Long, ByVal CNT As Long, ByVal swap As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ilpushfifo(id1, src, fifo, CNT, swap)
    ilpushfifo = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ibpopfifo(ByVal id1 As Integer, ByVal fifo As Long, ByVal dest As Long, ByVal CNT As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ibpopfifo(id1, fifo, dest, CNT)
    ibpopfifo = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function iwpopfifo(ByVal id1 As Integer, ByVal fifo As Long, ByVal dest As Long, ByVal CNT As Long, ByVal swap As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iwpopfifo(id1, fifo, dest, CNT, swap)
    iwpopfifo = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ilpopfifo(ByVal id1 As Integer, ByVal fifo As Long, ByVal dest As Long, ByVal CNT As Long, ByVal swap As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ilpopfifo(id1, fifo, dest, CNT, swap)
    ilpopfifo = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function siclcleanup() As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb__siclcleanup()
    siclcleanup = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function setsiclyield(ByVal yield_option As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb__setsiclyield(yield_option)
    setsiclyield = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Sub ibpoke(ByVal addr As Long, ByVal value As Byte)
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Force no error to be condition
    Call vb_icauseerr(id1, 0, 0)

    ' Call the function in the SICL DLL
    Call vb_ibpoke(addr, value)

    thisErrno = vb_igeterrno()
    If thisErrno <> 0 Then
        Err.Clear    ' set default values in the error object
        ' set the error string and raise the error
        tmp = vb_igeterrstr(thisErrno, myerrstr)
        Err.Description = myerrstr
        Err.Raise (thisErrno) 'Raise the error
    End If

End Sub


Sub iwpoke(ByVal addr As Long, ByVal value As Integer)
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Force no error to be condition
    Call vb_icauseerr(id1, 0, 0)

    ' Call the function in the SICL DLL
    Call vb_iwpoke(addr, value)

    thisErrno = vb_igeterrno()
    If thisErrno <> 0 Then
        Err.Clear    ' set default values in the error object
        ' set the error string and raise the error
        tmp = vb_igeterrstr(thisErrno, myerrstr)
        Err.Description = myerrstr
        Err.Raise (thisErrno) 'Raise the error
    End If

End Sub

Sub ilpoke(ByVal addr As Long, ByVal value As Long)
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Force no error to be condition
    Call vb_icauseerr(id1, 0, 0)

    ' Call the function in the SICL DLL
    Call vb_ilpoke(addr, value)

    thisErrno = vb_igeterrno()
    If thisErrno <> 0 Then
        Err.Clear    ' set default values in the error object
        ' set the error string and raise the error
        tmp = vb_igeterrstr(thisErrno, myerrstr)
        Err.Description = myerrstr
        Err.Raise (thisErrno) 'Raise the error
    End If

End Sub

Function ibpeek(ByVal addr As Long) As Byte
    Dim id As Byte

    ' Call the function in the SICL DLL and check for errors
    id = vb_ibpeek(addr)
    ibpeek = id

End Function

Function iwpeek(ByVal addr As Long) As Integer
    Dim id As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iwpeek(addr)
    iwpeek = id

End Function

Function ilpeek(ByVal addr As Long) As Long
    Dim id As Long

    ' Call the function in the SICL DLL and check for errors
    id = vb_ilpeek(addr)
    ilpeek = id

End Function


' This function truncates a string so that all characters
' following a carriage return or linefeed character are
' removed.  The truncated string is then returned.
Function strip_crlf(szString As String) As String
   Dim crlfpos As Integer

   crlfpos = InStr(szString, Chr$(13))
   If crlfpos Then
     szString = Left(szString, crlfpos - 1)
   End If
   crlfpos = InStr(szString, Chr$(10))
   If crlfpos Then
     szString = Left(szString, crlfpos - 1)
   End If

   strip_crlf = szString
End Function

Function icmd(ByVal id1 As Integer, ByVal cmd As Long, ByVal datalen As Integer, ByVal datawidth As Integer, ByRef pdata As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_icmd(id1, cmd, datalen, datawidth, pdata)
    icmd = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Public Function isprintf(ByRef s As String, ByVal fmt As String, ParamArray vararg() As Variant) As Integer
    Dim retVal As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer
    If (UBound(vararg) <> -1) Then
       ReDim va(LBound(vararg) To UBound(vararg)) As Variant
    Else
       ReDim va(0 To 0) As Variant
       va(0) = 0
    End If
    
    ' Force no error to be condition
    Call vb_icauseerr(0, 0, 0)
    
    For i = LBound(vararg) To UBound(vararg)
       va(i) = vararg(i)
    Next i

    ' Call the function in the SICL DLL and check for errors
    retVal = vb_iprint(1, 0, s, ByVal fmt, va)
    
    isprintf = retVal

    thisErrno = vb_igeterrno()
    If thisErrno <> 0 Then
        Err.Clear    ' set default values in the error object
        ' set the error string and raise the error
        tmp = vb_igeterrstr(thisErrno, myerrstr)
        Err.Description = myerrstr
        Err.Raise (thisErrno) 'Raise the error
    End If

End Function

Public Function iprintf(ByRef id As Integer, ByVal fmt As String, ParamArray vararg() As Variant) As Integer
    Dim retVal As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer
    If (UBound(vararg) <> -1) Then
       ReDim va(LBound(vararg) To UBound(vararg)) As Variant
    Else
       ReDim va(0 To 0) As Variant
       va(0) = 0
    End If
    
    ' Force no error to be condition
    Call vb_icauseerr(0, 0, 0)
    
    For i = LBound(vararg) To UBound(vararg)
       va(i) = vararg(i)
    Next i

    ' Call the function in the SICL DLL and check for errors
    retVal = vb_iprint(2, id, 0, ByVal fmt, va)
    
    iprintf = retVal

    thisErrno = vb_igeterrno()
    If thisErrno <> 0 Then
        Err.Clear    ' set default values in the error object
        ' set the error string and raise the error
        tmp = vb_igeterrstr(thisErrno, myerrstr)
        Err.Description = myerrstr
        Err.Raise (thisErrno) 'Raise the error
    End If

End Function

Public Function isscanf(ByVal s As String, ByVal fmt As String, Optional va1 As Variant, Optional va2 As Variant, Optional va3 As Variant, Optional va4 As Variant, Optional va5 As Variant, Optional va6 As Variant, Optional va7 As Variant, Optional va8 As Variant, Optional va9 As Variant, Optional va10 As Variant) As Integer
    Dim retVal As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer
    
    ' Force no error to be condition
    Call vb_icauseerr(0, 0, 0)
    
    ' Call the function in the SICL DLL and check for errors
    retVal = vb_iscan(3, 0, ByVal s, ByVal fmt, va1, va2, va3, va4, va5, va6, va7, va8, va9, va10)
    
    isscanf = retVal

    thisErrno = vb_igeterrno()
    If thisErrno <> 0 Then
        Err.Clear    ' set default values in the error object
        ' set the error string and raise the error
        tmp = vb_igeterrstr(thisErrno, myerrstr)
        Err.Description = myerrstr
        Err.Raise (thisErrno) 'Raise the error
    End If

End Function

Public Function iscanf(ByRef id As Integer, ByVal fmt As String, Optional va1 As Variant, Optional va2 As Variant, Optional va3 As Variant, Optional va4 As Variant, Optional va5 As Variant, Optional va6 As Variant, Optional va7 As Variant, Optional va8 As Variant, Optional va9 As Variant, Optional va10 As Variant) As Integer
    Dim retVal As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer
    
    ' Force no error to be condition
    Call vb_icauseerr(0, 0, 0)
   
    ' Call the function in the SICL DLL and check for errors
    retVal = vb_iscan(4, id, 0, ByVal fmt, va1, va2, va3, va4, va5, va6, va7, va8, va9, va10)
    
    iscanf = retVal

    thisErrno = vb_igeterrno()
    If thisErrno <> 0 Then
        Err.Clear    ' set default values in the error object
        ' set the error string and raise the error
        tmp = vb_igeterrstr(thisErrno, myerrstr)
        Err.Description = myerrstr
        Err.Raise (thisErrno) 'Raise the error
    End If

End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   DLL entry function declarations needed for GPIBnotify OLE control
'This function converts a hexadecimal string to an integer decimal
Function HEX2DEC(s As String) As Long 'byte
Dim i, bin_string_index, hex_char_index As Integer
Dim bin_char, hex_char As String
Const bin_string = "0000000100100011010001010110011110001001101010111100110111101111"
Const hex_string = "0123456789ABCDEF"

bin_char = ""
For i = 1 To Len(s)
    hex_char = Mid(s, i, 1)
    hex_char_index = InStr(hex_string, hex_char)
    bin_string_index = hex_char_index * 4 - 3
    bin_char = bin_char & Mid(bin_string, bin_string_index, 4)
    'Debug.Print bin_char
Next i

For i = 1 To Len(bin_char)
    If Mid(bin_char, i, 1) = "1" Then HEX2DEC = HEX2DEC + 2 ^ (Len(bin_char) - i)
Next i
HEX2DEC = CInt(HEX2DEC)

End Function

'This function converts a decimal to an binary string
Function DEC2BIN(dec As Long, Optional size As Long) As String
Dim i As Long
Dim HexChar As String
Dim Dec2Hex As String
Dim BinaryStringIndex, HexStringIndex As Long
Const BinaryString = "0000000100100011010001010110011110001001101010111100110111101111"
Const HexString = "0123456789ABCDEF"

Dec2Hex = Hex(dec)
For i = 1 To Len(Dec2Hex)
    HexChar = Mid(Dec2Hex, i, 1)
    HexStringIndex = InStr(HexString, HexChar)
    BinaryStringIndex = HexStringIndex * 4 - 3
    DEC2BIN = DEC2BIN & Mid(BinaryString, BinaryStringIndex, 4)
Next i
If size = 0 Then Exit Function
If Len(DEC2BIN) <> size Then
    For i = 1 To size - Len(DEC2BIN)
        DEC2BIN = "0" & DEC2BIN
    Next i
End If

End Function

'This function converts a decimal string to an hexadecimal string
Function Dec2Hex(dec As Long, size As String) As String
Dim temp As String
    
temp = Hex(dec)
If size < Len(temp) Then GoTo dec2hex_skip
Do While Len(temp) <> size
    temp = "0" & temp
Loop

dec2hex_skip:
Dec2Hex = CStr(temp)

End Function

'This function converts a binary string to an integer decimal
Function BIN2DEC(BinaryInput As Long) As Long
Dim LenBinaryInput As Long
Dim i As Long

LenBinaryInput = Len(CStr(BinaryInput))
For i = 1 To LenBinaryInput
    If Mid(BinaryInput, i, 1) = "1" Then BIN2DEC = BIN2DEC + 2 ^ (LenBinaryInput - i)
Next i
BIN2DEC = CLng(BIN2DEC)

End Function




'////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////
'This function is for delay
Sub wait(msec As Long)
Dim t As Single

If msec = 0 Then DoEvents: Exit Sub
t = Timer
Do
    DoEvents
Loop Until Abs((Timer - t) * 1000) >= msec

End Sub


Function CheckFile(filename As String) As Boolean
Dim ChkName, Index As String, x As Integer, index1 As String

CheckFile = False
For x = 0 To 1
If x = 0 Then Index = "\*.bmp": index1 = ".BMP" Else Index = "\*.tif": index1 = ".TIFF"
ChkName = Dir(ThisWorkbook.Path & Index)
Do While ChkName <> ""
    If ChkName = filename & index1 Then CheckFile = True: Exit Do
    ChkName = Dir
Loop
Next x
'ChkName = Dir(ThisWorkbook.Path & "\*.tif")
'Do While ChkName <> ""
'    If ChkName = filename & ".TIFF" Then CheckFile = True: Exit Do
'    ChkName = Dir
'Loop
End Function

Sub Bprint(id As Long, Command As Variant, Optional delayS As Long, Optional ProgMode As Boolean)
Dim Istart As Integer, Ilength As Long, data As String
Dim i As Long, CmdLen As Long
Dim segment As Long
Dim x As Long

On Error GoTo CmdERROR
CmdLen = Len(Command)
For i = 1 To CmdLen
If i = 1 Then Istart = 1
Ilength = Ilength + 1
data = data + Mid(Command, i, 1)
If i = CmdLen Then ivprintf id, data & Chr(10): GoTo ESKAPE
'If Mid(Command, i, 1) = "," Then
'data = Mid(Command, Istart, Ilength - 1): ivprintf id, data & Chr(10)
'wait delayS: Istart = i + 1: Ilength = 0: data = ""
'End If
Next i
CmdERROR:
MsgBox "You Have Encountered ERROR!!!", vbCritical, "Error Code": End
ESKAPE:
wait delayS
'If ProgMode = True Then
'For x = 20 To delayS Step 20
'If x = 0 Then PROGRESS1.ProgressBar1.value = Myprog + 20
'PROGRESS1.ProgressBar1.value = PROGRESS1.ProgressBar1.value + 20
'wait 20
'Next x
'Myprog = PROGRESS1.ProgressBar1.value
'End If
End Sub

Function Bread(id As Long, prow As Long, pcol As Long, delay As Long) As Single
Dim y As String * 30, z As Long, length As Long
Dim i As Long, Istart As Integer, count As Long, data As String
On Error GoTo CmdERROR
'y = "1.2345E-9"
wait delay
iread id, y, 200, &O0, z
length = Len(y)
For i = 1 To length
If Mid(y, i, 1) = "E" Then count = 1
data = data + Mid(y, i, 1)
If count > 0 Then count = count + 1
If count > 4 Then GoTo bend
Next i
CmdERROR:
MsgBox "You Have Encountered ERROR!!!", vbCritical, "Error Code": End
bend:
'MsgBox "READ COMPLETE"
ActiveSheet.Select
ActiveSheet.Cells(prow, pcol).value = CSng(data) * 1000000

End Function

Sub readtest()
BOpen YScope, 1
'Dim YScope As Integer
'YScope = iopen("gpib0,1")
Bprint YScope, "START"
Bprint YScope, "MEASURE:CHANNEL1:MIN:VALUE?"
Bread YScope, 1, 6, V
MsgBox "READ COMPLETE"
End Sub

Sub BOpen(id As Long, addrs As Long)
id = iopen("gpib0," & addrs)
'MsgBox "CHECK"
End Sub

Function OUTA(datain As Byte) As Boolean
    Dim result As Long
    Dim data(5) As Byte
    
    data(0) = &HA  'Command
    data(1) = datain
    
    result = USB_Send(handle, data(0), 2)
    
    If result = 0 Then
        OUTA = False
    Else
        OUTA = True
    End If
End Function

Function OUTB(datain As Byte) As Boolean
    Dim result As Long
    Dim data(5) As Byte
    
    data(0) = &HB  'Command
    data(1) = datain
    
    result = USB_Send(handle, data(0), 2)
    
    If result = 0 Then
        OUTB = False
    Else
        OUTB = True
    End If
End Function

Function OUTC(datain As Byte) As Boolean
    Dim result As Long
    Dim data(5) As Byte
    
    data(0) = &HC  'Command
    data(1) = datain
    
    result = USB_Send(handle, data(0), 2)
    
    If result = 0 Then
        OUTC = False
    Else
        OUTC = True
    End If
End Function

Function trimmer(q As String) As String
Dim counter As Integer
Dim fdata As String
Dim darray As Variant
Dim i As Long, lp As Integer, x As Integer
darray = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "+", "-", ".", "E")
'fdata = ""
'counter = 0
'y = data
lp = Len(y)

For x = 1 To lp
    For i = 0 To 13 Step 1
    If Mid(y, x, 1) = darray(i) Then fdata = fdata + Mid(y, x, 1)
    'If Mid(y, i, 1) <> darray(i - 1) Then GoTo dito
    Next i
If Mid(y, x, 1) = "E" Then counter = 1
If counter >= 1 Then counter = counter + 1
If counter > 3 Then GoTo dito
Next x
dito:
trimmer = fdata
End Function
Function trimmer8902(id As String) As String
Dim x As Integer
Dim count As Integer
count = 0
For x = 1 To Len(id)
If Mid(id, x, 1) = "E" Then trimmer8902 = Mid(id, 1, (x + 3)): Exit Function
Next x
End Function
Function TrimMe(MyString As String, MyChar As String) As String
Dim x As Integer
Dim CNT As Integer
Dim posi1 As Integer
Dim posi2 As Integer
For x = 1 To Len(MyString)
If Mid(MyString, x, 1) = MyChar Then
CNT = CNT + 1
    If CNT = 1 Then posi1 = x
    If CNT = 2 Then posi2 = x: GoTo baba
End If
Next x
baba:
TrimMe = Mid(MyString, posi1 + 1, posi2 - posi1 - 1)
End Function
Function trimmer5790(id As String) As Single
Dim x As Integer
Dim count As Integer
count = 0
For x = 1 To Len(id)
If Mid(id, x, 1) = "E" Then trimmer5790 = Mid(id, 1, (x + 3)): Exit Function
Next x
End Function
Function trimmer4284(id As String) As Single
Dim x As Integer
Dim count As Integer
count = 0
For x = 1 To Len(id)
If Mid(id, x, 1) = "E" Then trimmer4284 = Mid(id, 1, (x + 3)): Exit Function
Next x
End Function

Sub FindInstruments()

Dim defrm As Long, ViInLong&, VIRetCountLong&, InstAddr$, InstSession&, NewInstAddr$
Dim z As Long, figure As String * 500, IDNName() As String, i%

Call viOpenDefaultRM(defrm)
InstAddr = Space(50)
For i = 1 To 32
   If i = 1 Then Call viFindRsrc(defrm, "GPIB[0-9]*::?*INSTR", ViInLong, VIRetCountLong, InstAddr)
   If i <> 1 Then Call viFindNext(ViInLong, InstAddr)
     
   If VIRetCountLong = 0 Then
        MsgBox "No instruments were Found", vbOKOnly, "WARNING"
         GoTo JumpNext
   End If
     
   ReDim Preserve IDNName(VIRetCountLong) As String
   NewInstAddr = Mid(InstAddr, 1, InStr(1, InstAddr, Chr(&H0), vbBinaryCompare) - 1)
     
    Call viOpen(defrm, NewInstAddr, 0, 0, InstSession)
    Call viVPrintf(InstSession, "*IDN?" & Chr(10), 0)
    Call viRead(InstSession, figure, 50, z)
    IDNName(i) = figure
     'IDNName(i) = Mid(figure, 1, InStr(1, figure, Chr(10), vbBinaryCompare) - 1)
     Sheets(Sheet1).Cells(11 + i, 2).value = IDNName(i)
 '''    Debug.Print IDNName(i)
     Call viClose(InstSession)
     If i = VIRetCountLong Then Exit For
 Next i
 
JumpNext:
     Call viClose(ViInLong)
    Call viClose(defrm)
 If VIRetCountLong = 0 Then
    wait 100
    'Call GetDivision
    'Call GetJONumber
    'Call GetEngineerName
    'Call GetParameterID
    'Call GetDate
    End
 End If

End Sub
Function getdata(ByVal inst_id As Integer, Optional numonly As Boolean = True, _
Optional LocalMode As Boolean = False, Optional EndChar As String) As String
Dim Iscan As Variant
Iscan = Array("+", "-", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", _
            "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")

Dim z As Long, figure As String * 50
Dim x As Integer, lp As Integer
Dim IDENTITY As Boolean
If EndChar <> "" Then Call Termchar(inst_id, EndChar)
IDENTITY = False
getdata = vbNullString
'Select Case UCase(MySession)
 '   Case "AGILENT"
        Call vb_iread(1, inst_id, figure, 50, &O0, z)
        'MsgBox "monitor1" & figure
 '   Case "NI"
      '   Case "VISA"
 '       Call viRead(inst_id, figure, 50, Z)
'End Select
'If numonly = True Then getdata = read(figure) Else getdata = figure
'If LocalMode = True Then Call Instlocal(inst_id)
'getdata = Mid(figure, 1, 15)
'getdata = figure
For i = 1 To Len(figure)
IDENTITY = False
    For lp = 0 To 38
        If Mid(figure, i, 1) = Iscan(lp) Then IDENTITY = True: GoTo Skip
    Next lp
    If IDENTITY = False Then x = i: GoTo SKIP2
Skip:
    
Next i
SKIP2:

'MsgBox figure & "ABC"
getdata = Mid(figure, 1, i - 1)
'MsgBox "monitor2" & getdata
IDENTITY = False
End Function
Function getdata5790(ByVal inst_id As Integer, Optional numonly As Boolean = True, _
Optional LocalMode As Boolean = False, Optional EndChar As String) As String
Dim Iscan As Variant
Iscan = Array("+", "-", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", _
            "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")

Dim z As Long, figure As String * 50
Dim x As Integer, lp As Integer
Dim IDENTITY As Boolean
If EndChar <> "" Then Call Termchar(inst_id, EndChar)
IDENTITY = False
getdata5790 = vbNullString
'Select Case UCase(MySession)
 '   Case "AGILENT"
        Call vb_iread(1, inst_id, figure, 50, &O0, z)
        'MsgBox "monitor1" & figure
 '   Case "NI"
      '   Case "VISA"
 '       Call viRead(inst_id, figure, 50, Z)
'End Select
'If numonly = True Then getdata = read(figure) Else getdata = figure
'If LocalMode = True Then Call Instlocal(inst_id)
'getdata = Mid(figure, 1, 15)
'getdata = figure
'For i = 1 To Len(figure)
'IDENTITY = False
'    For lp = 0 To 38
'        If Mid(figure, i, 1) = Iscan(lp) Then IDENTITY = True: GoTo Skip
'    Next lp
'    If IDENTITY = False Then x = i: GoTo SKIP2
'Skip:
'
'Next i
'SKIP2:

'MsgBox figure & "ABC"
getdata5790 = figure
'MsgBox "monitor2" & getdata
IDENTITY = False
End Function

Function getdata4284(ByVal inst_id As Integer, Optional numonly As Boolean = True, _
Optional LocalMode As Boolean = False, Optional EndChar As String) As String
Dim Iscan As Variant
Iscan = Array("+", "-", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", _
            "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")

Dim z As Long, figure As String * 50
Dim x As Integer, lp As Integer
Dim IDENTITY As Boolean
If EndChar <> "" Then Call Termchar(inst_id, EndChar)
IDENTITY = False
getdata4284 = vbNullString
'Select Case UCase(MySession)
 '   Case "AGILENT"
        Call vb_iread(1, inst_id, figure, 50, &O0, z)
        'MsgBox "monitor1" & figure
 '   Case "NI"
      '   Case "VISA"
 '       Call viRead(inst_id, figure, 50, Z)
'End Select
'If numonly = True Then getdata = read(figure) Else getdata = figure
'If LocalMode = True Then Call Instlocal(inst_id)
'getdata = Mid(figure, 1, 15)
'getdata = figure
'For i = 1 To Len(figure)
'IDENTITY = False
'    For lp = 0 To 38
'        If Mid(figure, i, 1) = Iscan(lp) Then IDENTITY = True: GoTo Skip
'    Next lp
'    If IDENTITY = False Then x = i: GoTo SKIP2
'Skip:
'
'Next i
'SKIP2:

'MsgBox figure & "ABC"
getdata4284 = figure
'MsgBox "monitor2" & getdata
IDENTITY = False
End Function




Function getdata8902(ByVal inst_id As Integer, Optional numonly As Boolean = True, _
Optional LocalMode As Boolean = False, Optional EndChar As String) As String
Dim Iscan As Variant
Iscan = Array("+", "-", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", _
            "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")

Dim z As Long, figure As String * 50
Dim x As Integer, lp As Integer
Dim IDENTITY As Boolean
If EndChar <> "" Then Call Termchar(inst_id, EndChar)
IDENTITY = False
getdata8902 = vbNullString
'Select Case UCase(MySession)
 '   Case "AGILENT"
        Call vb_iread(1, inst_id, figure, 50, &O0, z)
        'MsgBox "monitor1" & figure
 '   Case "NI"
      '   Case "VISA"
 '       Call viRead(inst_id, figure, 50, Z)
'End Select
'If numonly = True Then getdata = read(figure) Else getdata = figure
'If LocalMode = True Then Call Instlocal(inst_id)
'getdata = Mid(figure, 1, 15)
'getdata = figure
For i = 1 To Len(figure)
IDENTITY = False
    For lp = 0 To 38
        If Mid(figure, i, 1) = Iscan(lp) Then IDENTITY = True: GoTo Skip
    Next lp
    If IDENTITY = False And i = 1 Then GoTo Skip
    If IDENTITY = False Then x = i: GoTo SKIP2
Skip:
    
Next i
SKIP2:

'MsgBox figure & "ABC"
getdata8902 = Mid(figure, 1, i - 1)
'MsgBox "monitor2" & getdata
IDENTITY = False
End Function

Sub Termchar(ByVal inst_id As Integer, Optional EndChar As String)

Select Case UCase(MySession)
    Case "AGILENT"
        Call vb_itermchr(inst_id, Asc(EndChar))
    Case "NI"

    Case "VISA"
        'Call viSetAttribute(inst_id, VI_ATTR_TERMCHAR, Asc(EndChar))
End Select
End Sub


Sub MyPBAR(pmin As Long, pmax As Long, myWait As Long)
Dim x As Long

'PROGRESS2.Progressbar1.Max = (pmax * myWait * 14) + 1
PROGRESS2.ProgressBar2.Max = (pmax * myWait * 14) + 1
'For x = (pmin * 1000) To (pmax * 1000) Step 1
For x = 1 To (myWait * 14) Step 1
PROGRESS2.ProgressBar2.value = PROGRESS2.ProgressBar2.value + 1
wait 0
Next x
'Sheet1.Cells(1, 5).value = CStr(Time)
PbarStart = PbarStart + 1: Exit Sub
End Sub

Sub mybartime()
Dim startTym As String
Dim stopTym As String
Dim x As Long
Dim y As Long
PROGRESS2.Show 0
wait 1000
PROGRESS2.ProgressBar1.Min = 0: wait 1000: PROGRESS2.ProgressBar1.Max = 675016
Sheet1.Cells(1, 4).value = CStr(Time)
    For lp = 1 To 15
        For x = 1 To 45000 Step 1
        PROGRESS2.ProgressBar1.value = PROGRESS2.ProgressBar1.value + 1
        wait 0
        'If PROGRESS2.Progressbar1.value = 675000 Then MsgBox y
        Next x
    Next lp
Sheet1.Cells(1, 5).value = CStr(Time)
End Sub

Function testcheck() As Boolean
If Form1.Test1.value = False And Form1.Test2.value = False And Form1.Test2DCV.value = False And _
   Form1.Test2OHM.value = False And _
   Form1.Test4.value = False And Form1.Test5.value = False And Form1.Test5LIN.value = False And _
   Form1.Test6.value = False And Form1.Test7.value = False And Form1.Test8.value = False And _
   Form1.Test9(1).value = False And Form1.Test10(1).value = False And Form1.Test11(1).value = False And _
   Form1.Test12.value = False And Form1.TestOpen.value = False Then
   '''''remove the form1.test3 and form1.test3DCV
If MsgBox("You have not chosen any test item in the list." & Chr(10) & _
       "Click ""OK"" to select test item/s." & Chr(10) & _
       "Click ""CANCEL"" to exit this program.", vbOKCancel + vbExclamation, "MP AUTOCAL SOFWARE") = vbOK Then testcheck = True Else End
End If
End Function

Sub MEASURE_BA()
CalErrLog = False
'On Error GoTo skip
Dim Mainlp As Integer
Dim PassFail As Integer
Dim DevFreq As Variant  'Test 2
Dim Ffreq As Variant
Dim stopBit As Integer
Dim CFreq500 As Variant 'Test13,Test 16
Dim CFreq300k As Variant
Dim DA As Variant
Dim Flev As Variant
Dim lp2 As Integer
Dim lp As Integer
Dim CNT As Integer
Dim NewFreq
Dim NewFreqDB
Dim BandRef

If testcheck = True Then Exit Sub
For Mainlp = 1 To 17
If Mainlp = 1 And Form1.Test1 = 1 Then GoTo Test1           '''OPEN
If Mainlp = 2 And Form1.Test2 = 1 Then GoTo Test2           '''ACV / DCV/ OHM ZERO
If Mainlp = 3 And Form1.Test2DCV = 1 Then GoTo Test2DCV     '''ACV / DCV/ OHM ZERO
If Mainlp = 4 And Form1.Test2OHM = 1 Then GoTo Test2OHM     '''ACV / DCV/ OHM ZERO
If Mainlp = 5 And Form1.Test3 = 1 Then GoTo Test3           '''Rear OHM ZERO
If Mainlp = 6 And Form1.Test3DCV = 1 Then GoTo Test3DCV     ''REAR DCV ZERO
If Mainlp = 7 And Form1.Test4 = 1 Then GoTo Test4           ''LOW I ZERO
If Mainlp = 8 And Form1.Test5 = 1 Then GoTo Test5           ''HI I ZERO
If Mainlp = 9 And Form1.Test5LIN = 1 Then GoTo Test5LIN     ''LINEARITY
If Mainlp = 10 And Form1.Test6 = 1 Then GoTo Test6          ''ACV GAIN
If Mainlp = 11 And Form1.Test7 = 1 Then GoTo Test7          ''VDC GAIN
If Mainlp = 12 And Form1.Test8 = 1 Then GoTo Test8          ''HI IDC GAIN
If Mainlp = 13 And Form1.Test9(1) = 1 Then GoTo Test9          ''HI IAC GAIN
If Mainlp = 14 And Form1.Test10(1) = 1 Then GoTo Test10        ''LOW IAC GAIN
If Mainlp = 15 And Form1.Test11(1) = 1 Then GoTo Test11        ''LOW IDC GAIN
If Mainlp = 16 And Form1.Test12 = 1 Then GoTo Test12        ''OHM GAIN
If Mainlp = 17 And Form1.TestOpen = 1 Then GoTo TestOpen
'If Mainlp = 9 And Form1.Test9 = 1 Then GoTo Test9
'If Mainlp = 10 And Form1.Test10 = 1 Then GoTo Test10
'If Mainlp = 11 And Form1.Test11 = 1 Then GoTo Test11

GoTo Skip:
'**********************************************************************************************
'**********************************************************************************************
TestOpen:   '''UNLOCK/LOCK
MyTestOpen
If MainSkipFlag = True Then MainSkipFlag = False: Form1.Show: GoTo SkipGlobal2
GoTo Skip:
'**********************************************************************************************
'**********************************************************************************************
Test1:   'ÓPEN
MyTest1
If MainSkipFlag = True Then MainSkipFlag = False: Form1.Show: GoTo SkipGlobal2
GoTo Skip:
'*****************************************************************************************************
'*****************************************************************************************************
Test2:   ' 'ACV / DCV/ OHM ZERO
MyTest2
If MainSkipFlag = True Then MainSkipFlag = False: Form1.Show: GoTo SkipGlobal2
GoTo Skip:
'*****************************************************************************************************
'*****************************************************************************************************
Test2DCV:    'ACV / DCV/ OHM ZERO
MyTest2DCV
If MainSkipFlag = True Then MainSkipFlag = False: Form1.Show: GoTo SkipGlobal2
GoTo Skip:
'*****************************************************************************************************
'*****************************************************************************************************
Test2OHM:    'ACV / DCV/ OHM ZERO
MyTest2OHM
If MainSkipFlag = True Then MainSkipFlag = False: Form1.Show: GoTo SkipGlobal2
GoTo Skip:
'*****************************************************************************************************
'*****************************************************************************************************
Test3:   '''Rear OHM ZERO
MyTest3
If MainSkipFlag = True Then MainSkipFlag = False: Form1.Show: GoTo SkipGlobal2
GoTo Skip:

'*****************************************************************************************************
Test3DCV:   ''''Rear DCV ZERO
MyTest3DCV
If MainSkipFlag = True Then MainSkipFlag = False: Form1.Show: GoTo SkipGlobal2
GoTo Skip:
'*****************************************************************************************************
Test4:   ''LOW I ZERO
MyTest4
If MainSkipFlag = True Then MainSkipFlag = False: Form1.Show: GoTo SkipGlobal2
GoTo Skip:
'*****************************************************************************************************
Test5:    ''HI I ZERO
MyTest5
If MainSkipFlag = True Then MainSkipFlag = False: Form1.Show: GoTo SkipGlobal2
GoTo Skip:
'*****************************************************************************************************
Test5LIN:  'LINEARITY
MyTest5LIN
If MainSkipFlag = True Then MainSkipFlag = False: Form1.Show: GoTo SkipGlobal2
GoTo Skip:
'*****************************************************************************************************

Test6:   ''ACV GAIN
MyTest6
If MainSkipFlag = True Then MainSkipFlag = False: Form1.Show: GoTo SkipGlobal2
GoTo Skip:
'*****************************************************************************************************

Test7:   ''VDC GAIN
MyTest7
If MainSkipFlag = True Then MainSkipFlag = False: Form1.Show: GoTo SkipGlobal2
GoTo Skip:
'*****************************************************************************************************

Test8:     ''HI IDC GAIN
MyTest8
If MainSkipFlag = True Then MainSkipFlag = False: Form1.Show: GoTo SkipGlobal2
GoTo Skip:
'*****************************************************************************************************
''*****************************************************************************************************
Test9:  ''HI IAC GAIN
MyTest9
If MainSkipFlag = True Then MainSkipFlag = False: Form1.Show: GoTo SkipGlobal2
GoTo Skip:
'*****************************************************************************************************
Test10:    ''LOW IAC GAIN
MyTest10
If MainSkipFlag = True Then MainSkipFlag = False: Form1.Show: GoTo SkipGlobal2
GoTo Skip:
'*****************************************************************************************************
Test11:     ''LOW IDC GAIN
MyTest11
If MainSkipFlag = True Then MainSkipFlag = False: Form1.Show: GoTo SkipGlobal2
GoTo Skip:
'*****************************************************************************************************
Test12:   ''OHM GAIN
MyTest12
If MainSkipFlag = True Then MainSkipFlag = False: Form1.Show: GoTo SkipGlobal2
GoTo Skip:
'*****************************************************************************************************

Skip:
Next Mainlp
SkipGlobal:
If MsgBox("Performance Test (as found) finished" & Chr(10) & _
       "Would you like to perform another test?", vbYesNo + vbInformation, "MPC AUTOCAL SOFTWARE") = vbNo Then GoTo FINISHED Else Form1.Show: Exit Sub
FINISHED:
If MsgBox("Do you like to exit this program?", vbYesNo + vbInformation, "MPC AUTOCAL SOFTWARE") = vbNo Then Form1.Show: Exit Sub Else PROGRESS1.Show
SkipGlobal2:
End Sub

'Function TRIMKO(id As String) As String
'Dim x As Integer
'MsgBox id
'MsgBox Len(id)
'For x = 1 To Len(id)
'If Mid(id, x, 1) = "," Then MsgBox "x=" & x: TRIMKO = cstrMid(id, 1, (x - 1)): GoTo skip
'Next x
'skip:
'End Function
Function getdata3458(ByVal inst_id As Integer, Optional numonly As Boolean = True, _
Optional LocalMode As Boolean = False, Optional EndChar As String) As String
Dim Iscan As Variant
Iscan = Array("+", "-", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", _
              "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", " ")

Dim z As Long, figure As String * 50
Dim x As Integer, lp As Integer
Dim IDENTITY As Boolean
If EndChar <> "" Then Call Termchar(inst_id, EndChar)
IDENTITY = False
getdata3458 = vbNullString
'Select Case UCase(MySession)
 '   Case "AGILENT"
        Call vb_iread(1, inst_id, figure, 50, &O0, z)
 '   Case "NI"
        
 '   Case "VISA"
 '       Call viRead(inst_id, figure, 50, Z)
'End Select
'If numonly = True Then getdata = read(figure) Else getdata = figure
'If LocalMode = True Then Call Instlocal(inst_id)
'getdata = Mid(figure, 1, 15)
figure = Mid(figure, 1, (Len(figure) - 1))
For i = 1 To Len(figure)
IDENTITY = False
    For lp = 0 To 39
        If Mid(figure, i, 1) = Iscan(lp) Then IDENTITY = True: GoTo Skip
    Next lp
    If IDENTITY = False Then x = i: GoTo SKIP2
Skip:
    
Next i
SKIP2:
getdata3458 = Mid(figure, 1, i - 1)
IDENTITY = False
End Function

Private Function Log10(ByVal Number As Single) As Double
       Log10 = Log(Number) / Log(10)
End Function


Sub CHECK_DMM(code3458 As String, code34401 As String)
DMMflag = False
If MsgBox("Did you select " & Form1.CalName_3458 & " in the main menu?" & Chr(10) & _
          "Press ""Yes"" if you did." & Chr(10) & _
          "Press ""No"" to go back.", vbYesNo + vbExclamation) = vbNo Then DMMflag = True: GoTo SkipGlobal
If Mid(Form1.CalName_3458, 1, 5) = "3458A" Then
blikterm2:
            Bprint CalInst2, "PRESET", 1000
            Bprint CalInst2, "TERM?", 1000
            If CalErrLog = True Then CalErrLog = False: GoTo SkipGlobal
            y = getdata(CInt(CalInst2))
            
            If CInt(y) = 2 Or CInt(y) = 0 Then MsgBox "Please set the measuring terminal to FRONT (released)": GoTo blikterm2
            'Bprint CalInst2, "RESET", 1000
            'Bprint CalInst2, "FUNC DCV", 100
            Bprint CalInst2, "FUNC " & code3458, 100
            Bprint CalInst2, "ARANGE", 100: ilocal CalInst2
End If
If Mid(Form1.CalName_3458, 1, 5) = "34410" Then
    Bprint CalInst2, "*RST", 1000 ': ilocal CalInst2
    'Bprint CalInst2, "SENSE:FUNCTION ""VOLT:DC""": wait 200: ilocal CalInst2
    Bprint CalInst2, "CONF:VOLT:" & code34401: wait 200: ilocal CalInst2
End If
SkipGlobal:
End Sub
'ALWAYS ON TOP (1)
      Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) _
         As Long

         If Topmost = True Then 'Make the window topmost
            SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, _
               0, FLAGS)
         Else
            SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, _
               0, 0, FLAGS)
            SetTopMostWindow = False
         End If
         
      End Function


Function TEST_SETUP(SetupNum As String, MyCaption As String) As Boolean
TEST_SETUP = False
Form1.Hide: MyPIC.Image1.Picture = LoadPicture("C:\vee_user\fluke\8845a adj\8845A_Adjust_" & SetupNum & ".jpg"): MyPIC.Caption = "MP AUTOCAL SOFTWARE " & MyCaption
''Form1.Hide: MyPIC.Image1.Picture = LoadPicture("C:\temp\hp\34970A_vb\34970A_vb_SETUP1_3.jpg"): MyPIC.Caption = "MP AUTOCAL SOFTWARE " & MyCaption

MyPIC.Label1.Caption = "Connect equipments same as above." & Chr(10) & _
                         "When connection is ""OK"" click to start measurement."
MyPIC.Show 1: If TestExit = True Then TestExit = False: TEST_SETUP = True: Exit Function
If TestExit = True Then TestExit = False: TEST_SETUP = True: Exit Function
End Function

Sub Bopen_All(Dev34401 As Long, Optional CInst_5700 As Long, Optional Cinst2 As Long, Optional Cinst3 As Long, _
 Optional Cinst4 As Long, Optional Cinst7 As Long, Optional Cinst8 As Long)
    If Dev34401 = 1 Then BOpen DevInst, CLng(Form1.InstAdd.Text): If CalErrLog = True Then Exit Sub
    If CInst_5700 = 1 Then BOpen CalInst, CLng(Form1.CalAdd_5700.Text): If CalErrLog = True Then Exit Sub
End Sub

