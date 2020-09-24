Attribute VB_Name = "Module1"
Option Explicit

'I'm a Chinese undergraduate student
'excuse my poor English ~_~!
'Code By TZWSOHO

Private Const WS_CHILD = &H40000000
Private Const WS_VISIBLE = &H10000000

Private Const SWP_NOMOVE = &H2&
Private Const SWP_NOZORDER = &H4&
Private Const SWP_NOSENDCHANGING = &H400&

Private Const WM_USER = &H400
Private Const WM_CAP_START = WM_USER
Private Const WM_CAP_DLG_VIDEOFORMAT = WM_CAP_START + 41
Private Const WM_CAP_DLG_VIDEOCOMPRESSION = WM_CAP_START + 46
Private Const WM_CAP_DLG_VIDEOSOURCE = WM_CAP_START + 42
Private Const WM_CAP_DRIVER_CONNECT = WM_CAP_START + 10
Private Const WM_CAP_DRIVER_GET_CAPS = WM_CAP_START + 14
Private Const WM_CAP_SET_PREVIEWRATE = WM_CAP_START + 52
Private Const WM_CAP_SET_PREVIEW = WM_CAP_START + 50
Private Const WM_CAP_GET_STATUS = WM_CAP_START + 54
Private Const WM_CAP_GRAB_FRAME = WM_CAP_START + 60
Private Const WM_CAP_FILE_SAVEDIB = WM_CAP_START + 25
Private Const WM_CAP_UNICODE_START As Long = WM_USER + 100
Private Const WM_CAP_FILE_SAVEDIBW As Long = (WM_CAP_UNICODE_START + 25)

Private Type POINTAPI
     X As Long
     Y As Long
End Type

Private Type CAPDRIVERCAPS
     wDeviceIndex As Long
     fHasOverlay As Long
     fHasDlgVideoSource As Long
     fHasDlgVideoFormat As Long
     fHasDlgVideoDisplay As Long
     fCaptureInitialized As Long
     fDriverSuppliesPalettes As Long
     hVideoIn As Long
     hVideoOut As Long
     hVideoExtIn As Long
     hVideoExtOut As Long
End Type

Private Type CAPSTATUS
     uiImageWidth As Long
     uiImageHeight As Long
     fLiveWindow As Long
     fOverlayWindow As Long
     fScale As Long
     ptScroll As POINTAPI
     fUsingDefaultPalette As Long
     fAudioHardware As Long
     fCapFileExists As Long
     dwCurrentVideoFrame As Long
     dwCurrentVideoFramesDropped As Long
     dwCurrentWaveSamples As Long
     dwCurrentTimeElapsedMS As Long
     hPalCurrent As Long
     fCapturingNow As Long
     dwReturn As Long
     wNumVideoAllocated As Long
     wNumAudioAllocated As Long
End Type

Private Type SECURITY_ATTRIBUTES
     nLength As Long
     lpSecurityDescriptor As Long
     bInheritHandle As Long
End Type

Private Declare Function capGetDriverDescription Lib "avicap32.dll" Alias "capGetDriverDescriptionA" (ByVal dwDriverIndex As Long, ByVal lpszName As String, ByVal cbName As Long, ByVal lpszVer As String, ByVal cbVer As Long) As Long
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long
Private Declare Function SendMessage_Long Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage_Any Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SendMessage_String Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

Dim CapHWnd As Long 'Í¼Ïñ³éÈ¡´°¿ÚµÄHWnd handle of video source window

Sub Get_CaptureWindow(Optional ByVal nm As Long = 0)
'create a capture window
'nm is the id of the camera(default is 0)
'½¨Á¢Ò»¸ö¿É³éÈ¡µÄ´°¿Ú
'nm Îª,Èô²»Ö»Ò»¸ö³éÈ¡×°ÖÃµÄ»°,Ö¸¶¨×°ÖÃ´úºÅ
CapHWnd = capCreateCaptureWindow("", WS_CHILD Or WS_VISIBLE, 0, 0, 160, 120, Screen.ActiveForm.hwnd, 0)
Call Connect_CaptureDriver(nm)
End Sub

Function Connect_CaptureDriver(ByVal nDriverIndex As Long) As Boolean
'link to the camera
'Á´½Óµ½³éÈ¡×°ÖÃ
Dim retVal As Boolean
Dim Caps As CAPDRIVERCAPS
Dim I As Long
'Debug.Assert (nDriverIndex < 10) And (nDriverIndex >= 0)
'link to the interface of video source window
'Á´½Óµ½³éÈ¡´°¿ÚµÄ½çÃæ
retVal = SendMessage_Long(CapHWnd, WM_CAP_DRIVER_CONNECT, nDriverIndex, 0&)
If retVal = False Then Exit Function
'return the ability of capture interface
'·µ»Ø³éÈ¡½çÃæµÄÄÜÁ¦
retVal = SendMessage_Any(CapHWnd, WM_CAP_DRIVER_GET_CAPS, Len(Caps), Caps)
'set the rate of preview (per millisecond)
'ÉèÖÃÃ¿ºÁÃëÔ¤ÀÀµÄËÙ¶È
Call Set_PreviewRate(CapHWnd, 66) '15 FPS
'activate the preview of camera
'¼¤»îÉãÓ°»úµÄÔ¤ÀÀÍ¼Ïñ
Call Set_Preview
'readjust the capture window to the full image size
'ÖØÐÂµ÷Õû³éÈ¡´°¿ÚÎªÈ«²¿Õ¼ÂúÍ¼Ïñ
Call ResizeCaptureWindow
Connect_CaptureDriver = True
End Function

Private Function Set_PreviewRate(ByVal hCapWnd As Long, ByVal wMS As Long) As Boolean
'set the rate of preview (per millisecond)
'ÉèÖÃÃ¿ºÁÃëÔ¤ÀÀµÄËÙ¶È
Set_PreviewRate = SendMessage_Long(hCapWnd, WM_CAP_SET_PREVIEWRATE, wMS, 0&)
End Function

Private Function Set_Preview() As Boolean
'activate the preview of camera
'¼¤»îÉãÓ°»úµÄÔ¤ÀÀÍ¼Ïñ
Set_Preview = SendMessage_Long(CapHWnd, WM_CAP_SET_PREVIEW, True, 0&)
End Function

Sub ResizeCaptureWindow()
'readjust the size of capture window
'ÖØÐÂµ÷Õû³éÈ¡´°¿ÚµÄ´óÐ¡
Dim B As Boolean
Dim capStat As CAPSTATUS
'return the capture window's status
'·µ»Ø³éÈ¡´°¿ÚµÄ×´Ì¬
B = Get_CaptureWindow_Status(CapHWnd, capStat)
If B = True Then
    'readjust the size of capture window
    'ÖØÐÂµ÷Õû³éÈ¡´°¿ÚµÄ´óÐ¡
    Call SetWindowPos(CapHWnd, 0&, 0&, 0&, capStat.uiImageWidth, capStat.uiImageHeight, SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSENDCHANGING)
    Screen.ActiveForm.Width = capStat.uiImageWidth * Screen.TwipsPerPixelX * 2 + 120
    Screen.ActiveForm.Height = capStat.uiImageHeight * Screen.TwipsPerPixelY + 780
    
    Form1.P1.Left = capStat.uiImageWidth
    Form1.P1.Width = capStat.uiImageWidth
    Form1.P1.Height = capStat.uiImageHeight
    
End If
End Sub

Function Get_CaptureWindow_Status(ByVal hCapWnd As Long, ByRef capStat As CAPSTATUS) As Boolean
'return the capture window's status
'·µ»Ø³éÈ¡´°¿ÚµÄ×´Ì¬
Get_CaptureWindow_Status = SendMessage_Any(hCapWnd, WM_CAP_GET_STATUS, Len(capStat), capStat)
End Function

Function Set_VideoFormat() As Boolean
'set the capture image's resolution
'ÉèÖÃ³éÈ¡»­ÃæµÄ·Ö±æÂÊ
Set_VideoFormat = SendMessage_Long(CapHWnd, WM_CAP_DLG_VIDEOFORMAT, 0&, 0&)
Call ResizeCaptureWindow
End Function

Function Get_SimpleWindow(arr() As Byte) As Boolean
'capture single image
'n is the filename
'³éÈ¡µ¥»­Ãæ
'n ÎªÎÄ¼þÃû
Dim B As Long, n As String, FrFl As Integer
n = ".\CAP.BMP"
Call SendMessage_Long(CapHWnd, WM_CAP_GRAB_FRAME, 0&, 0&)
B = SendMessage_String(CapHWnd, WM_CAP_FILE_SAVEDIB, 0&, ByVal n) 'Ascii ·½Ê½ ASCII method
'B = SendMessage_String(CapHWnd, WM_CAP_FILE_SAVEDIBW, 0&, ByVal StrConv(n, vbUnicode)) 'Unicode ·½Ê½ Unicode method
If B <> 0 And FileLen(n) <> 0 Then
    FrFl = FreeFile
    Open n For Binary Access Read As #FrFl
    ReDim arr(LOF(FrFl) - 1) As Byte
    Get #FrFl, , arr
    Close #FrFl
End If
'the following line is for preventing from freezing after captured
Call Set_Preview '¼ÓÕâÒ»ÐÐ,²Å²»»á³éÈ¡áá,»­Ãæ³ÊÏÖ¶³½á(Freeze)×´Ì¬
Get_SimpleWindow = B
End Function

Function Set_CaptureSource() As Boolean
'set the capture source camera
'ÉèÖÃ³éÈ¡Ô´
Set_CaptureSource = SendMessage_Long(CapHWnd, WM_CAP_DLG_VIDEOSOURCE, 0&, 0&)
End Function

Sub DrawCap(arrDIBs() As Byte)
Dim capStat As CAPSTATUS
Dim hCapDC As Long, B As Boolean
hCapDC = GetDC(CapHWnd)
B = Get_CaptureWindow_Status(CapHWnd, capStat)
If B Then
    'BitBlt Form1.P1.hDC, 0, 0, capStat.uiImageWidth, capStat.uiImageHeight, hCapDC, 0, 0, vbSrcCopy
    'Form1.P1.Refresh
    Call SaveDIB(hCapDC, 8, arrDIBs)
End If
End Sub

'Function Set_CompressRate() As Boolean
'ÉèÖÃ³éÑùËÙÂÊ
'Set_CompressRate = SendMessage_Long(CapHWnd, WM_CAP_DLG_VIDEOCOMPRESSION, 0&, 0&)
'End Function
