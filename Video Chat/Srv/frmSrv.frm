VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSrv 
   AutoRedraw      =   -1  'True
   Caption         =   "·þÎñ¶Ë Server"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   8130
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin MSWinsockLib.Winsock Wsk1 
      Left            =   3960
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   8686
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Í¼ÏóÑ¹Ëõ"
      Height          =   180
      Left            =   3360
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   4455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "¶Ô·½Í¼Ïó"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "±¾µØÍ¼Ïó"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "¶Ï¿ªÁ¬½Ó"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmSrv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'I'm a Chinese undergraduate student
'excuse my poor English ~_~!
'Code By TZWSOHO

Private Type ImageInfo
    imgWidth As Long '¿í¶È width
    imgHeight As Long '¸ß¶È height
    imgOrgSize As Long 'Í¼ÏóÔ­´óÐ¡ original size
    imgCmpSize As Long 'Í¼ÏóÑ¹Ëõºó´óÐ¡ compressed size
    lBitCount As Long 'Í¼ÏóÎ»É«Êý
    lPtr As Long 'Ñ¹ËõºóµÄÖ¸Õë compressed pointer
End Type

Private Declare Function SetDIBitsToDevice Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, ByRef Bits As Any, ByRef BitsInfo As BitMapInfo256, ByVal wUsage As Long) As Long

Private Sub Command1_Click()
List1.AddItem "ÓÃ»§¶Ï¿ªÁ¬½Ó£¡" 'user disconnected
Wsk1.Close
Wsk1.Listen
Command1.Enabled = False
Command3.Enabled = False
Unload frmSelf
Unload frmOpp
End Sub

Private Sub Command2_Click()
frmSelf.Show
Call Get_CaptureWindow(0)
End Sub

Private Sub Command3_Click()
frmOpp.Show
Wsk1.SendData Chr$(0) '·¢ËÍ»ñÈ¡Ô¶³Ì¼ÆËã»úÊÓÆµÍ¼ÏóµÄÃüÁî request to getting video
Command3.Enabled = False
End Sub

Private Sub Form_Load()
Wsk1.Listen
List1.AddItem "±¾µØ¶Ë¿Ú '8686' ÒÑ¾­´ò¿ª£¬³ÌÐòÕýÔÚ¼àÌý..." 'listening
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmSelf
Unload frmOpp
Unload frmSrv
End
End Sub

Private Sub Wsk1_Close()
List1.AddItem "Ô¶³Ì¼ÆËã»ú¶Ï¿ªÁ¬½Ó£¡" 'disconnected
Wsk1.Close
Wsk1.Listen
Command1.Enabled = False
Command3.Enabled = False
Unload frmSelf
Unload frmOpp
End Sub

Private Sub Wsk1_ConnectionRequest(ByVal requestID As Long)
'connection request
List1.AddItem "Ô¶³Ì¼ÆËã»ú '" & Wsk1.RemoteHostIP & "' ·¢À´Á¬½ÓÇëÇó£¬´úºÅ '" & requestID & "' ..."
List1.AddItem "ÒÑ¾­½¨Á¢Á¬½Ó£¡" 'connection established
Wsk1.Close
Wsk1.Accept requestID
Command1.Enabled = True
Command3.Enabled = True
End Sub

Private Sub Wsk1_DataArrival(ByVal bytesTotal As Long)
'arrDIB ±¾µØÊÓÆµÍ¼ÏóµÄÐÅÏ¢ local video data
'arrDIBRec ½ÓÊÕµ½µÄÔ¶³Ì¼ÆËã»úÊÓÆµÍ¼ÏóÐÅÏ¢ received video data
On Error GoTo er
Static bRecDIB As Boolean, dstImgInfo As ImageInfo
Static arrDIB() As Byte, arrDIBRec() As Byte
Dim arrData() As Byte, I As Long
Dim srcImgInfo As ImageInfo, BMPInfo As BitMapInfo256
Wsk1.GetData arrData, vbArray Or vbByte
If bRecDIB Then
    If ArrIsNull(arrDIBRec) Then
        ReDim arrDIBRec(UBound(arrData))
        Call CopyMemory(arrDIBRec(0), arrData(0), UBound(arrData) + 1)
    Else
        With dstImgInfo
            If UBound(arrDIBRec) < .imgCmpSize - UBound(arrData) - 1 Then
                I = UBound(arrDIBRec) + 1
                ReDim Preserve arrDIBRec(I + UBound(arrData))
                Call CopyMemory(arrDIBRec(I), arrData(0), UBound(arrData) + 1)
            Else
                'received the last array of data and make it
                'the background of frmOpp form
                '½ÓÊÕ×îºóÒ»×éÊý¾Ý²¢´òÓ¡³É frmOpp µÄ±³¾°
                I = UBound(arrDIBRec) + 1
                ReDim Preserve arrDIBRec(I + UBound(arrData))
                Call CopyMemory(arrDIBRec(I), arrData(0), UBound(arrData) + 1)
                BMPInfo = CreateBMInfo(.imgWidth, .imgHeight, .lBitCount)
                If .imgCmpSize < .imgOrgSize Then
                    DoEvents
                    Call Compress(arrDIBRec, .lPtr, .imgOrgSize) '½âÑ¹ËõÍ¼Ïó decompress video
                End If
                Call SetDIBitsToDevice(frmOpp.hdc, 0, 0, .imgWidth, .imgHeight, 0, 0, 0, .imgHeight, arrDIBRec(0), BMPInfo, DIB_RGB_COLORS)
                DoEvents '²»¼ÓÕâ¾ä»­Ãæ½«»á²»Á÷³© make the video fluent
                Erase arrDIBRec: bRecDIB = False
                Wsk1.SendData Chr$(1) '¼ÌÐø½ÓÊÕÏÂÒ»·ùÍ¼Ïó receive the next image
            End If
        End With
    End If
Else
    Select Case arrData(0)
        Case 0 'Ô¶³Ì¼ÆËã»úÒªÇó½ÓÊÕ±¾µØ¼ÆËã»úµÄÊÓÆµÍ¼Ïó requestion of receiving local video
            List1.AddItem "Ô¶³Ì¼ÆËã»ú¿ªÊ¼½ÓÊÕ±¾µØÊÓÆµÍ¼Ïó..."
            frmSelf.Show
            Call Get_CaptureWindow(0)
            Call Get_CaptureDIB(arrDIB, 8) '256É«
            With srcImgInfo
                .lBitCount = 8
                .imgOrgSize = UBound(arrDIB)
                .imgWidth = frmSelf.ScaleWidth
                .imgHeight = frmSelf.ScaleHeight
                
                'the following line is for compressing the video
                'ÏÂÃæÕâÒ»ÐÐÊÇÑ¹ËõÍ¼ÏóµÄ¹ý³Ì£¬
                '¾­ÎÒ²âÊÔÊ¹ÓÃÏÂÃæÒ»ÐÐºó CPU Õ¼ÓÃÂÊ±ÈÃ»Ê¹ÓÃÊ±¸ß 10% ×óÓÒ
                '²»¹ýËÆºõ¶Ô·½µÄÍ¼Ïó»áºÜ¿¨£¬²»Ì«ÍÆ¼öÊ¹ÓÃ
                
                If CBool(Check1.Value) Then Call Compress(arrDIB, .lPtr, UBound(arrDIB)) 'Ñ¹ËõÍ¼Ïó
                
                .imgCmpSize = UBound(arrDIB) 'Ñ¹ËõºóµÄÍ¼Ïó´óÐ¡£¨Ã»Ñ¹Ëõ¾ÍµÈÓÚ imgOrgSize£©
                '·¢ËÍÍ¼ÏóÐÅÏ¢ send the video data
                Wsk1.SendData Chr$(2) & CStr(.lBitCount) & CStr(.lPtr) & "|" & CStr(.imgOrgSize) & "|" & CStr(.imgCmpSize) & "|" & CStr(.imgWidth) & "|" & CStr(.imgHeight)
            End With
        Case 1 'Ô¶³Ì¼ÆËã»ú¼ÌÐø¿ªÊ¼½ÓÊÕÍ¼Ïó continue receiving video
            Call Get_CaptureDIB(arrDIB, 8) '256É«
            If ArrIsNull(arrDIB) Then Exit Sub
            With srcImgInfo
                .lBitCount = 8
                .imgOrgSize = UBound(arrDIB)
                .imgWidth = frmSelf.ScaleWidth
                .imgHeight = frmSelf.ScaleHeight
                If CBool(Check1.Value) Then Call Compress(arrDIB, .lPtr, UBound(arrDIB)) 'Ñ¹ËõÍ¼Ïó
                .imgCmpSize = UBound(arrDIB)
                '·¢ËÍÍ¼ÏóÐÅÏ¢ send video data
                Wsk1.SendData Chr$(2) & CStr(.lBitCount) & CStr(.lPtr) & "|" & CStr(.imgOrgSize) & "|" & CStr(.imgCmpSize) & "|" & CStr(.imgWidth) & "|" & CStr(.imgHeight)
            End With
        Case 2 'Ô¶³Ì¼ÆËã»ú·¢À´Í¼ÏóÐÅÏ¢ remote video data
            Dim v As Variant
            v = Split(Mid(StrConv(arrData, vbUnicode), 3), "|")
            With dstImgInfo
                .lBitCount = CByte(Mid(StrConv(arrData, vbUnicode), 2, 1))
                .lPtr = CLng(v(0))
                .imgOrgSize = CLng(v(1))
                .imgCmpSize = CLng(v(2))
                .imgWidth = CLng(v(3))
                .imgHeight = CLng(v(4))
                If frmOpp.WindowState = vbMinimized Then
                    Wsk1.SendData Chr$(8)
                    Exit Sub
                Else
                    frmOpp.Width = .imgWidth * Screen.TwipsPerPixelX + 90
                    frmOpp.Height = .imgHeight * Screen.TwipsPerPixelY + 510
                End If
            End With
            bRecDIB = True
            Wsk1.SendData Chr$(3) 'ÒÑ¾­×¼±¸ºÃ½ÓÊÕÊÓÆµÍ¼Ïó ready to receive
        Case 3 'Ô¶³Ì¼ÆËã»úÒªÇó½ÓÊÕ±¾µØÊÓÆµÍ¼Ïó remote requestion of receiving video
            Wsk1.SendData arrDIB
            Erase arrDIB
        Case 4 'Ô¶³Ì¼ÆËã»úÍ£Ö¹½ÓÊÕÊÓÆµÍ¼Ïó remote stop receiving video
            List1.AddItem "Ô¶³Ì¼ÆËã»úÍ£Ö¹½ÓÊÕÊÓÆµÍ¼Ïó"
            Erase arrDIB
            Unload frmSelf
        Case 5 'Ô¶³Ì¼ÆËã»úÍ£Ö¹·¢ËÍÍ¼Ïó remote stop sending video
            List1.AddItem "Ô¶³Ì¼ÆËã»úÍ£Ö¹·¢ËÍÍ¼Ïó"
            Erase arrDIBRec
            Unload frmOpp
        Case 6 'Ô¶³Ì¼ÆËã»úÔÝÍ£·¢ËÍÍ¼Ïó remote pause on receiving video
            List1.AddItem "Ô¶³Ì¼ÆËã»úÔÝÍ£·¢ËÍÍ¼Ïó"
            Erase arrDIBRec
        Case 7 'Ô¶³Ì¼ÆËã»úÑ¯ÎÊÊÇ·ñ¼ÌÐø½ÓÊÕÍ¼ÏóÐÅÏ¢ remote ask if continue receiving video
            frmOpp.Show
            'continue receiving remote video
            Wsk1.SendData Chr$(1) '·¢ËÍ¼ÌÐø»ñÈ¡Ô¶³Ì¼ÆËã»úÊÓÆµÍ¼ÏóµÄÃüÁî
            Command3.Enabled = False
        Case 8 'ÔÝÍ£½ÓÊÕ±¾µØÍ¼Ïó stop receiving local video
            List1.AddItem "ÔÝÍ£½ÓÊÕ±¾µØÍ¼Ïó"
            Erase arrDIB
    End Select
End If
Exit Sub
er:
List1.AddItem "´íÎó:" & Err.Description
Debug.Print Err.Description
End Sub

Private Function ArrIsNull(arr() As Byte) As Boolean
On Error GoTo er
Dim I As Long
I = UBound(arr)
ArrIsNull = False
Exit Function
er:
ArrIsNull = True
End Function
