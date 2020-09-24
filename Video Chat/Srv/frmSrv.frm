VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSrv 
   AutoRedraw      =   -1  'True
   Caption         =   "服务端 Server"
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
      Caption         =   "图象压缩"
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
      Caption         =   "对方图象"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "本地图象"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "断开连接"
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
    imgWidth As Long '宽度 width
    imgHeight As Long '高度 height
    imgOrgSize As Long '图象原大小 original size
    imgCmpSize As Long '图象压缩后大小 compressed size
    lBitCount As Long '图象位色数
    lPtr As Long '压缩后的指针 compressed pointer
End Type

Private Declare Function SetDIBitsToDevice Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, ByRef Bits As Any, ByRef BitsInfo As BitMapInfo256, ByVal wUsage As Long) As Long

Private Sub Command1_Click()
List1.AddItem "用户断开连接！" 'user disconnected
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
Wsk1.SendData Chr$(0) '发送获取远程计算机视频图象的命令 request to getting video
Command3.Enabled = False
End Sub

Private Sub Form_Load()
Wsk1.Listen
List1.AddItem "本地端口 '8686' 已经打开，程序正在监听..." 'listening
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmSelf
Unload frmOpp
Unload frmSrv
End
End Sub

Private Sub Wsk1_Close()
List1.AddItem "远程计算机断开连接！" 'disconnected
Wsk1.Close
Wsk1.Listen
Command1.Enabled = False
Command3.Enabled = False
Unload frmSelf
Unload frmOpp
End Sub

Private Sub Wsk1_ConnectionRequest(ByVal requestID As Long)
'connection request
List1.AddItem "远程计算机 '" & Wsk1.RemoteHostIP & "' 发来连接请求，代号 '" & requestID & "' ..."
List1.AddItem "已经建立连接！" 'connection established
Wsk1.Close
Wsk1.Accept requestID
Command1.Enabled = True
Command3.Enabled = True
End Sub

Private Sub Wsk1_DataArrival(ByVal bytesTotal As Long)
'arrDIB 本地视频图象的信息 local video data
'arrDIBRec 接收到的远程计算机视频图象信息 received video data
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
                '接收最后一组数据并打印成 frmOpp 的背景
                I = UBound(arrDIBRec) + 1
                ReDim Preserve arrDIBRec(I + UBound(arrData))
                Call CopyMemory(arrDIBRec(I), arrData(0), UBound(arrData) + 1)
                BMPInfo = CreateBMInfo(.imgWidth, .imgHeight, .lBitCount)
                If .imgCmpSize < .imgOrgSize Then
                    DoEvents
                    Call Compress(arrDIBRec, .lPtr, .imgOrgSize) '解压缩图象 decompress video
                End If
                Call SetDIBitsToDevice(frmOpp.hdc, 0, 0, .imgWidth, .imgHeight, 0, 0, 0, .imgHeight, arrDIBRec(0), BMPInfo, DIB_RGB_COLORS)
                DoEvents '不加这句画面将会不流畅 make the video fluent
                Erase arrDIBRec: bRecDIB = False
                Wsk1.SendData Chr$(1) '继续接收下一幅图象 receive the next image
            End If
        End With
    End If
Else
    Select Case arrData(0)
        Case 0 '远程计算机要求接收本地计算机的视频图象 requestion of receiving local video
            List1.AddItem "远程计算机开始接收本地视频图象..."
            frmSelf.Show
            Call Get_CaptureWindow(0)
            Call Get_CaptureDIB(arrDIB, 8) '256色
            With srcImgInfo
                .lBitCount = 8
                .imgOrgSize = UBound(arrDIB)
                .imgWidth = frmSelf.ScaleWidth
                .imgHeight = frmSelf.ScaleHeight
                
                'the following line is for compressing the video
                '下面这一行是压缩图象的过程，
                '经我测试使用下面一行后 CPU 占用率比没使用时高 10% 左右
                '不过似乎对方的图象会很卡，不太推荐使用
                
                If CBool(Check1.Value) Then Call Compress(arrDIB, .lPtr, UBound(arrDIB)) '压缩图象
                
                .imgCmpSize = UBound(arrDIB) '压缩后的图象大小（没压缩就等于 imgOrgSize）
                '发送图象信息 send the video data
                Wsk1.SendData Chr$(2) & CStr(.lBitCount) & CStr(.lPtr) & "|" & CStr(.imgOrgSize) & "|" & CStr(.imgCmpSize) & "|" & CStr(.imgWidth) & "|" & CStr(.imgHeight)
            End With
        Case 1 '远程计算机继续开始接收图象 continue receiving video
            Call Get_CaptureDIB(arrDIB, 8) '256色
            If ArrIsNull(arrDIB) Then Exit Sub
            With srcImgInfo
                .lBitCount = 8
                .imgOrgSize = UBound(arrDIB)
                .imgWidth = frmSelf.ScaleWidth
                .imgHeight = frmSelf.ScaleHeight
                If CBool(Check1.Value) Then Call Compress(arrDIB, .lPtr, UBound(arrDIB)) '压缩图象
                .imgCmpSize = UBound(arrDIB)
                '发送图象信息 send video data
                Wsk1.SendData Chr$(2) & CStr(.lBitCount) & CStr(.lPtr) & "|" & CStr(.imgOrgSize) & "|" & CStr(.imgCmpSize) & "|" & CStr(.imgWidth) & "|" & CStr(.imgHeight)
            End With
        Case 2 '远程计算机发来图象信息 remote video data
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
            Wsk1.SendData Chr$(3) '已经准备好接收视频图象 ready to receive
        Case 3 '远程计算机要求接收本地视频图象 remote requestion of receiving video
            Wsk1.SendData arrDIB
            Erase arrDIB
        Case 4 '远程计算机停止接收视频图象 remote stop receiving video
            List1.AddItem "远程计算机停止接收视频图象"
            Erase arrDIB
            Unload frmSelf
        Case 5 '远程计算机停止发送图象 remote stop sending video
            List1.AddItem "远程计算机停止发送图象"
            Erase arrDIBRec
            Unload frmOpp
        Case 6 '远程计算机暂停发送图象 remote pause on receiving video
            List1.AddItem "远程计算机暂停发送图象"
            Erase arrDIBRec
        Case 7 '远程计算机询问是否继续接收图象信息 remote ask if continue receiving video
            frmOpp.Show
            'continue receiving remote video
            Wsk1.SendData Chr$(1) '发送继续获取远程计算机视频图象的命令
            Command3.Enabled = False
        Case 8 '暂停接收本地图象 stop receiving local video
            List1.AddItem "暂停接收本地图象"
            Erase arrDIB
    End Select
End If
Exit Sub
er:
List1.AddItem "错误:" & Err.Description
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
