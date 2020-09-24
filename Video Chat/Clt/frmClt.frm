VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "客户端 Client"
   ClientHeight    =   3345
   ClientLeft      =   10800
   ClientTop       =   300
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4560
   Begin VB.CheckBox Check1 
      Caption         =   "图象压缩"
      Height          =   180
      Left            =   3360
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "本地图象"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "对方图象"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   1440
      TabIndex        =   2
      Text            =   "127.0.0.1"
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "断开连接"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "建立连接"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock Wsk1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "远程计算机IP："
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1260
   End
End
Attribute VB_Name = "frmClt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'I'm a Chinese undergraduate student
'excuse my poor English ~_~!
'Code By TZWSOHO

Private Type ImageInfo
    imgWidth As Long '宽度width
    imgHeight As Long '高度height
    imgOrgSize As Long '图象原大小original size
    imgCmpSize As Long '图象压缩后大小compressed size
    lBitCount As Long '图象位色数
    lPtr As Long '压缩后的指针compressed pointer
End Type

Private Declare Function SetDIBitsToDevice Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, ByRef Bits As Any, ByRef BitsInfo As BitMapInfo256, ByVal wUsage As Long) As Long

Private Sub Command1_Click()
Wsk1.Close
Wsk1.Connect Text1.Text, 8686
End Sub

Private Sub Command2_Click()
Wsk1.Close
'disconnect the connection with Wsk1.RemoteHostIP
List1.AddItem "用户断开与 '" & Wsk1.RemoteHostIP & "' 的连接！"
Unload frmSelf
Unload frmOpp
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
End Sub

Private Sub Command3_Click()
frmOpp.Show
'sending the commands of getting the images from the remote machine
Wsk1.SendData Chr$(0) '发送获取远程计算机视频图象的命令
Command3.Enabled = False
End Sub

Private Sub Command4_Click()
frmSelf.Show
Call Get_CaptureWindow(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmOpp
Unload frmSelf
Unload frmClt
End
End Sub

Private Sub Wsk1_Close()
'disconnected with the remote machine
List1.AddItem "已经和远程计算机 '" & Wsk1.RemoteHostIP & "' 断开连接！"
Unload frmSelf
Unload frmOpp
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
End Sub

Private Sub Wsk1_Connect()
'connected the remote machine
List1.AddItem "已经连接上远程计算机 '" & Wsk1.RemoteHostIP & "'！"
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = True
End Sub

Private Sub Wsk1_DataArrival(ByVal bytesTotal As Long)
'arrDIB 本地视频图象的信息 local images data
'arrDIBRec 接收到的远程计算机视频图象信息 received remote images data
On Error GoTo er
Static bRecDIB As Boolean, dstImgInfo As ImageInfo
Static arrDIB() As Byte, arrDIBRec() As Byte, BMPInfo As BitMapInfo256
Dim srcImgInfo As ImageInfo
Dim arrData() As Byte, I As Long
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
                    Call Compress(arrDIBRec, .lPtr, .imgOrgSize) '解压缩图象 decompress images
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
                .imgHeight = frmSelf.ScaleHeight - 20
                
                'the following line is for compressing the video
                '下面这一行是压缩图象的过程，
                '经我测试使用下面一行后 CPU 占用率比没使用时高 10% 左右！
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
                .imgHeight = frmSelf.ScaleHeight - 20
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
        Case 3 '远程计算机要求接收本地视频图象图片 remote requestion of receiving video
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
