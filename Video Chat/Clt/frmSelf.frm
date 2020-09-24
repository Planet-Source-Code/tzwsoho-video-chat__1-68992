VERSION 5.00
Begin VB.Form frmSelf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "自己图象"
   ClientHeight    =   3600
   ClientLeft      =   10515
   ClientTop       =   3870
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Menu mnu_Option 
      Caption         =   "选项(&O)"
      Begin VB.Menu mnu_VideoFormat 
         Caption         =   "改变视频格式..."
      End
      Begin VB.Menu mnu_VideoSource 
         Caption         =   "选择视频截取来源..."
      End
   End
End
Attribute VB_Name = "frmSelf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'I'm a Chinese undergraduate student
'excuse my poor English ~_~!
'Code By TZWSOHO

Private Sub Form_Resize()
On Error Resume Next
Static bSend As Boolean
If frmSelf.WindowState = vbMinimized Then
    bSend = True
    frmClt.Wsk1.SendData Chr$(6) '本地停止发送图象信息 stop sending video
ElseIf frmSelf.WindowState = vbNormal Then
    If bSend Then
        bSend = False
        'ask if continue receiving video
        frmClt.Wsk1.SendData Chr$(7) '询问远程计算机是否继续接收图象信息
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
frmClt.Wsk1.SendData Chr$(5) '停止发送图象信息 stop sending video
End Sub

Private Sub mnu_VideoFormat_Click()
Call Set_VideoFormat
End Sub

Private Sub mnu_VideoSource_Click()
Call Set_CaptureSource
End Sub
