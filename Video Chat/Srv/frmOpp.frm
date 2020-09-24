VERSION 5.00
Begin VB.Form frmOpp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "对方图象"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   4275
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "frmOpp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'I'm a Chinese undergraduate student
'excuse my poor English ~_~!
'Code By TZWSOHO

Private Sub Form_DblClick()
'captured to the App.Path
If Get_SimpleWindow Then frmSrv.List1.AddItem "画面已经截取到程序目录！"
End Sub

Private Sub Form_Resize()
On Error Resume Next
Static bSend As Boolean
If frmOpp.WindowState = vbMinimized Then
    bSend = True
    frmSrv.Wsk1.SendData Chr$(8) '暂停接收图象信息 pause on receiving video
ElseIf frmOpp.WindowState = vbNormal Then
    If bSend Then
        bSend = False
        frmSrv.Wsk1.SendData Chr$(1) '继续接收图象信息 continue receiving video
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
frmSrv.Command3.Enabled = True
frmSrv.Wsk1.SendData Chr(4) '发送停止接收视频图象的命令 stop receiving video
End Sub
