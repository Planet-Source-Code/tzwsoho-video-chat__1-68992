VERSION 5.00
Begin VB.Form frmSelf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "×Ô¼ºÍ¼Ïó"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Menu mnu_Option 
      Caption         =   "Ñ¡Ïî(&O)"
      Begin VB.Menu mnu_VideoFormat 
         Caption         =   "¸Ä±äÊÓÆµ¸ñÊ½..."
      End
      Begin VB.Menu mnu_VideoSource 
         Caption         =   "Ñ¡ÔñÊÓÆµ½ØÈ¡À´Ô´..."
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
    frmSrv.Wsk1.SendData Chr$(6) '±¾µØÍ£Ö¹·¢ËÍÍ¼ÏóÐÅÏ¢ stop sending video
ElseIf frmSelf.WindowState = vbNormal Then
    If bSend Then
        bSend = False
        'ask if continue receiving video
        frmSrv.Wsk1.SendData Chr$(7) 'Ñ¯ÎÊÔ¶³Ì¼ÆËã»úÊÇ·ñ¼ÌÐø½ÓÊÕÍ¼ÏóÐÅÏ¢
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
frmSrv.Wsk1.SendData Chr$(5) 'Í£Ö¹·¢ËÍÍ¼ÏóÐÅÏ¢ stop sending video
End Sub

Private Sub mnu_VideoFormat_Click()
Call Set_VideoFormat
End Sub

Private Sub mnu_VideoSource_Click()
Call Set_CaptureSource
End Sub
