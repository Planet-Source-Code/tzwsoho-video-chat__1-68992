VERSION 5.00
Begin VB.Form frmOpp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "¶Ô·½Í¼Ïó"
   ClientHeight    =   3600
   ClientLeft      =   10515
   ClientTop       =   7410
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
If Get_SimpleWindow Then frmClt.List1.AddItem "»­ÃæÒÑ¾­½ØÈ¡µ½³ÌÐòÄ¿Â¼£¡"
End Sub

Private Sub Form_Resize()
On Error Resume Next
Static bSend As Boolean
If frmOpp.WindowState = vbMinimized Then
    bSend = True
    frmClt.Wsk1.SendData Chr$(8) 'ÔÝÍ£½ÓÊÕÍ¼ÏóÐÅÏ¢ pause on receiving video
ElseIf frmOpp.WindowState = vbNormal Then
    If bSend Then
        bSend = False
        frmClt.Wsk1.SendData Chr$(1) '¼ÌÐø½ÓÊÕÍ¼ÏóÐÅÏ¢ continue receiving video
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
frmClt.Command3.Enabled = True
frmClt.Wsk1.SendData Chr(4) '·¢ËÍÍ£Ö¹½ÓÊÕÊÓÆµÍ¼ÏóµÄÃüÁî stop receive video
End Sub
