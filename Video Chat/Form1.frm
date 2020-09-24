VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   4320
      Top             =   1080
   End
   Begin VB.PictureBox P1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   4800
      ScaleHeight     =   3615
      ScaleWidth      =   4800
      TabIndex        =   0
      Top             =   0
      Width           =   4800
   End
   Begin VB.Menu mnu_Option 
      Caption         =   "Ñ¡Ïî(&O)"
      Begin VB.Menu mnu_ImageFormat 
         Caption         =   "±ä¸üÍ¼Ïó¸ñÊ½..."
      End
      Begin VB.Menu mnu_CaptureSource 
         Caption         =   "ÉèÖÃ³éÈ¡À´Ô´..."
      End
      Begin VB.Menu mnu_Compression 
         Caption         =   "Ñ¹Ëõ±È..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnu_Capture 
      Caption         =   "³éÈ¡(&C)"
      Begin VB.Menu mnu_BMP 
         Caption         =   "µ¥»­Ãæ"
      End
   End
   Begin VB.Menu mnu_Compress1 
      Caption         =   "Ñ¹ËõËã·¨Ò»"
   End
   Begin VB.Menu mnu_Compress2 
      Caption         =   "Ñ¹ËõËã·¨¶þ(Huffman)"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'I'm a Chinese undergraduate student
'excuse my poor English ~_~!
'Code By TZWSOHO

Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BitMapInfo256, ByVal wUsage As Long) As Long

Private bWay As Boolean

Private Sub Form_Load()
Me.Show
Call Get_CaptureWindow(0)
End Sub

Private Sub mnu_BMP_Click()
On Error Resume Next
Dim arrBMP() As Byte
Call Get_SimpleWindow(arrBMP()) '½ØÈ¡µ±Ç°Í¼Ïóµ½³ÌÐòÄ¿Â¼µÄ CAP.BMP capture the current image to ".\CAP.BMP"
'Debug.Print Chr(arrBMP(0)), Chr(arrBMP(1))
End Sub

Private Sub mnu_CaptureSource_Click()
Call Set_CaptureSource
End Sub

Private Sub mnu_Compress1_Click()
bWay = False
End Sub

Private Sub mnu_Compress2_Click()
bWay = True
End Sub

'Private Sub mnu_Compression_Click()
'Call Set_CompressRate
'End Sub

Private Sub mnu_ImageFormat_Click()
Call Set_VideoFormat
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
'DIBs() Ã¿Ò»·ùÍ¼ÏóµÄÄÚÈÝ data of every image
'DIBPtr Í¼ÏóÑ¹ËõºóµÄÖ¸Õë pointer of compressed image
'I Ô­Í¼ÏóµÄ´óÐ¡ original size of image
Dim I As Long, DIBPtr As Long
Dim DIBs() As Byte, BMPInfo As BitMapInfo256
Call DrawCap(DIBs) '»ñÈ¡Í¼ÏóÄÚÈÝ get data of image
BMPInfo = CreateBMInfo(P1.Width, P1.Height, 8)
I = UBound(DIBs)

If Not bWay Then
    'the first compressing method
    'µÚÒ»ÖÖÑ¹Ëõ·½·¨
    Call Compress(DIBs, DIBPtr, UBound(DIBs)) 'Ñ¹Ëõ compress
    Me.Caption = "´óÐ¡(×Ö½Ú):Ô­Í¼Ïó/Ñ¹Ëõºó " & I & "/" & UBound(DIBs) & " (Ëã·¨Ò»)"
    Call Compress(DIBs, DIBPtr, I) '½âÑ¹ decompress
Else
    'the second compressing method(huffman)
    'µÚ¶þÖÖÑ¹Ëõ·½·¨ (»ô·òÂüËã·¨)
    Call Compress_Huffman_Dynamic(DIBs) 'Ñ¹Ëõ compress
    Me.Caption = "´óÐ¡(×Ö½Ú):Ô­Í¼Ïó/Ñ¹Ëõºó " & I & "/" & UBound(DIBs) & " (Ëã·¨¶þ)"
    Call DeCompress_Huffman_Dynamic(DIBs) '½âÑ¹ decompress
End If

Call SetDIBitsToDevice(P1.hDC, 0, 0, P1.Width, P1.Height, 0, 0, 0, P1.Height, DIBs(0), BMPInfo, DIB_RGB_COLORS)
End Sub
