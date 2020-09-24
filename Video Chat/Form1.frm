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
   StartUpPosition =   3  '窗口缺省
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
      Caption         =   "选项(&O)"
      Begin VB.Menu mnu_ImageFormat 
         Caption         =   "变更图象格式..."
      End
      Begin VB.Menu mnu_CaptureSource 
         Caption         =   "设置抽取来源..."
      End
      Begin VB.Menu mnu_Compression 
         Caption         =   "压缩比..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnu_Capture 
      Caption         =   "抽取(&C)"
      Begin VB.Menu mnu_BMP 
         Caption         =   "单画面"
      End
   End
   Begin VB.Menu mnu_Compress1 
      Caption         =   "压缩算法一"
   End
   Begin VB.Menu mnu_Compress2 
      Caption         =   "压缩算法二(Huffman)"
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
Call Get_SimpleWindow(arrBMP()) '截取当前图象到程序目录的 CAP.BMP capture the current image to ".\CAP.BMP"
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
'DIBs() 每一幅图象的内容 data of every image
'DIBPtr 图象压缩后的指针 pointer of compressed image
'I 原图象的大小 original size of image
Dim I As Long, DIBPtr As Long
Dim DIBs() As Byte, BMPInfo As BitMapInfo256
Call DrawCap(DIBs) '获取图象内容 get data of image
BMPInfo = CreateBMInfo(P1.Width, P1.Height, 8)
I = UBound(DIBs)

If Not bWay Then
    'the first compressing method
    '第一种压缩方法
    Call Compress(DIBs, DIBPtr, UBound(DIBs)) '压缩 compress
    Me.Caption = "大小(字节):原图象/压缩后 " & I & "/" & UBound(DIBs) & " (算法一)"
    Call Compress(DIBs, DIBPtr, I) '解压 decompress
Else
    'the second compressing method(huffman)
    '第二种压缩方法 (霍夫曼算法)
    Call Compress_Huffman_Dynamic(DIBs) '压缩 compress
    Me.Caption = "大小(字节):原图象/压缩后 " & I & "/" & UBound(DIBs) & " (算法二)"
    Call DeCompress_Huffman_Dynamic(DIBs) '解压 decompress
End If

Call SetDIBitsToDevice(P1.hDC, 0, 0, P1.Width, P1.Height, 0, 0, 0, P1.Height, DIBs(0), BMPInfo, DIB_RGB_COLORS)
End Sub
