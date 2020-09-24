Attribute VB_Name = "Module2"
Option Explicit

'I'm a Chinese undergraduate student
'excuse my poor English ~_~!
'Code By TZWSOHO

Private Const BI_RGB = 0&

Public Const DIB_RGB_COLORS = 0

Private Type BITMAPINFOHEADER 'BMP ÐÅÏ¢Í· BMP information header
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Public Type BitMapInfo256 '16/256É«ÓÃ for 16/256 bit count
     bmiHeader As BITMAPINFOHEADER
     bmiColors(0 To 255) As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BitMapInfo256, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BitMapInfo256, ByVal wUsage As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Sub SaveDIB(ByVal hSrcDC As Long, ByVal BitCount As Long, arrDIBs() As Byte)
'Dim nt As Single: nt = Timer 'for counting time
Dim I As Long
Dim BMInfo As BitMapInfo256
Dim Wth As Long, Hgt As Long
Dim hDIB As Long, iDC As Long
Wth = Form1.P1.Width: Hgt = Form1.P1.Height
BMInfo = CreateBMInfo(Wth, Hgt, BitCount, I)
ReDim arrDIBs(Wth * Hgt * BitCount / 8) As Byte
iDC = CreateCompatibleDC(0)
hDIB = CreateDIBSection(iDC, BMInfo, DIB_RGB_COLORS, I, 0&, 0&)
Call SelectObject(iDC, hDIB)
Call BitBlt(iDC, 0, 0, Wth, Hgt, hSrcDC, 0, 0, vbSrcCopy)
Call GetDIBits(iDC, hDIB, 0, Hgt, arrDIBs(0), BMInfo, 0)
Call DeleteDC(iDC)
Call DeleteObject(hDIB)
'Debug.Print BitCount, Timer - nt
End Sub

Function CreateBMInfo(ByVal Wth As Long, ByVal Hgt As Long, ByVal BitCount As Long, Optional Num As Long) As BitMapInfo256
Dim I As Long
Dim R As Long, G As Long, B As Long
With CreateBMInfo
    With .bmiHeader
        .biSize = Len(CreateBMInfo.bmiHeader)
        .biWidth = Wth
        .biHeight = Hgt
        .biPlanes = 1
        .biBitCount = BitCount '256 É« = 8, 16 É« = 4
        .biCompression = BI_RGB
        .biSizeImage = Wth * Hgt
    End With
    If .bmiHeader.biBitCount = 8 Then '256 É«
        For B = 0 To &HE0 Step &H20
            For G = 0 To &HE0 Step &H20
                For R = 0 To &HC0 Step &H40
                    .bmiColors(I) = IIf(B = &HE0, &HFF, B) * &H10000 + IIf(G = &HE0, &HFF, G) * &H100 + IIf(R = &HC0, &HFF, R)
                    I = I + 1
                Next R
            Next G
        Next B
    ElseIf .bmiHeader.biBitCount = 4 Then '16 É«
        For I = 0 To 15
            .bmiColors(I) = QBColor(I)
        Next I
    End If
End With
If Not IsMissing(Num) Then Num = I
End Function

'the following sub is for compress the images
'I pick it up in the web
Sub Compress(Data() As Byte, suPtr As Long, Optional ByVal I As Long)
Dim e As Long  '¼ÇÂ¼Êý×éÖ¸Õë array pointer
Dim a1 As Long '¼ÇÂ¼Êý×éÖµÖØ¸´µÄ¸öÊý count of repeated arrays
Dim a2 As Long '¼ÇÂ¼µÚÒ»¸öÖØ¸´µÄÐòºÅ first number of repeated one
Dim su() As Byte 'Ñ¹ËõºóµÄÊý×é compressed array
Dim a3 As Long    '¼ÇÂ¼Ñ¹ËõºóµÄÊý×éµÄÖ¸Õë compressed array pointer
If I <= UBound(Data) Then
    'Ñ¹ËõËã·¨ compressing method
    ReDim su(I)
    Do While e < I
        DoEvents
        If (a1 = 255) Then
            su(a3) = Data(e)
            su(a3 + 1) = a1
            a3 = a3 + 2
            a1 = 0
            If (e = I - 1) Then
                su(a3) = Data(I)
                su(a3 + 1) = 0
                ReDim Preserve su(a3 + 5)
            End If
        Else
            If Data(e) = Data(e + 1) Then
                a1 = a1 + 1
                If (e = I - 1) Then
                    su(a3) = Data(e)
                    su(a3 + 1) = a1
                    ReDim Preserve su(a3 + 5)
                End If
            Else
                su(a3) = Data(e)
                su(a3 + 1) = a1
                a3 = a3 + 2
                a1 = 0
                If (e = I - 1) Then
                    su(a3) = Data(I)
                    su(a3 + 1) = 0
                    ReDim Preserve su(a3 + 5)
                End If
            End If
        End If
        e = e + 1
    Loop
    suPtr = a3
    ReDim Data(UBound(su)) As Byte
    CopyMemory Data(0), su(0), UBound(su) + 1
    Exit Sub
Else
    '»¹Ô­Ëã·¨ restore method
    Dim msu() As Byte '»¹Ô­ºóÒª·ÅÈëµÄÊý×é decompressed array
    Dim mi As Long    '¼ÇÂ¼»¹Ô­Ê±¶Á³öµÄÖ¸Õë decompressed pointer
    Dim mx As Long    'ÖØ¸´µÄ¸öÊýµÄµÝÔö repeating count's increasing
    Dim ma As Long    '¼ÇÂ¼Ð´ÈëµÄÖ¸Õë writed pointer
    mi = 0: mx = 0: ma = 0
    ReDim msu(I)
    Do While mi < suPtr
        Do While mx <= Data(mi + 1)
            msu(ma) = Data(mi)
            mx = mx + 1
            ma = ma + 1
        Loop
        mx = 0
        mi = mi + 2
    Loop
    ReDim Data(UBound(msu)) As Byte
    CopyMemory Data(0), msu(0), UBound(msu) + 1
End If
End Sub

