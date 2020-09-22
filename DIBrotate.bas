Attribute VB_Name = "DIBrotate"
Private Const Pi As Single = 3.1415926
Private Const Trans As Single = (Pi / 180) 'used to be a double
Private Const Pd2 As Single = (Pi / 2)     'used to be a double

Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs
Type BITMAPINFOHEADER '40 bytes
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
Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type
Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

'=== example only ====
'source DIB buffer stuff
Public HsrcBufDC As Long, HsrcDIBitmap As Long
Public bBytesSRC() As Byte
Public SrcBI24 As BITMAPINFO

'destination DIB buffer stuff
Public HdestBufDC As Long, HdestDIBitmap As Long
Public bBytesDest() As Byte
Public DestBI24 As BITMAPINFO


'Dim iBitmap As Long, iDC As Long
'Dim bitmap As BITMAPINFO, bBytes() As Byte, Cnt As Long
'GERotateSprite (BITMAPINFO* SRCbi24BitInfo, char* bSrcBytes[], long SRCbyteCount, HDC hDestDC, BITMAPINFO* DESTbi24BitInfo, char* bDestBytes[], long DESTbyteCount, long Angle) {
Public Sub DIB_rotate(SRCbi24BitInfo As BITMAPINFO, bSrcBytes() As Byte, SRCbyteCount As Long, _
           hDestDC As Long, _
           DESTbi24BitInfo As BITMAPINFO, _
           bDestBytes() As Byte, _
           DESTbyteCount As Long, _
           angle As Long) '{
'//void __stdcall GERotateSprite (BITMAPINFO SRCbi24BitInfo, char bSrcBytes[], int SRCbyteCount,
'// HDC hDestDC, BITMAPINFO DESTbi24BitInfo, char bDestBytes[],
'// int DESTbyteCount, int Angle) {
'Things that should be done by the owner program:
'EG.
'dim bitmap as BITMAPINFO
'// gets passed to this function as SRCbi24BitInfo

'    With bitmap.bmiHeader
'        .biBitCount = 24
'        .biCompression = BI_RGB
'        .biPlanes = 1
'        .biSize = Len(bitmap.bmiHeader)
'        .biWidth = 100 'width of bitmap in pixels
'        .biHeight = 100 'height of bitmap in pixels
'    End With
'
'       iDC = CreateCompatibleDC(0)
'        iBitmap = CreateDIBSection(iDC, bitmap, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
'        SelectObject iDC, iBitmap
'        ^^^
'//IDC and iBitmap are not needed by this function
'// because it bypasses that layer of the GDI and
'// only deals with the bitarray, HOWEVER - IMPORTANT:
'// that last step of creating the diBitmap object
'// and selecting it into the source DC is VERY important!
'// or else it wont work: NOTE: use DeleteObject and
'// deleteDC to delete the source & destination diBitmap
'// objects aswell as the DC...
'
'
''retrieve source bits
'GetDIBits iDC, iBitmap, 0, bitmap.bmiHeader.biHeight, bBytes(1), bitmap, DIB_RGB_COLORS
'//bBytes() is passed to us as BsrcBYTES()
'
'//do the same thing for the destination buffer
'/* note for C++ porting: same variable declaration as
'the first rotation sub, but plus the definition of bbytes!
    Dim c1x As Long, c1y As Long
    Dim c2x As Long, c2y As Long
    Dim a As Single
    Dim p1x As Long, p1y As Long
    Dim p2x As Long, p2y As Long
    Dim n As Long, R As Double
    Dim theta As Single 'this is angle * trans
        Dim ApT2 As Single 'A + theta (value=0 at declaration)

    Dim PCoord As Long, PCoord2 As Long
    
    'Dim Result As String 'for testing only
    'Dim TR As Boolean, BL As Boolean
    Dim D As Long
'// MUCH OF THE BELOW VARIABLE SETUP
'CAN BE THE EXACT SAME AS IN THE EARLIER
'C++ VERSION OF THIS FUNCTION...
'IMPORTANT NOTICE, THE COLORS ARE NOT COLORREFS ANYMORE!
'DOWN TO THE SETUP OF N
'long theta=angle * trans;
'but in vb we do it like this:
    theta = (angle * Trans)
    
    '!= is the same as <>
    
    '//these can also be calculated when declared
    '(center of destination)
    c1x = (SRCbi24BitInfo.bmiHeader.biWidth \ 2)
    c1y = (SRCbi24BitInfo.bmiHeader.biHeight \ 2)
    c2x = (DESTbi24BitInfo.bmiHeader.biWidth \ 2)
    c2y = (DESTbi24BitInfo.bmiHeader.biHeight \ 2)
    
    n = (c2y - 1)
   
       ' n = (c1x - 1) / 2
    '#####
    
    '//n is used to determine when the for loops
    '//including the nested loop should end...
    '//THE END OF THE SETUP OF VARIABLES

    For p2x = 0 To n
        For p2y = 0 To n
        If p2x = 0 Then a = Pd2 Else a = Atn(p2y / p2x) 'in c++ this = atan
           R = Sqr((p2x * p2x) + (p2y * p2y))
           '//minor optimization (avoid doing the same calculation twice)
           '// NOTE: this value can NOT be set when its declared
           '// because the value of A changes in every cycle of this loop
            ApT = (a + theta)
                p1x = R * Cos(ApT) 'used to be (a + theta) written twice
                p1y = R * Sin(ApT)
            '//retrieve pixel color
            'c0& = GetPixel(HsrcDC, c1x + p1x, c1y + p1y)
            
    'BOTTOM-RIGHT
   PCoord = (((SRCbi24BitInfo.bmiHeader.biHeight - (c1y + p1y)) * SRCbi24BitInfo.bmiHeader.biHeight + (c1x + p1x)) * 3) - 2
    If (PCoord + 1) < SRCbyteCount And PCoord > 0 Then
       
       'by eliminating these in-between
       'variables C0[r][g][b]
       'we eliminate 3 copy operators
       'per pixel, AND 9 variable declarations,
       'AND we avoid processing any pixels outside
       'of the rotation's 'edges'..
       '(however, that may cause some problems
       'and actually make thigns take longer
       'if the program has to use a bigger
       'sprite just to make sure there's enough
       'padding to fill the edges properly, so I
       'am questioning whether to keep this or not?)
       
       PCoord2 = (((DESTbi24BitInfo.bmiHeader.biHeight - (c2y + p2y + 1)) * DESTbi24BitInfo.bmiHeader.biHeight + (c2x + p2x)) * 3) + 1
       bDestBytes(PCoord2) = bSrcBytes(PCoord)     'blue
       bDestBytes(PCoord2 + 1) = bSrcBytes(PCoord + 1) 'green
       bDestBytes(PCoord2 + 2) = bSrcBytes(PCoord + 2) 'red
    End If
         
      
    '//top left
    PCoord = (((SRCbi24BitInfo.bmiHeader.biHeight - (c1y - p1y)) * SRCbi24BitInfo.bmiHeader.biHeight + (c1x - p1x)) * 3) - 2
    If (PCoord + 1) < SRCbyteCount And PCoord > 0 Then
        '//TOP-LEFT
         PCoord2 = (((DESTbi24BitInfo.bmiHeader.biHeight - (c2y - p2y)) * DESTbi24BitInfo.bmiHeader.biHeight + (c2x - p2x)) * 3) - 2
         bDestBytes(PCoord2) = bSrcBytes(PCoord)            '//blue
         bDestBytes(PCoord2 + 1) = bSrcBytes(PCoord + 1)    '//green
         bDestBytes(PCoord2 + 2) = bSrcBytes(PCoord + 2)    '//red
    End If
         
         
    '//top-right
    PCoord = (((SRCbi24BitInfo.bmiHeader.biHeight - (c1y - p1x)) * SRCbi24BitInfo.bmiHeader.biHeight + (c1x + p1y)) * 3) - 2
    If (PCoord + 1) < SRCbyteCount And PCoord > 0 Then
        'TOP RIGHT
         PCoord2 = (((DESTbi24BitInfo.bmiHeader.biHeight - (c2y - p2x)) * DESTbi24BitInfo.bmiHeader.biHeight + (c2x + p2y)) * 3) + 1
         bDestBytes(PCoord2) = bSrcBytes(PCoord)     'blue
         bDestBytes(PCoord2 + 1) = bSrcBytes(PCoord + 1) 'green
         bDestBytes(PCoord2 + 2) = bSrcBytes(PCoord + 2) 'red
'tripple comment is when the ifs were seperate
   End If
    
    'bottom-left!
    PCoord = (((SRCbi24BitInfo.bmiHeader.biHeight - (c1y + p1x)) * SRCbi24BitInfo.bmiHeader.biHeight + (c1x - p1y)) * 3) - 2
   If (PCoord + 1) < SRCbyteCount And PCoord > 0 Then
        'BOTTOM-LEFT
         PCoord2 = (((DESTbi24BitInfo.bmiHeader.biHeight - (c2y + p2x + 1)) * DESTbi24BitInfo.bmiHeader.biHeight + (c2x - p2y)) * 3) - 2
         bDestBytes(PCoord2) = bSrcBytes(PCoord)     'blue
         bDestBytes(PCoord2 + 1) = bSrcBytes(PCoord + 1) 'green
         bDestBytes(PCoord2 + 2) = bSrcBytes(PCoord + 2) 'red
    End If
  Next
Next
    'changed index number of bits from 1 to 0
SetDIBitsToDevice hDestDC, 0, 0, SRCbi24BitInfo.bmiHeader.biWidth, SRCbi24BitInfo.bmiHeader.biHeight, 0, 0, 0, SRCbi24BitInfo.bmiHeader.biHeight, bDestBytes(1), SRCbi24BitInfo, DIB_RGB_COLORS


'testing...
'Result = "TR = " & TR & ", BL = " & BL & ": D = " & D
'Clipboard.Clear
'Clipboard.SetText Result
End Sub
