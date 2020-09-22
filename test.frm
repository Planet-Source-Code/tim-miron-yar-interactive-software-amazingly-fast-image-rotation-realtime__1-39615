VERSION 5.00
Begin VB.Form frm_rotate 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rotate a bitmap"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5370
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "test.frx":0000
   ScaleHeight     =   269
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   358
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Cleanup DIB stuff!"
      Enabled         =   0   'False
      Height          =   600
      Left            =   3900
      TabIndex        =   5
      Top             =   3390
      Width           =   1395
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SetupDIB"
      Height          =   495
      Left            =   3900
      TabIndex        =   4
      Top             =   2265
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DrawDIB rotate"
      Enabled         =   0   'False
      Height          =   600
      Left            =   3900
      TabIndex        =   3
      Top             =   2775
      Width           =   1395
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3195
      Left            =   3435
      Max             =   360
      TabIndex        =   0
      Top             =   750
      Width           =   345
   End
   Begin VB.Label lblFPS 
      BackStyle       =   0  'Transparent
      Caption         =   "FPS"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3915
      TabIndex        =   2
      Top             =   1920
      Width           =   1020
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3870
      TabIndex        =   1
      Top             =   1050
      Width           =   855
   End
End
Attribute VB_Name = "frm_rotate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NOTES:

'KNOWN RESTRICTIONS/ERRORS: The picture must be a square
'meaning that the width and height must be equel.  Also,
'the dimensions of the image and the buffer must be
'multiples of 4, or errors may occur, for best results
'stick to binary numbers [16, 32, 64, 128 etc.].

Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0
Private Const TstDiment = 48
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub Check1_Click()
If Check1.Value = 1 Then
   Picture2.Visible = True
   Else
   Picture2.Visible = False
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim T1 As Long, T2 As Long
'test DIB
T2 = GetTickCount
DIB_rotate SrcBI24, bBytesSRC, UBound(bBytesSRC), _
HdestBufDC, DestBI24, bBytesDest, UBound(bBytesDest), VScroll1.Value
'GERotateSprite SrcBI24, bBytesSRC(1), UBound(bBytesSRC), _
frm_rotate.hdc, DestBI24, bBytesDest(1), UBound(bBytesDest), VScroll1.Value

T1 = GetTickCount
'blit from buffer to form
BitBlt frm_rotate.hdc, 0, 0, TstDiment, TstDiment, HdestBufDC, 0, 0, vbSrcCopy

Label2.Caption = VScroll1.Value & "'"
lblFPS.Caption = Round(1000 / (T1 - T2), 2) & " Fps"
End Sub

Private Sub Command2_Click()
'NOTE: changing the Constant TstDiment will change the
'dimensions of the rotation picture, so if you load
'a larger image onto the form, you should change the
'TstDiment constant to be the width of the image

'source buffer...
      With SrcBI24.bmiHeader
          .biBitCount = 24
          .biCompression = BI_RGB
          .biPlanes = 1
          .biSize = Len(SrcBI24.bmiHeader)
          .biWidth = TstDiment
          .biHeight = TstDiment
      End With
      
      'resize array of bytes to hold the all the pixel's
      'information of the bitmap
     ReDim bBytesSRC(1 To SrcBI24.bmiHeader.biWidth * SrcBI24.bmiHeader.biHeight * 3) As Byte
    HsrcBufDC = CreateCompatibleDC(0)
   HsrcDIBitmap = CreateDIBSection(HsrcBufDC, SrcBI24, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
  SelectObject HsrcBufDC, HsrcDIBitmap
  
  'copy the picture from the form into the buffer...
 BitBlt HsrcBufDC, 0, 0, SrcBI24.bmiHeader.biWidth, SrcBI24.bmiHeader.biHeight, frm_rotate.hdc, 0, 0, vbSrcCopy
 
 'read the picture information into the array of bytes
 'so we can pass it to the rotation function.
GetDIBits HsrcBufDC, HsrcDIBitmap, 0, SrcBI24.bmiHeader.biHeight, bBytesSRC(1), SrcBI24, DIB_RGB_COLORS

'clear the form
frm_rotate.Cls
Set frm_rotate.Picture = Nothing


'destination buffer...
      With DestBI24.bmiHeader
          .biBitCount = 24
          .biCompression = BI_RGB
          .biPlanes = 1
          .biSize = Len(DestBI24.bmiHeader)
          .biWidth = TstDiment
          .biHeight = TstDiment
      End With
      'resize array of bytes to hold the entire bitmap
     ReDim bBytesDest(1 To DestBI24.bmiHeader.biWidth * DestBI24.bmiHeader.biHeight * 3) As Byte
    HdestBufDC = CreateCompatibleDC(0)
   HdestDIBitmap = CreateDIBSection(HdestBufDC, DestBI24, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
  SelectObject HdestBufDC, HdestDIBitmap
GetDIBits HdestBufDC, HdestDIBitmap, 0, DestBI24.bmiHeader.biHeight, bBytesDest(1), DestBI24, DIB_RGB_COLORS


Command1.Enabled = True
Command3.Enabled = True
Command2.Enabled = False
MsgBox "IF I'M NOT COMPILED I'M ALOT SLOWER!"
MsgBox "Drag the scrollbar to rotate! :)"
End Sub

Private Sub Command3_Click()
'delete device contexts
DeleteDC HsrcBufDC
DeleteDC HdestBufDC

'delete bitmap objects
DeleteObject HsrcDIBitmap
DeleteObject HdestDIBitmap

MsgBox "All cleaned up! ;)"
Command1.Enabled = False
Command3.Enabled = False
Command2.Enabled = True

'unload form and erase from memory
Unload Me
Set frm_rotate = Nothing
End
End Sub

'when the scrollbar moves or changes values,
'do the rotation
Private Sub VScroll1_Change()
On Error Resume Next
Command1_Click
End Sub

Private Sub VScroll1_Scroll()
Command1_Click
End Sub
