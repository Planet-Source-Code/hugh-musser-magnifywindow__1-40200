<div align="center">

## MagnifyWindow


</div>

### Description

This code magnifies the area under the mouse as you move around the desktop. To use code, open a project with a window.

'Add a timer named Timer1.

'Add a picturebox named Picture1.

'Add a textbox named Text1.

'Add a UpDown control named UpDown1.

'Copy code below to your form.
 
### More Info
 
Uses mouse location to determine area to be magnified.

Uses API calls.

Draws a section of the desktop magnified.

None Known.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Hugh Musser](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/hugh-musser.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/hugh-musser-magnifywindow__1-40200/archive/master.zip)





### Source Code

```
Option Explicit
Private Type POINTAPI
 x As Long
 y As Long
End Type
Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "USER32" (lpPoint As POINTAPI) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetWindowDC Lib "USER32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "USER32" () As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, _
  ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
  ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
  ByVal ySrc As Long, ByVal nSrcWidth As Long, _
  ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Sub Form_Load()
Me.Move 10, 10, 2775, 3390 'position form
UpDown1.Value = 50
Text1.Text = UpDown1.Value & "%"
Me.AutoRedraw = True
Timer1.Interval = 1
End Sub
Private Sub Form_Resize()
Picture1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - 315
Text1.Move 0, Me.ScaleHeight - 315, 765, 315
UpDown1.Move 765, Me.ScaleHeight - 315, 195, 315
End Sub
Private Sub Timer1_Timer()
Dim rv As Long, mXY As POINTAPI, magFCT As Single
Dim hWP As Long, hPP As Long, maxWIDTH As Long, maxHEIGHT As Long
Dim src_LEFT As Long, src_TOP As Long, src_WIDTH As Long, src_HEIGHT As Long
Dim dst_LEFT As Long, dst_TOP As Long, dst_WIDTH As Long, dst_HEIGHT As Long
Dim dst_centerX As Long, dst_centerY As Long
Dim src_HANDLE As Long, src_DC As Long
Dim meW As Long, meH As Long
 magFCT = 1 - (UpDown1.Value / 100) 'magnification factor
 rv = GetCursorPos(mXY) 'get mouse position
 hWP = WindowFromPoint(mXY.x, mXY.y) 'handle to window under mouse
 hPP = GetParent(hWP) 'handle to parent to window under mouse
 If hPP = 0 Then hPP = hWP 'if no parent, use window from point
 If hPP <> Me.hwnd Then 'do not magnify our form
  '--describe area to accept magnification----------------
  dst_centerX = Picture1.ScaleWidth / 2
  dst_centerY = Picture1.ScaleHeight / 2
  dst_LEFT = 0
  dst_TOP = 0
  dst_WIDTH = Picture1.ScaleWidth / Screen.TwipsPerPixelX
  dst_HEIGHT = Picture1.ScaleHeight / Screen.TwipsPerPixelY
  '--describe area of screen to magnify------
  meH = (Picture1.ScaleHeight / Screen.TwipsPerPixelX) * magFCT
  meW = (Picture1.ScaleWidth / Screen.TwipsPerPixelY) * magFCT
  src_LEFT = mXY.x - (meW / 2)
  src_TOP = mXY.y - (meH / 2)
  src_WIDTH = meW
  src_HEIGHT = meH
  '--adjust for edge of screen----------------------------
  maxWIDTH = Screen.Width / Screen.TwipsPerPixelX
  maxHEIGHT = Screen.Height / Screen.TwipsPerPixelY
  If src_LEFT < 0 Then
   dst_centerX = dst_centerX + (src_LEFT * (Screen.TwipsPerPixelX / magFCT))
   src_LEFT = 0
  ElseIf src_LEFT + src_WIDTH > maxWIDTH Then
   dst_centerX = dst_centerX + (src_LEFT + src_WIDTH - maxWIDTH) * (Screen.TwipsPerPixelX / magFCT)
   src_LEFT = src_LEFT - (src_LEFT + src_WIDTH - maxWIDTH)
  End If
  If src_TOP < 0 Then
   dst_centerY = dst_centerY + (src_TOP * (Screen.TwipsPerPixelY / magFCT))
   src_TOP = 0
  ElseIf src_TOP + src_HEIGHT > maxHEIGHT Then
   dst_centerY = dst_centerY + (src_TOP + src_HEIGHT - maxHEIGHT) * (Screen.TwipsPerPixelY / magFCT)
   src_TOP = src_TOP - (src_TOP + src_HEIGHT - maxHEIGHT)
  End If
  src_HANDLE = GetDesktopWindow() 'get a handle to screen
  src_DC = GetWindowDC(src_HANDLE) 'get device context to screen
  '--copy section of screen to form-----------------------------
  StretchBlt Picture1.hdc, _
     dst_LEFT, dst_TOP, dst_WIDTH, dst_HEIGHT, _
     src_DC, _
     src_LEFT, src_TOP, src_WIDTH, src_HEIGHT, vbSrcCopy
  rv = ReleaseDC(src_HANDLE, src_DC) ' release screen dc
  '--draw spotter on screen
  Picture1.Line (dst_centerX, dst_centerY - 300)-(dst_centerX, dst_centerY + 300)
  Picture1.Line (dst_centerX - 300, dst_centerY)-(dst_centerX + 300, dst_centerY)
 End If
 Me.Caption = "(x,y)=" & mXY.x * Screen.TwipsPerPixelX & "," & mXY.y * Screen.TwipsPerPixelY
End Sub
Private Sub UpDown1_Change()
Text1.Text = UpDown1.Value & "%"
End Sub
```

