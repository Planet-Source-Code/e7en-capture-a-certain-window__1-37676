<div align="center">

## Capture a Certain Window


</div>

### Description

This will capture the specified window and paste it into a picturebox. Really Simple and easy to follow. Please Vote and Comment!
 
### More Info
 
Just add a Picturebox to a form and use the line

CaptureWindow hWnd, PictureBox

eg. CaptureWindow Me.hWnd, Picture1


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[ï¿½e7eN](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/e7en.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/e7en-capture-a-certain-window__1-37676/archive/master.zip)





### Source Code

```
Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
  End Type
Sub CaptureWindow(WindowhWnd As Long, Output As PictureBox)
    Dim Ret As Long
    Dim WindowRect As RECT
    Dim WindowhWnd As Long
    Dim nHeight As Long, nWidth As Long
  Output.Cls 'Clear the picturebox
  Ret = GetWindowRect(WindowhWnd, WindowRect) 'Get the windows co-ordinates
  nWidth = WindowRect.Right - WindowRect.Left 'Get the windows Width
  nHeight = WindowRect.Bottom - WindowRect.Top ' Get the windows height
  Ret = BitBlt(Output.hDC, 0, 0, nWidth, nHeight, GetWindowDC(WindowhWnd), 0, 0, vbSrcCopy)'Get the windows image and copy it to the Picturebox
End Sub
```

