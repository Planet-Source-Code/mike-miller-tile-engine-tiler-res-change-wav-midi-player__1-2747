<div align="center">

## Tile Engine \(tiler,res change,wav & midi player\)


</div>

### Description

This will make it easy for someone to make a cool tile game. The coding on their part will take basicly nothing and produce a quality game. It also has a wav player, midi player and some other stuff. Check it out! The reason I made this is because every place i went to look for a tile engine, either didn't have one or the code was all in the form and was really jacked up. With the engine that I made(rattyrat13@aol.com) it is all in a moudle and very easy to understand. It currently supports up to 35 diferent tiles but that can be changed to make it more. MAKE SURE THAT AUTO-REDRAW IS ON! If autoredraw isn't true then you will have to make sure that all the picture boxes that are being used as the input are still visable to the user.

'Newly updated 

----

I forgot to add about a transparent bitmap in here, so people who want just that, just steal that and take it. Also I want everybody to know that this engine is fast, because it uses BitBlit ( Bit Blit or Bitmap Blaster) not paintpicture. There is no reason that I can see to use direct-x for a tile engine, BitBlit is fast enough.
 
### More Info
 
The map file, if it isn't self explanitory, email me. And the PictureBoxes. Other than that there is some stuff if you want to use the non-engine part of the moudle. The File Name For the WAV, Or MIDI, and the requried inputs for the transparent bliter.

TURN AUTOREDRAW ON!!!!!!

If something is wrong, just try refreshing the form, in a timer with intraval of about 500 do

me.refresh

timer1.enabled = false

end sub

This will make everything better if it doesn't work.

A BitBlited Form of a large picture produced by many tiles, or even just 1 tile, it doesn't really matter.

Beware of an extreamly cool game made by you with this engine. None Other than that though, if you find one, e-mail it to me and I will correct it and send you the corrected code along with everyone else who reads it, and you shall get some credit for helping me.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mike Miller](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mike-miller.md)
**Level**          |Unknown
**User Rating**    |6.0 (605 globes from 101 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Games](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/games__1-38.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mike-miller-tile-engine-tiler-res-change-wav-midi-player__1-2747/archive/master.zip)

### API Declarations

```
'''''''''''''''''''''''TM
'''Funky Tile Engine'''
'''Mike Miller '''
'''1999  '''
'''''''''''''''''''''''
'autoredraw must be true!
'RattyRat13@aol.com
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpSound As String, ByVal flag As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long
Const CCDEVICENAME = 32
Const CCFORMNAME = 32
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000
Public Type RECT
 Left As Long
 Top As Long
 Right As Long
 Bottom As Long
End Type
Private Type DEVMODE
 dmDeviceName As String * CCDEVICENAME
 dmSpecVersion As Integer
 dmDriverVersion As Integer
 dmSize As Integer
 dmDriverExtra As Integer
 dmFields As Long
 dmOrientation As Integer
 dmPaperSize As Integer
 dmPaperLength As Integer
 dmPaperWidth As Integer
 dmScale As Integer
 dmCopies As Integer
 dmDefaultSource As Integer
 dmPrintQuality As Integer
 dmColor As Integer
 dmDuplex As Integer
 dmYResolution As Integer
 dmTTOption As Integer
 dmCollate As Integer
 dmFormName As String * CCFORMNAME
 dmUnusedPadding As Integer
 dmBitsPerPel As Integer
 dmPelsWidth As Long
 dmPelsHeight As Long
 dmDisplayFlags As Long
 dmDisplayFrequency As Long
End Type
 Dim DevM As DEVMODE
```


### Source Code

```
Sub ChangeRes(iWidth As Single, iHeight As Single)
'Just Call Changeres(1600,1200) or whatever you want in load
 Dim a As Boolean
 Dim i&
 i = 0
 Do
 a = EnumDisplaySettings(0&, i&, DevM)
 i = i + 1
 Loop Until (a = False)
 Dim b&
 DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
 DevM.dmPelsWidth = iWidth
 DevM.dmPelsHeight = iHeight
 b = ChangeDisplaySettings(DevM, 0)
End Sub
Public Sub TilePicture(frmDest As Form, source As PictureBox, X, Y)
'This is not the sub that you want to use, it may be a good one to modify though
'If you think you need Direct-X or just want to see what will work.
 Dim pw As Integer
 Dim ph As Integer
 Dim fw As Integer
 Dim fh As Integer
 Dim rst As Integer
 source.ScaleMode = 3
 pw = source.ScaleWidth
 ph = source.ScaleHeight
 fw = frmDest.Width / Screen.TwipsPerPixelX
 fh = frmDest.Height / Screen.TwipsPerPixelY
iResult = BitBlt(frmDest.hdc, X, Y, iPicWidth, iPicHeight, picSource.hdc, 0, 0, vbSrcCopy)
End Sub
Public Sub LoadMap(InvisText As TextBox, mapname As String)
'maps constists of numbers like
'
'00012301
'12321455
'51000102
'and so forth, if anyone wants to make a map editor, that would be cool,
'but I don't got the time(5:00 - 9:00pm) in football practice
 Dim lFileLength As Long
 Dim iFileNum As Integer
 iFileNum = FreeFile
 Open mapname For Input As iFileNum
 lFileLength = LOF(iFileNum)
 Text1.Text = Input(lFileLength, #iFileNum)
 Close iFileNum
End Sub
Public Sub OpenMidi()
'dont call this, it needs a few mods
 Dim sFile As String
 Dim sShortFile As String * 67
 Dim lResult As Long
 Dim sError As String * 255
 sFile = App.Path & "\midtest.mid"
 lResult = GetShortPathName(sFile, sShortFile, Len(sShortFile))
 sFile = Left$(sShortFile, lResult)
 lResult = mciSendString("open " & sFile & _
 " type sequencer alias mcitest", ByVal 0&, 0, 0)
 If lResult Then
 lResult = mciGetErrorString(lResult, sError, 255)
 Debug.Print "open: " & sError
 End If
End Sub
Public Sub PlayMidi()
'see above
 Dim lResult As Integer
 Dim sError As String * 255
 lResult = mciSendString("play mcitest", ByVal 0&, 0, 0)
 If lResult Then
 lResult = mciGetErrorString(lResult, sError, 255)
 Debug.Print "play: " & sError
 End If
End Sub
Public Sub CloseMidi()
'again see above, i am sorry I will update soon
 Dim lResult As Integer
 Dim sError As String * 255
 lResult = mciSendString("close mcitest", "", 0&, 0&)
 If lResult Then
 lResult = mciGetErrorString(lResult, sError, 255)
 Debug.Print "stop: " & sError
 End If
End Sub
Sub PlayWave(sFileName As String)
 On Error GoTo Play_Err
 Dim iReturn As Integer
 If sFileName > "" Then
 If UCase$(Right$(sFileName, 3)) = "WAV" Then
  If Dir(sFileName) > "" Then
  iReturn = sndPlaySound(sFileName, 0)
  End If
 End If
 End If
 Exit Sub
Play_Err:
 Exit Sub
End Sub
Function TileWalkable(Tilesize As Integer, LoadedMap As TextBox, X As Integer, Y As Integer, LineWidth As Integer) As Boolean
'Funky Tile Engine Note:
'Most pic boxes use twip, so divide pic.width by screen.twipsperpixelx and same for height, execpt for y insted.
'I also suggest that you modify this if you are tring to make a more customized
'engine, because this at this time gives you 18 unwalkables
Dim xx As Integer
Dim yy As Integer
Dim temp As Integer
Dim a As String
Dim b As String
xx = X / Tilesize
yy = Y / Tilesize
If Y < Tilesize Then
 a = Left(LoadedMap, xx)
 b = Mid(a, xx, 1): GoTo 1
End If
temp = yy * LineWidth + 2
a = Left(LoadedMap, xx + temp)
b = Mid(a, xx + temp, 1): GoTo 1
1
MsgBox b
If b = "0" Then TileWalkable = False: Exit Function
If b = "1" Then TileWalkable = False: Exit Function
If b = "2" Then TileWalkable = False: Exit Function
If b = "3" Then TileWalkable = False: Exit Function
If b = "4" Then TileWalkable = False: Exit Function
If b = "5" Then TileWalkable = False: Exit Function
If b = "6" Then TileWalkable = False: Exit Function
If b = "7" Then TileWalkable = False: Exit Function
If b = "8" Then TileWalkable = False: Exit Function
If b = "9" Then TileWalkable = False: Exit Function
If b = "a" Then TileWalkable = False: Exit Function
If b = "b" Then TileWalkable = False: Exit Function
If b = "c" Then TileWalkable = False: Exit Function
If b = "d" Then TileWalkable = False: Exit Function
If b = "e" Then TileWalkable = False: Exit Function
If b = "f" Then TileWalkable = False: Exit Function
If b = "g" Then TileWalkable = False: Exit Function
TileWalkable = True
End Function
Sub Tilemake(LoadedMap As TextBox, MapXLength As Integer, MapYLength, PicWidth As Integer, Dest As Form, Optional pic0 As PictureBox, Optional pic1 As PictureBox, Optional pic2 As PictureBox, Optional pic3 As PictureBox, Optional pic4 As PictureBox, Optional pic5 As PictureBox, Optional pic6 As PictureBox, Optional pic7 As PictureBox, Optional pic8 As PictureBox, Optional pic9 As PictureBox, Optional pic10 As PictureBox, Optional pic11 As PictureBox, Optional pic12 As PictureBox, Optional pic13 As PictureBox, Optional pic14 As PictureBox, Optional pic15 As PictureBox, Optional pic16 As PictureBox, Optional pic17 As PictureBox, Optional pic18 As PictureBox, Optional pic19 As PictureBox, Optional pic20 As PictureBox, Optional pic21 As PictureBox, Optional pic22 As PictureBox, Optional pic23 As PictureBox, Optional pic24 As PictureBox, Optional pic25 As PictureBox, Optional pic26 As PictureBox, Optional pic27 As PictureBox, Optional pic28 As PictureBox, Optional pic29 As PictureBox, _
Optional pic30 As PictureBox, Optional pic31 As PictureBox, Optional pic32 As PictureBox, Optional pic33, Optional pic34 As PictureBox, Optional pic35 As PictureBox)
'this is what you call
'all pictureboxes are optional, so you don't have to use them all
'Put me in the form paint
'after 0123456789 comes a - z
'be creative if you want more, ~!@#$%^&*()_+
cc = 0
aa = 0
bb = 0
1
For i = 0 To MapXLength
a = Mid(LoadedMap, i + aa + 1, 1)
dd = i * PicWidth
dd = dd + 224
If a = "0" Then Call TilePicture(Dest, pic0, dd, cc)
If a = "1" Then Call TilePicture(Dest, pic1, dd, cc)
If a = "2" Then Call TilePicture(Dest, pic2, dd, cc)
If a = "3" Then Call TilePicture(Dest, pic3, dd, cc)
If a = "4" Then Call TilePicture(Dest, pic4, dd, cc)
If a = "5" Then Call TilePicture(Dest, pic5, dd, cc)
If a = "6" Then Call TilePicture(Dest, pic6, dd, cc)
If a = "7" Then Call TilePicture(Dest, pic7, dd, cc)
If a = "8" Then Call TilePicture(Dest, pic8, dd, cc)
If a = "9" Then Call TilePicture(Dest, pic9, dd, cc)
If a = "a" Then Call TilePicture(Dest, pic10, dd, cc)
If a = "b" Then Call TilePicture(Dest, pic11, dd, cc)
If a = "c" Then Call TilePicture(Dest, pic12, dd, cc)
If a = "d" Then Call TilePicture(Dest, pic13, dd, cc)
If a = "e" Then Call TilePicture(Dest, pic14, dd, cc)
If a = "f" Then Call TilePicture(Dest, pic15, dd, cc)
If a = "g" Then Call TilePicture(Dest, pic16, dd, cc)
If a = "h" Then Call TilePicture(Dest, pic17, dd, cc)
If a = "i" Then Call TilePicture(Dest, pic18, dd, cc)
If a = "j" Then Call TilePicture(Dest, pic19, dd, cc)
If a = "k" Then Call TilePicture(Dest, pic20, dd, cc)
If a = "l" Then Call TilePicture(Dest, pic21, dd, cc)
If a = "m" Then Call TilePicture(Dest, pic22, dd, cc)
If a = "n" Then Call TilePicture(Dest, pic23, dd, cc)
If a = "o" Then Call TilePicture(Dest, pic24, dd, cc)
If a = "p" Then Call TilePicture(Dest, pic25, dd, cc)
If a = "q" Then Call TilePicture(Dest, pic26, dd, cc)
If a = "r" Then Call TilePicture(Dest, pic27, dd, cc)
If a = "s" Then Call TilePicture(Dest, pic28, dd, cc)
If a = "t" Then Call TilePicture(Dest, pic29, dd, cc)
If a = "u" Then Call TilePicture(Dest, pic30, dd, cc)
If a = "v" Then Call TilePicture(Dest, pic31, dd, cc)
If a = "w" Then Call TilePicture(Dest, pic32, dd, cc)
'If a = "x" Then Call TilePicture(Dest, pic33, dd, cc)
'If a = "y" Then Call TilePicture(Dest, pic34, dd, cc)
'If a = "z" Then Call TilePicture(Dest, pic35, dd, cc)
Next i
cc = cc + PicWidth
aa = aa + MapXLength + 2
bb = bb + 1
If bb > MapYLength Then Exit Sub
GoTo 1
End Sub
'Private Sub TransparentBlt(OutDstDC As Long, DstDC As Long, SrcDC As Long, SrcRect As RECT, DstX As Integer, DstY As Integer, TransColor As Long)
' Dim nRet As Long, W As Integer, H As Integer
' Dim MonoMaskDC As Long, hMonoMask As Long
' Dim MonoInvDC As Long, hMonoInv As Long
' Dim ResultDstDC As Long, hResultDst As Long
' Dim ResultSrcDC As Long, hResultSrc As Long
' Dim hPrevMask As Long, hPrevInv As Long
' Dim hPrevSrc As Long, hPrevDst As Long
' W = SrcRect.Right - SrcRect.Left + 1
' H = SrcRect.Bottom - SrcRect.Top + 1
' MonoMaskDC = CreateCompatibleDC(DstDC)
' MonoInvDC = CreateCompatibleDC(DstDC)
' hMonoMask = CreateBitmap(W, H, 1, 1, ByVal 0&)
' hMonoInv = CreateBitmap(W, H, 1, 1, ByVal 0&)
' hPrevMask = SelectObject(MonoMaskDC, hMonoMask)
' hPrevInv = SelectObject(MonoInvDC, hMonoInv)
' ResultDstDC = CreateCompatibleDC(DstDC)
' ResultSrcDC = CreateCompatibleDC(DstDC)
' hResultDst = CreateCompatibleBitmap(DstDC, W, H)
' hResultSrc = CreateCompatibleBitmap(DstDC, W, H)
' hPrevDst = SelectObject(ResultDstDC, hResultDst)
' hPrevSrc = SelectObject(ResultSrcDC, hResultSrc)
' Dim OldBC As Long
' OldBC = SetBkColor(SrcDC, TransColor)
' nRet = BitBlt(MonoMaskDC, 0, 0, W, H, SrcDC, SrcRect.Left, SrcRect.Top, vbSrcCopy)
' TransColor = SetBkColor(SrcDC, OldBC)
' nRet = BitBlt(MonoInvDC, 0, 0, W, H, MonoMaskDC, 0, 0, vbNotSrcCopy)
' nRet = BitBlt(ResultDstDC, 0, 0, W, H, DstDC, DstX, DstY, vbSrcCopy)
' nRet = BitBlt(ResultDstDC, 0, 0, W, H, MonoMaskDC, 0, 0, vbSrcAnd)
' nRet = BitBlt(ResultSrcDC, 0, 0, W, H, SrcDC, SrcRect.Left, SrcRect.Top, vbSrcCopy)
' nRet = BitBlt(ResultSrcDC, 0, 0, W, H, MonoInvDC, 0, 0, vbSrcAnd)
' nRet = BitBlt(ResultDstDC, 0, 0, W, H, ResultSrcDC, 0, 0, vbSrcInvert)
' nRet = BitBlt(OutDstDC, DstX, DstY, W, H, ResultDstDC, 0, 0, vbSrcCopy)
' hMonoMask = SelectObject(MonoMaskDC, hPrevMask)
' DeleteObject hMonoMask
' hMonoInv = SelectObject(MonoInvDC, hPrevInv)
' DeleteObject hMonoInv
' hResultDst = SelectObject(ResultDstDC, hPrevDst)
' DeleteObject hResultDst
' hResultSrc = SelectObject(ResultSrcDC, hPrevSrc)
' DeleteObject hResultSrc
' DeleteDC MonoMaskDC
' DeleteDC MonoInvDC
' DeleteDC ResultDstDC
' DeleteDC ResultSrcDC
'End Sub
'Dim R As RECT
' With R
' .Left = 0
' .Top = 0
' .Right = Picture1.ScaleWidth
' .Bottom = Picture1.ScaleHeight
'End With
'
'TransparentBlt Form1.hDC, Form1.hDC, Picture1.hDC, R, 20, 20, vbblack
```

