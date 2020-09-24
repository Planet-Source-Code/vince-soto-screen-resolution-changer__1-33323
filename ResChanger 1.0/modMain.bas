Attribute VB_Name = "modMain"
Option Explicit

'Public application variables
Public DM1() As DEVMODE
Public Res(50)
Public oRES As String
Public NumModes As Integer
Public CurrentRes As String
'API function declarations
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpInitData As DEVMODE, ByVal dwFlags As Long) As Long
Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (lpszDeviceName As Any, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
'Public Constants:
Public Const HORZRES = 8
Public Const VERTRES = 10
Public Const BITSPIXEL = 12
Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32
Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const SRCPAINT = &HEE0086
Public Const NOTSRCCOPY = &H330008
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const DM_DISPLAYFLAGS = &H200000
Public Const DM_DISPLAYFREQUENCY = &H400000
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1
Public Const DISP_CHANGE_FAILED = -1
Public Const DISP_CHANGE_BADMODE = -2
Public Const DISP_CHANGE_NOTUPDATED = -3
Public Const DISP_CHANGE_BADFLAGS = -4
Public Const DISP_CHANGE_BADPARAM = -5
Public Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4
Public Const CDS_UPDATEREGISTRY = &H1
'Public types:
'used for display mode settings:
Public Type DEVMODE
   dmDeviceName As String * CCHDEVICENAME
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
   dmFormName As String * CCHFORMNAME
   dmUnusedPadding As Integer
   dmBitsPerPel As Long
   dmPelsWidth As Long
   dmPelsHeight As Long
   dmDisplayFlags As Long
   dmDisplayFrequency As Long
End Type
'used for transparentblt function:
Public Type BITMAP
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type

Sub TransparentBlt(dest As Control, ByVal srcBmp As Integer, ByVal destX As Integer, ByVal destY As Integer, ByVal TransColor As Long)
   'this subroutine replaces transparentblt in MSimg32.dll, which has a memory leak
   'per microsoft article: Q94961
   Const PIXEL = 3
   'variable declarations:
   Dim destScale As Integer
   Dim srcDC As Integer
   Dim saveDC As Integer
   Dim maskDC As Integer
   Dim invDC As Integer
   Dim resultDC As Integer
   Dim bmp As BITMAP
   Dim hResultBmp As Integer
   Dim hSaveBmp As Integer
   Dim hMaskBmp As Integer
   Dim hInvBmp As Integer
   Dim hPrevBmp As Integer
   Dim hSrcPrevBmp As Integer
   Dim hSavePrevBmp As Integer
   Dim hDestPrevBmp As Integer
   Dim hMaskPrevBmp As Integer
   Dim hInvPrevBmp As Integer
   Dim OrigColor As Long
   Dim Success As Integer
   'need to verifiy if form also:
   If TypeOf dest Is PictureBox Then
      destScale = dest.ScaleMode
      dest.ScaleMode = PIXEL
      Success = GetObj(srcBmp, Len(bmp), bmp)
      'create memory device contexts:
      srcDC = CreateCompatibleDC(dest.hDC)
      saveDC = CreateCompatibleDC(dest.hDC)
      maskDC = CreateCompatibleDC(dest.hDC)
      invDC = CreateCompatibleDC(dest.hDC)
      resultDC = CreateCompatibleDC(dest.hDC)
      'create compatible bitmap in DCs:
      hMaskBmp = CreateBitmap(bmp.bmWidth, bmp.bmHeight, 1, 1, ByVal 0&)
      hInvBmp = CreateBitmap(bmp.bmWidth, bmp.bmHeight, 1, 1, ByVal 0&)
      hResultBmp = CreateCompatibleBitmap(dest.hDC, bmp.bmWidth, bmp.bmHeight)
      hSaveBmp = CreateCompatibleBitmap(dest.hDC, bmp.bmWidth, bmp.bmHeight)
      'set bitmaptypes
      hSrcPrevBmp = SelectObject(srcDC, srcBmp)
      hSavePrevBmp = SelectObject(saveDC, hSaveBmp)
      hMaskPrevBmp = SelectObject(maskDC, hMaskBmp)
      hInvPrevBmp = SelectObject(invDC, hInvBmp)
      hDestPrevBmp = SelectObject(resultDC, hResultBmp)
      'do masking work
      Success = BitBlt(saveDC, 0, 0, bmp.bmWidth, bmp.bmHeight, srcDC, 0, 0, SRCCOPY)
      OrigColor = SetBkColor(srcDC, TransColor)
      Success = BitBlt(maskDC, 0, 0, bmp.bmWidth, bmp.bmHeight, srcDC, 0, 0, SRCCOPY)
      TransColor = SetBkColor(srcDC, OrigColor)
      Success = BitBlt(invDC, 0, 0, bmp.bmWidth, bmp.bmHeight, maskDC, 0, 0, NOTSRCCOPY)
      Success = BitBlt(resultDC, 0, 0, bmp.bmWidth, bmp.bmHeight, dest.hDC, destX, destY, SRCCOPY)
      Success = BitBlt(resultDC, 0, 0, bmp.bmWidth, bmp.bmHeight, maskDC, 0, 0, SRCAND)
      Success = BitBlt(srcDC, 0, 0, bmp.bmWidth, bmp.bmHeight, invDC, 0, 0, SRCAND)
      Success = BitBlt(resultDC, 0, 0, bmp.bmWidth, bmp.bmHeight, srcDC, 0, 0, SRCPAINT)
      Success = BitBlt(dest.hDC, destX, destY, bmp.bmWidth, bmp.bmHeight, resultDC, 0, 0, SRCCOPY)
      Success = BitBlt(srcDC, 0, 0, bmp.bmWidth, bmp.bmHeight, saveDC, 0, 0, SRCCOPY)
      'reselects bitmap DCs for removal:
      hPrevBmp = SelectObject(srcDC, hSrcPrevBmp)
      hPrevBmp = SelectObject(saveDC, hSavePrevBmp)
      hPrevBmp = SelectObject(resultDC, hDestPrevBmp)
      hPrevBmp = SelectObject(maskDC, hMaskPrevBmp)
      hPrevBmp = SelectObject(invDC, hInvPrevBmp)
      'deletes bitmap objects:
      Success = DeleteObject(hSaveBmp)
      Success = DeleteObject(hMaskBmp)
      Success = DeleteObject(hInvBmp)
      Success = DeleteObject(hResultBmp)
      'frees memory device context:
      Success = DeleteDC(srcDC)
      Success = DeleteDC(saveDC)
      Success = DeleteDC(invDC)
      Success = DeleteDC(maskDC)
      Success = DeleteDC(resultDC)
      'reset scalemode (if necessary)
      dest.ScaleMode = destScale
   End If
End Sub

Public Sub ResChange(strRes As String)
   'variable declaration:
   Dim ValidRes As Long, UpdtReg As Long, FoundRes As Long, CheckRes As Integer
   'check to see if resolution to be changed is in table Res():
   For CheckRes = 0 To NumModes
      If Res(CheckRes) = strRes Then FoundRes = CheckRes
   Next CheckRes
   'updates system registry to new resolution
   UpdtReg = CDS_UPDATEREGISTRY
   'change display settings:
   ValidRes = ChangeDisplaySettings(DM1(FoundRes), UpdtReg)
   'error handling / if reboot is necessary, prompt to restart computer:
   Select Case ValidRes
      Case DISP_CHANGE_RESTART
         ValidRes = MsgBox("This change will not take effect until you reboot the system.  Reboot now?", vbYesNo)
         If ValidRes = vbYes Then
            UpdtReg = 0
            ValidRes = ExitWindowsEx(EWX_REBOOT, UpdtReg)
         End If
      Case DISP_CHANGE_SUCCESSFUL
         CurrentRes = strRes
      Case Else
         MsgBox "Error changing resolution! Returned: " & ValidRes
   End Select
End Sub

Public Sub ResCheck()
   'variable declaration
   Dim lBits As Integer, lWidth As Integer, lHeight As Integer, Flag As Boolean
   'pull screen resolution information:
   lBits = GetDeviceCaps(frmMain.hDC, BITSPIXEL)
   lWidth = GetDeviceCaps(frmMain.hDC, HORZRES)
   lHeight = GetDeviceCaps(frmMain.hDC, VERTRES)
   'set original resolution to return to:
   oRES = lWidth & "x" & lHeight & "x" & lBits
   'initialize variables:
   NumModes = 0
   Flag = True
   'set maximum resolutions of device (increase this if necessary):
   ReDim DM1(100) As DEVMODE
   'create resolution table Res() and stop when valid resolutions are reached:
   Do While Flag
      Res(NumModes) = DM1(NumModes).dmPelsWidth & "x" _
         & DM1(NumModes).dmPelsHeight & "x" _
         & DM1(NumModes).dmBitsPerPel
      NumModes = NumModes + 1
      Flag = EnumDisplaySettings(ByVal 0, NumModes, DM1(NumModes))
   Loop
   'adjustment for number of modes:
   NumModes = NumModes - 1
   CurrentRes = oRES
End Sub

Sub Main()
   'creates resolution table:
   Call ResCheck
   'run frmMain:
   frmMain.Show
End Sub
