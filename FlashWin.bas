Attribute VB_Name = "FlashWindow"
'tel, 2005
Option Explicit
Private Declare Function FlashWindowEx Lib "user32" (pfwi As PFLASHWINFO) As Long

Private Type PFLASHWINFO
  cbSize        As Long 'Size of the structure, in bytes.
  hWnd          As Long 'Handle to the window to be flashed. The window can be either opened or minimized.
  dwFlags       As Long 'Flash status. This parameter can be one or more of the following values.
  uCount        As Long 'Number of times to flash the window.
  dwTimeout     As Long 'Rate at which the window is to be flashed, in milliseconds. If dwTimeout is zero, the function uses the default cursor blink rate.
End Type

'dwFlags
Public Enum FLASH_FLAGS
    FLASHW_STOP = 0         'Stop flash
    FLASHW_TRAY = 2         'Flash the taskbar button.
    FLASHW_CAPTION = 1      'caption only
    FLASHW_ALL = 3          'caption + task bar
    FLASHW_TIMERNOFG = 12   'Flash continuously until the window comes to the foreground. (doesnt seem to work)
End Enum

Dim FLASHWINFO As PFLASHWINFO

''Use
'FlashWin Me.hWnd, FLASHW_TRAY
'FlashWin Me.hWnd, FLASHW_STOP

Sub FlashWin(hWnd As Long, flags As FLASH_FLAGS)
    FLASHWINFO.cbSize = Len(FLASHWINFO)
    FLASHWINFO.hWnd = hWnd
    FLASHWINFO.dwTimeout = 0
    FLASHWINFO.uCount = 255
    FLASHWINFO.dwFlags = flags
    Call FlashWindowEx(FLASHWINFO)
End Sub


