Attribute VB_Name = "modGetDisplay"
Option Explicit

Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CXSCREEN = 0 'X Size of screen
Public Const SM_CYSCREEN = 1 'Y Size of Screen
Public Const SM_CXVSCROLL = 2 'X Size of arrow in vertical scroll bar.
Public Const SM_CYHSCROLL = 3 'Y Size of arrow in horizontal scroll bar
Public Const SM_CYCAPTION = 4 'Height of windows caption
Public Const SM_CXBORDER = 5 'Width of no-sizable borders
Public Const SM_CYBORDER = 6 'Height of non-sizable borders
Public Const SM_CXDLGFRAME = 7 'Width of dialog box borders
Public Const SM_CYDLGFRAME = 8 'Height of dialog box borders
Public Const SM_CYHTHUMB = 9 'Height of scroll box on horizontal scroll bar
Public Const SM_CXHTHUMB = 10 ' Width of scroll box on horizontal scroll bar
Public Const SM_CXICON = 11 'Width of standard icon
Public Const SM_CYICON = 12 'Height of standard icon
Public Const SM_CXCURSOR = 13 'Width of standard cursor
Public Const SM_CYCURSOR = 14 'Height of standard cursor
Public Const SM_CYMENU = 15 'Height of menu
Public Const SM_CXFULLSCREEN = 16 'Width of client area of maximized window
Public Const SM_CYFULLSCREEN = 17 'Height of client area of maximized window
Public Const SM_CYKANJIWINDOW = 18 'Height of Kanji window
Public Const SM_MOUSEPRESENT = 19 'True is a mouse is present
Public Const SM_CYVSCROLL = 20 'Height of arrow in vertical scroll bar
Public Const SM_CXHSCROLL = 21 'Width of arrow in vertical scroll bar
Public Const SM_DEBUG = 22 'True if deugging version of windows is running
Public Const SM_SWAPBUTTON = 23 'True if left and right buttons are swapped.
Public Const SM_CXMIN = 28 'Minimum width of window
Public Const SM_CYMIN = 29 'Minimum height of window
Public Const SM_CXSIZE = 30 'Width of title bar bitmaps
Public Const SM_CYSIZE = 31 'height of title bar bitmaps
Public Const SM_CXMINTRACK = 34 'Minimum tracking width of window
Public Const SM_CYMINTRACK = 35 'Minimum tracking height of window
Public Const SM_CXDOUBLECLK = 36 'double click width
Public Const SM_CYDOUBLECLK = 37 'double click height
Public Const SM_CXICONSPACING = 38 'width between desktop icons
Public Const SM_CYICONSPACING = 39 'height between desktop icons
Public Const SM_MENUDROPALIGNMENT = 40 'Zero if popup menus are aligned to the left of the memu bar item. True if it is aligned to the right.
Public Const SM_PENWINDOWS = 41 'The handle of the pen windows DLL if loaded.
Public Const SM_DBCSENABLED = 42 'True if double byte characteds are enabled
Public Const SM_CMOUSEBUTTONS = 43 'Number of mouse buttons.
Public Const SM_CMETRICS = 44 'Number of system metrics
Public Const SM_CLEANBOOT = 67 'Windows 95 boot mode. 0 = normal, 1 = safe, 2 = safe with network
Public Const SM_CXMAXIMIZED = 61 'default width of win95 maximised window
Public Const SM_CXMAXTRACK = 59 'maximum width when resizing win95 windows
Public Const SM_CXMENUCHECK = 71 'width of menu checkmark bitmap
Public Const SM_CXMENUSIZE = 54 'width of button on menu bar
Public Const SM_CXMINIMIZED = 57 'width of rectangle into which minimised windows must fit.
Public Const SM_CYMAXIMIZED = 62 'default height of win95 maximised window
Public Const SM_CYMAXTRACK = 60 'maximum width when resizing win95 windows
Public Const SM_CYMENUCHECK = 72 'height of menu checkmark bitmap
Public Const SM_CYMENUSIZE = 55 'height of button on menu bar
Public Const SM_CYMINIMIZED = 58 'height of rectangle into which minimised windows must fit.
Public Const SM_CYSMCAPTION = 51 'height of windows 95 small caption
Public Const SM_MIDEASTENABLED = 74 'Hebrw and Arabic enabled for windows 95
Public Const SM_NETWORK = 63 'bit o is set if a network is present. Const SM_SECURE = 44 'True if security is present on windows 95 system
Public Const SM_SLOWMACHINE = 73 'true if machine is too slow to run win95.

