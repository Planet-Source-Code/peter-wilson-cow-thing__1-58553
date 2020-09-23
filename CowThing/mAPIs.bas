Attribute VB_Name = "mAPIs"
Option Explicit

' SetWindowPos Flags
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

' The GetWindowLong function retrieves information about the specified window.
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

' The SetWindowLong function changes an attribute of the specified window.
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

' The SetLayeredWindowAttributes function sets the opacity and transparency color key of a layered window.
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_TRANSPARENT = &H20&
Public Const LWA_COLORKEY = &H1&
Public Const LWA_ALPHA = &H2&



' ===============================================================================================================
' The TransparentBlt function performs a bit-block transfer of the color data corresponding
' to a rectangle of pixels from the specified source device context into a destination device context.
' ---------------------------------------------------------------------------------------------------------------
'    hdcDest
'       [in] Handle to the destination device context.
'    nXOriginDest
'       [in] Specifies the x-coordinate, in logical units, of the upper-left corner of the destination rectangle.
'    nYOriginDest
'       [in] Specifies the y-coordinate, in logical units, of the upper-left corner of the destination rectangle.
'    nWidthDest
'       [in] Specifies the width, in logical units, of the destination rectangle.
'    hHeightDest
'       [in] Handle to the height, in logical units, of the destination rectangle.
'    hdcSrc
'       [in] Handle to the source device context.
'    nXOriginSrc
'       [in] Specifies the x-coordinate, in logical units, of the source rectangle.
'    nYOriginSrc
'       [in] Specifies the y-coordinate, in logical units, of the source rectangle.
'    nWidthSrc
'       [in] Specifies the width, in logical units, of the source rectangle.
'    nHeightSrc
'       [in] Specifies the height, in logical units, of the source rectangle.
'    crTransparent
'       [in] The RGB color in the source bitmap to treat as transparent.
' ===============================================================================================================
Public Declare Function TransparentBlt Lib "msimg32" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long


