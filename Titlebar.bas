Attribute VB_Name = "Module1"
Option Explicit

Global ParentHwnd As Long
Global ParentMinimized As Boolean

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

Public Enum ShowCommands
    SW_gHIDE = 0
    SW_gSHOWNORMAL = 1
    SW_gNORMAL = 1
    SW_gSHOWMINIMIZED = 2
    SW_gSHOWMAXIMIZED = 3
    SW_gMAXIMIZE = 3
    SW_gSHOWNOACTIVATE = 4
    SW_gSHOW = 5
    SW_gMINIMIZE = 6
    SW_gSHOWMINNOACTIVE = 7
    SW_gSHOWNA = 8
    SW_gRESTORE = 9
    SW_gSHOWDEFAULT = 10
    SW_gMAX = 10
End Enum
