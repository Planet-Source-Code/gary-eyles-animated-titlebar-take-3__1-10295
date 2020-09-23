VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TitlebarCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents PictureBox As PictureBox
Attribute PictureBox.VB_VarHelpID = -1
Public WithEvents sCloseButton As CommandButton
Attribute sCloseButton.VB_VarHelpID = -1
Public WithEvents sRestoreButton As CommandButton
Attribute sRestoreButton.VB_VarHelpID = -1
Public WithEvents sMinimizeButton As CommandButton
Attribute sMinimizeButton.VB_VarHelpID = -1

Private Const WM_PAINT = &HF
Private Const LF_FACESIZE = 32

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(1 To LF_FACESIZE) As Byte
End Type

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function GetObjectAPIBynum Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByVal lpObject As Long) As Long
Private Const SYSTEM_FONT = 13
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Const WM_SETFONT = &H30
Private Const WM_GETFONT = &H31
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const SWW_HPARENT = (-8)
Private Const WM_MOVE = &H3
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const WM_SIZE = &H5
Private Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" (ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()

Private Declare Function ExcludeClipRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function OffsetClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, _
    lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
    
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, ByVal heightSrc As Long, ByVal blendFunct As Long) As Boolean

Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Dim TmpPict As ascMemoryBitmap
Dim AlphaValue As Integer
Dim CloseOnClick As Boolean

Public Event DrawTitlebar()

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function DrawCaption Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long, pcRect As RECT, ByVal un As Long) As Long
 
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Enum TitlebarColors
    COLOR_GRADIENTACTIVECAPTION = 27
    COLOR_GRADIENTINACTIVECAPTION = 28
End Enum
 
Private Enum DrawC
    DC_ACTIVE = &H1
    DC_SMALLCAP = &H2
    DC_ICON = &H4
    DC_TEXT = &H8
    DC_INBUTTON = &H10
    DC_GRADIENT = &H20
End Enum

Implements ISubclass
Private m_emr As EMsgResponse

'Private Type RECT
'        Left As Long
'        Top As Long
'        Right As Long
'        Bottom As Long
'End Type

'Private Type POINTAPI
'        X As Long
'        Y As Long
'End Type

Private Const WM_GETSYSMENU = &H313
Private Const WM_NCPAINT = &H85
Private Const WM_DRAWITEM = &H2B
Private Const WM_ACTIVATE = &H6

Private Type SystemButtons
    IsThere As Boolean
    TheButton As CommandButton
End Type

Enum dButtons
    pCloseButton = 0
    pRestoreButton = 1
    pMinimizeButton = 2
End Enum

Dim gFrm As Form
Dim gBox As PictureBox
Dim FocusIn As Boolean
Dim CloseButton As SystemButtons
Dim RestoreButton As SystemButtons
Dim MinimizeButton As SystemButtons

Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
' Redraw window:
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_ERASE = &H4
Private Const RDW_ERASENOW = &H200
Private Const RDW_FRAME = &H400
Private Const RDW_INTERNALPAINT = &H2
Private Const RDW_INVALIDATE = &H1
Private Const RDW_NOCHILDREN = &H40
Private Const RDW_NOERASE = &H20
Private Const RDW_NOFRAME = &H800
Private Const RDW_NOINTERNALPAINT = &H10
Private Const RDW_UPDATENOW = &H100
Private Const RDW_VALIDATE = &H8

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_VCENTER = &H4
Private Const DT_TOP = &H0
Private Const DT_SINGLELINE = &H20
Private Const DT_RIGHT = &H2
Private Const DT_WORDBREAK = &H10
Private Const DT_CALCRECT = &H400
Private Const DT_WORD_ELLIPSIS = &H40000

Private Type DRAWITEMSTRUCT
   CtlType As Long
   CtlID As Long
   itemID As Long
   itemAction As Long
   itemState As Long
   hwndItem As Long
   hDC As Long
   rcItem As RECT
   itemData As Long
End Type

Enum SysMet
    SM_CXSCREEN = 0
    SM_CYSCREEN = 1
    SM_CXVSCROLL = 2
    SM_CYHSCROLL = 3
    SM_CYCAPTION = 4
    SM_CXBORDER = 5
    SM_CYBORDER = 6
    SM_CXDLGFRAME = 7
    SM_CYDLGFRAME = 8
    SM_CYVTHUMB = 9
    SM_CXHTHUMB = 10
    SM_CXICON = 11
    SM_CYICON = 12
    SM_CXCURSOR = 13
    SM_CYCURSOR = 14
    SM_CYMENU = 15
    SM_CXFULLSCREEN = 16
    SM_CYFULLSCREEN = 17
    SM_CYKANJIWINDOW = 18
    SM_MOUSEPRESENT = 19
    SM_CYVSCROLL = 20
    SM_CXHSCROLL = 21
    SM_DEBUG = 22
    SM_SWAPBUTTON = 23
    SM_RESERVED1 = 24
    SM_RESERVED2 = 25
    SM_RESERVED3 = 26
    SM_RESERVED4 = 27
    SM_CXMIN = 28
    SM_CYMIN = 29
    SM_CXSIZE = 30
    SM_CYSIZE = 31
    SM_CXFRAME = 32
    SM_CYFRAME = 33
    SM_CXMINTRACK = 34
    SM_CYMINTRACK = 35
    SM_CXDOUBLECLK = 36
    SM_CYDOUBLECLK = 37
    SM_CXICONSPACING = 38
    SM_CYICONSPACING = 39
    SM_MENUDROPALIGNMENT = 40
    SM_PENWINDOWS = 41
    SM_DBCSENABLED = 42
    SM_CMOUSEBUTTONS = 43
    SM_CMETRICS = 44
End Enum

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Sub TopZ(frm As Form, Optional YesNo As Boolean = True)
    Dim lRt As Long
    
If YesNo Then
    lRt = SetWindowPos(frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE)
Else
    lRt = SetWindowPos(frm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE)
End If
End Sub

Private Sub RepaintWindow( _
        ByRef objThis As Object, _
        Optional ByVal bClientAreaOnly As Boolean = True _
    )
Dim tR As RECT
Dim tP As POINTAPI
    If (bClientAreaOnly) Then
        GetClientRect objThis.hWnd, tR
    Else
        GetWindowRect objThis.hWnd, tR
        tP.X = tR.Left: tP.Y = tR.Top
        ScreenToClient objThis.hWnd, tP
        tR.Left = tP.X: tR.Top = tP.Y
        tP.X = tR.Right: tP.Y = tR.Bottom
        ScreenToClient objThis.hWnd, tP
        tR.Right = tP.X: tR.Bottom = tP.Y
    End If
    InvalidateRect objThis.hWnd, tR, 1
    
End Sub
        
Private Function MAKELONG(wLow As Long, wHigh As Long) As Long
    MAKELONG = LOWORD(wLow) Or (&H10000 * LOWORD(wHigh))
End Function

Private Function HiWord(ByVal l As Long) As Long
    l = l \ &H10000
    HiWord = Val("&H" & Hex$(l))
End Function

Private Function LOWORD(dwValue As Long) As Long
    CopyMemory LOWORD, dwValue, 2
End Function

Public Sub DrawAlpha(destDC, destLeft, destTop, destWidth, destHeight, srcDC, srcLeft, srcTop, srcWidth, srcHeight, sldBlend)
    Dim Blend As BLENDFUNCTION, BlendLng As Long
    
    Blend.SourceConstantAlpha = 255 - sldBlend
    
    CopyMemory BlendLng, Blend, 4
    
    AlphaBlend destDC, destLeft, destTop, destWidth, destHeight, _
        srcDC, srcLeft, srcTop, srcWidth, srcHeight, BlendLng
End Sub

Public Function GetSysMetrics(index As SysMet) As Long
    GetSysMetrics = GetSystemMetrics(index)
End Function

Private Sub PictureBox_DblClick()
On Error GoTo errOccurred

gFrm.SetFocus
'If GetForegroundWindow <> ParentHwnd Then Exit Sub

Dim xyC As POINTAPI
Dim xyFrame As Long
'Get the size of the frame surrouning
'the form
xyFrame = GetSysMetrics(SM_CXFRAME)
'Get the cursor position
GetCursorPos xyC

If CloseOnClick Then
    If xyC.X - gFrm.Left / 15 - xyFrame < gBox.ScaleHeight Then
        Unload gFrm
        Exit Sub
    End If
End If
    
If gFrm.WindowState = vbNormal Then
    ShowWindow gFrm.hWnd, ShowCommands.SW_gMAXIMIZE
Else
    ShowWindow gFrm.hWnd, ShowCommands.SW_gNORMAL
End If

Call Refresh

errOccurred:
End Sub

Private Sub sDrawCaption(tHwnd As Long, tHdc As Long, tRect As RECT, tFlags As DrawC)
    DrawCaption tHwnd, tHdc, tRect, tFlags
End Sub

Public Sub DrawDefaultCaption(Gradient As Boolean, HasIcon As Boolean, DrawText As Boolean)
Dim rRect As RECT
Dim TheFlags As DrawC
rRect.Left = 0
rRect.Top = 0
rRect.Right = gBox.ScaleWidth
If CloseButton.IsThere = True And RestoreButton.IsThere = False And MinimizeButton.IsThere = False Then
    rRect.Right = rRect.Right - CloseButton.TheButton.Width - 2
ElseIf CloseButton.IsThere Then
    rRect.Right = rRect.Right - CloseButton.TheButton.Width
End If

If RestoreButton.IsThere Then rRect.Right = rRect.Right - RestoreButton.TheButton.Width
If MinimizeButton.IsThere Then rRect.Right = rRect.Right - MinimizeButton.TheButton.Width
rRect.Bottom = gBox.ScaleHeight

If Gradient Then TheFlags = TheFlags Or DC_GRADIENT
If HasIcon Then TheFlags = TheFlags Or DC_ICON
If DrawText Then TheFlags = TheFlags Or DC_TEXT
If FocusIn Then
    TheFlags = TheFlags Or DC_ACTIVE
    gBox.BackColor = GetSysColor(TitlebarColors.COLOR_GRADIENTACTIVECAPTION)
Else
    gBox.BackColor = GetSysColor(TitlebarColors.COLOR_GRADIENTINACTIVECAPTION)
End If

'gBox.BackColor = QBColor(12)
gBox.Cls
sDrawCaption gFrm.hWnd, gBox.hDC, rRect, TheFlags
End Sub

Sub Refresh()
    RaiseEvent DrawTitlebar
    If CloseButton.IsThere Then Redraw CloseButton.TheButton.hWnd
    If RestoreButton.IsThere Then Redraw RestoreButton.TheButton.hWnd
    If MinimizeButton.IsThere Then Redraw MinimizeButton.TheButton.hWnd
End Sub

Private Sub PictureBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo finish

gFrm.SetFocus

Dim xyFrame As Long
Dim tbHeight As Long
xyFrame = GetSysMetrics(SM_CXBORDER)
tbHeight = GetSysMetrics(SM_CYCAPTION)

If X < gBox.ScaleHeight And Button = 1 And CloseOnClick = True Then
    Dim tmpC As POINTAPI
    GetCursorPos tmpC
    SystemMenu gFrm.Left / 15 + xyFrame * 2 + 2, gFrm.Top / 15 + xyFrame + tbHeight + 2
End If

Exit Sub
finish:
ReleaseCapture
Beep
End Sub

Private Sub PictureBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    'Titlebar doesn't wont the click event
    Call ReleaseCapture
ElseIf Button = 2 Then
    'Display system menu when right mouse
    'button is release between the form icon
    'and titlebar buttons
    Dim tmpC As POINTAPI
    Dim TmpX As Long
    GetCursorPos tmpC
    If CloseButton.IsThere Then TmpX = CloseButton.TheButton.Left
    If RestoreButton.IsThere Then TmpX = RestoreButton.TheButton.Left
    If MinimizeButton.IsThere Then TmpX = MinimizeButton.TheButton.Left
    
    If tmpC.X - gFrm.Left / 15 < TmpX Then
        SystemMenu tmpC.X, tmpC.Y
    End If
    Call ReleaseCapture
End If
End Sub

Private Sub PictureBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errOccurred

'gFrm.SetFocus
If GetForegroundWindow <> gFrm.hWnd Then
    Exit Sub
End If

If Button = 1 Then
    Me.MoveTitlebar
End If

errOccurred:
End Sub

Public Sub Redraw(hWnd As Long)
   RedrawWindow hWnd, ByVal 0&, RDW_UPDATENOW, 1
End Sub

Sub SystemMenu(xx As Long, yy As Long)
    SendMessage gFrm.hWnd, WM_GETSYSMENU, 0, ByVal MAKELONG(xx, yy)
End Sub

Public Property Get Focus() As Boolean
   Focus = FocusIn
End Property

Public Property Get Alpha() As Integer
   Alpha = AlphaValue
End Property

Public Property Let Alpha(aValue As Integer)
    On Error GoTo AnError
    If aValue < 0 Then
        aValue = 0
    ElseIf aValue > 255 Then
        aValue = 255
    End If
    AlphaValue = aValue
    
    Exit Property

AnError:
AlphaValue = 255
End Property

Public Property Get HasAnIcon() As Boolean
    HasAnIcon = CloseOnClick
End Property

Public Property Let HasAnIcon(aValue As Boolean)
    CloseOnClick = aValue
End Property

Public Sub DrawTextEx(sText As String, lleft, ttop, rright, bbottom)
Dim tmprect As RECT
tmprect.Left = lleft
tmprect.Top = ttop
tmprect.Right = rright
tmprect.Bottom = bbottom

    DrawText gBox.hDC, sText, -1, tmprect, DT_LEFT Or _
    DT_SINGLELINE Or DT_WORD_ELLIPSIS Or DT_CALCRECT

tmprect.Top = (gBox.ScaleHeight / 2) - (tmprect.Bottom / 2)
tmprect.Bottom = ttop + tmprect.Bottom
'tmprect.Left = gBox.ScaleHeight
tmprect.Left = lleft
'tmprect.Right = tmprect.Right + tmprect.Left
tmprect.Right = rright

    DrawText gBox.hDC, sText, -1, tmprect, DT_LEFT Or _
    DT_SINGLELINE Or DT_WORD_ELLIPSIS
End Sub

Sub MoveTitlebar()
If gFrm.hWnd = ParentHwnd Then
    Dim cFrm As Form
    For Each cFrm In Forms
        If cFrm.Visible And cFrm.hWnd <> ParentHwnd Then
            'TopZ Form1, False
            SetWindowPos ParentHwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
            TopZ cFrm
        End If
    Next
End If

Call ReleaseCapture
SendMessage gFrm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Sub SetButton(bType As dButtons, butt As CommandButton)
If bType = pCloseButton Then
    CloseButton.IsThere = True
    Set CloseButton.TheButton = butt
    Set sCloseButton = butt
ElseIf bType = pRestoreButton Then
    RestoreButton.IsThere = True
    Set RestoreButton.TheButton = butt
    Set sRestoreButton = butt
ElseIf bType = pMinimizeButton Then
    MinimizeButton.IsThere = True
    Set MinimizeButton.TheButton = butt
    Set sMinimizeButton = butt
End If
End Sub

Sub TitleBar(frm As Form, pBox As PictureBox, Optional pHwnd As Boolean = False)
'These two lines make it possible for the
'picturebox to be displayed above the form

If pHwnd Then
    ParentHwnd = frm.hWnd
End If

'These are all the windows messages I want
'to capture
AttachMessage Me, frm.hWnd, WM_MOVE
AttachMessage Me, frm.hWnd, WM_SIZE
AttachMessage Me, frm.hWnd, WM_NCPAINT
AttachMessage Me, frm.hWnd, WM_ACTIVATE
AttachMessage Me, pBox.hWnd, WM_MOVE
AttachMessage Me, pBox.hWnd, WM_DRAWITEM
AttachMessage Me, frm.hWnd, &H401
AttachMessage Me, frm.hWnd, &H402

Set gFrm = frm
Set gBox = pBox
Set Me.PictureBox = pBox

SetParent pBox.hWnd, 0
SetWindowLong pBox.hWnd, SWW_HPARENT, frm.hWnd

FocusIn = True
AlphaValue = 255
End Sub

Sub UnTitlebar()
'Set the picturebox back to the form
'so VB doesn't get confused.
SetWindowLong gFrm.hWnd, SWW_HPARENT, 0

'We have to stop capturing the windows
'messages other wise we make get some
'errors
DetachMessage Me, gFrm.hWnd, WM_MOVE
DetachMessage Me, gFrm.hWnd, WM_SIZE
DetachMessage Me, gFrm.hWnd, WM_NCPAINT
DetachMessage Me, gFrm.hWnd, WM_ACTIVATE
DetachMessage Me, gBox.hWnd, WM_MOVE
DetachMessage Me, gBox.hWnd, WM_DRAWITEM
DetachMessage Me, gFrm.hWnd, &H401
DetachMessage Me, gFrm.hWnd, &H402

If gFrm.hWnd = ParentHwnd Then
    ParentHwnd = 0
End If

Set gBox = Nothing
Set gFrm = Nothing
Set PictureBox = Nothing
Set sCloseButton = Nothing
Set sRestoreButton = Nothing
Set sMinimizeButton = Nothing
End Sub

Private Property Let ISubClass_MsgResponse(ByVal RHS As EMsgResponse)
m_emr = RHS
End Property

Private Property Get ISubClass_MsgResponse() As EMsgResponse
ISubClass_MsgResponse = m_emr
End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim xyFrame As Long
Dim tmpWrct As RECT
Dim tbHeight As Long
Dim ccFrm As Form
  
If iMsg = &H401 Then
    If wParam = 1 Then
        gBox.Visible = False
    Else
        gBox.Visible = True
        TopZ gFrm
    End If
End If

If iMsg = &H402 Then
    If wParam = 1 Or wParam = 2 Then
        FocusIn = True
    Else
        FocusIn = False
        TopZ gFrm, False
    End If

    Call Refresh
End If
   
'Get the position and size of our form
GetWindowRect gFrm.hWnd, tmpWrct

'Get the size of the Caption(titlebar) and
'the size of the frame, which surrounds
'the form
xyFrame = GetSystemMetrics(SysMet.SM_CXFRAME)
tbHeight = GetSystemMetrics(SysMet.SM_CYCAPTION)



If iMsg = WM_MOVE And hWnd = gFrm.hWnd Then
    'When moving the form we want to move
    'the picturebox as well. Other wise
    'it won't look like a titlebar
    SetWindowPos gBox.hWnd, _
        0, LOWORD(lParam), _
        gFrm.Top / 15 + xyFrame, _
        tmpWrct.Right - tmpWrct.Left - xyFrame * 2, tbHeight - 1, 0
        
    Call Refresh
ElseIf iMsg = WM_SIZE And hWnd = gFrm.hWnd Then
    'Changes the size of titlebar
    'depending on the size of the form
    gBox.Width = gFrm.Width - xyFrame * 30

    If wParam = 1 Then
        ParentMinimized = True
    Else
        ParentMinimized = False
    End If
         
    For Each ccFrm In Forms
        SendMessage ccFrm.hWnd, &H401, wParam, 0
    Next
                 
    Call Refresh
        
    CallOldWindowProc hWnd, iMsg, wParam, lParam
    Exit Function

ElseIf iMsg = WM_NCPAINT Then
    'If we don't stop the old titlebar from
    'redrawing we can see it when we move
    'the form about
    Dim tmpDC As Long
    Dim hRgn As Long
    tmpDC = GetWindowDC(gFrm.hWnd)
    
    'I'm not too sure about all these commands
    'but they do seem to work quite well
    With tmpWrct
        hRgn = CreateRectRgn(.Left, .Top, .Right, .Bottom)
    End With
    ExcludeClipRect tmpDC, xyFrame, xyFrame, tmpWrct.Right - tmpWrct.Left - 4, xyFrame + tbHeight - 1 ' + 20
    OffsetClipRgn tmpDC, tmpWrct.Left, tmpWrct.Top
    GetClipRgn tmpDC, hRgn
    ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, hRgn, lParam)
    DeleteObject hRgn

ElseIf hWnd = gBox.hWnd And iMsg = WM_DRAWITEM Then
    'We draw are buttons here
    Dim tDis As DRAWITEMSTRUCT
    Dim DrawSty As Long
    CopyMemory tDis, ByVal lParam, Len(tDis)

    If tDis.hwndItem = CloseButton.TheButton.hWnd Then
        'Set drawing style to 'Close'
        DrawSty = &H0
    ElseIf tDis.hwndItem = RestoreButton.TheButton.hWnd Then
        If gFrm.WindowState = vbNormal Then
        'Set drawing style to 'Maximize'
            DrawSty = &H2
        Else
        'Set drawing style to 'Restore'
            DrawSty = &H3
        End If
    ElseIf tDis.hwndItem = MinimizeButton.TheButton.hWnd Then
        'Set drawing style to 'Minimize'
        DrawSty = &H1
    End If
    
    If tDis.itemState = 1 Then
    'Set drawing style to 'button pushed'
        DrawSty = DrawSty Or &H200
    ElseIf tDis.itemState = 4 Then
    'Set drawing style to 'disabled'
        DrawSty = DrawSty Or &H100
    End If
'    'Api that allows you to draw system buttons
'    'with the greatest of ease


    Dim tmprect As RECT
    Dim tmprect2 As RECT
    GetWindowRect tDis.hwndItem, tmprect
    GetWindowRect gBox.hWnd, tmprect2

    Set TmpPict = New ascMemoryBitmap
    TmpPict.CreateBlank tDis.rcItem.Right, tDis.rcItem.Bottom
    
Dim aRct As RECT

If tDis.hwndItem = CloseButton.TheButton.hWnd Then
    aRct.Bottom = tDis.rcItem.Bottom - 2
    aRct.Left = tDis.rcItem.Left
    aRct.Right = tDis.rcItem.Right - 2
    aRct.Top = tDis.rcItem.Top + 2
    BitBlt TmpPict.hDC, 0, 0, tDis.rcItem.Right, tDis.rcItem.Bottom, gBox.hDC, CloseButton.TheButton.Left, 0, vbSrcCopy
ElseIf tDis.hwndItem = RestoreButton.TheButton.hWnd Then
    aRct.Bottom = tDis.rcItem.Bottom - 2
    aRct.Left = tDis.rcItem.Left
    aRct.Right = tDis.rcItem.Right - 2
    aRct.Top = tDis.rcItem.Top + 2
    BitBlt TmpPict.hDC, 0, 0, tDis.rcItem.Right, tDis.rcItem.Bottom, gBox.hDC, RestoreButton.TheButton.Left, 0, vbSrcCopy
ElseIf tDis.hwndItem = MinimizeButton.TheButton.hWnd Then
    aRct.Bottom = tDis.rcItem.Bottom - 2
    aRct.Left = tDis.rcItem.Left + 2
    aRct.Right = tDis.rcItem.Right
    aRct.Top = tDis.rcItem.Top + 2
    BitBlt TmpPict.hDC, 0, 0, tDis.rcItem.Right, tDis.rcItem.Bottom, gBox.hDC, MinimizeButton.TheButton.Left, 0, vbSrcCopy
End If
    
    DrawFrameControl TmpPict.hDC, aRct, 1, DrawSty
    
    If AlphaValue <> 255 Then
        'DrawAlpha TmpPict.hDC, 0, 0, tDis.rcItem.Right, tDis.rcItem.Bottom, gBox.hDC, tmprect.Left - gFrm.Left / 15 - 3, tmprect.Top - gFrm.Top / 15 - 3, tDis.rcItem.Right, tDis.rcItem.Bottom, AlphaValue
        DrawAlpha TmpPict.hDC, 0, 0, tDis.rcItem.Right, tDis.rcItem.Bottom, gBox.hDC, tmprect.Left - tmprect2.Left, 0, tDis.rcItem.Right, tDis.rcItem.Bottom, AlphaValue
    End If
    
    BitBlt tDis.hDC, 0, 0, tDis.rcItem.Right, tDis.rcItem.Bottom, TmpPict.hDC, 0, 0, vbSrcCopy
    TmpPict.ClearUp
    Set TmpPict = Nothing

    Exit Function

ElseIf iMsg = WM_ACTIVATE Then
    'This message is called every time
    'the form either gets the focus
    'or loses its focus
    
    For Each ccFrm In Forms
            SendMessage ccFrm.hWnd, &H402, wParam, 0
    Next
End If

'Changes the size and position of the
'buttons depending of the size of the
'titlebar and the form

If CloseButton.IsThere Then
    CloseButton.TheButton.Top = 0
    CloseButton.TheButton.Width = tbHeight - 1
    CloseButton.TheButton.Left = gFrm.ScaleWidth - CloseButton.TheButton.Width
    CloseButton.TheButton.Height = tbHeight - 1
End If

If RestoreButton.IsThere Then
    RestoreButton.TheButton.Top = 0
    RestoreButton.TheButton.Width = tbHeight - 1
    RestoreButton.TheButton.Left = gFrm.ScaleWidth - (RestoreButton.TheButton.Width * 2)
    RestoreButton.TheButton.Height = tbHeight - 1
End If

If MinimizeButton.IsThere Then
    MinimizeButton.TheButton.Top = 0
    MinimizeButton.TheButton.Width = tbHeight - 1
    MinimizeButton.TheButton.Left = gFrm.ScaleWidth - (MinimizeButton.TheButton.Width * 3)
    MinimizeButton.TheButton.Height = tbHeight - 1
End If

Exit Function

If CloseButton.IsThere Then
    CloseButton.TheButton.Top = 2
    CloseButton.TheButton.Width = tbHeight - 3
    CloseButton.TheButton.Left = gFrm.ScaleWidth - CloseButton.TheButton.Width - xyFrame / 2
    CloseButton.TheButton.Height = tbHeight - 5
End If

If RestoreButton.IsThere Then
    RestoreButton.TheButton.Top = 2
    RestoreButton.TheButton.Width = tbHeight - 3
    RestoreButton.TheButton.Left = gFrm.ScaleWidth - (RestoreButton.TheButton.Width * 2) - xyFrame / 2 - 2
    RestoreButton.TheButton.Height = tbHeight - 5
End If

If MinimizeButton.IsThere Then
    MinimizeButton.TheButton.Top = 2
    MinimizeButton.TheButton.Width = tbHeight - 3
    MinimizeButton.TheButton.Left = gFrm.ScaleWidth - (MinimizeButton.TheButton.Width * 3) - xyFrame / 2 - 2
    MinimizeButton.TheButton.Height = tbHeight - 5
End If
End Function

Private Sub sCloseButton_Click()
Unload gFrm
End Sub

Private Sub sCloseButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errOccurred
    
gFrm.SetFocus
Exit Sub

errOccurred:
ReleaseCapture
Beep
End Sub

Private Sub sMinimizeButton_Click()
'Minimize the form
ShowWindow gFrm.hWnd, ShowCommands.SW_gMINIMIZE
End Sub

Private Sub sMinimizeButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errOccurred
    
gFrm.SetFocus
Exit Sub

errOccurred:
ReleaseCapture
Beep
End Sub

Private Sub sRestoreButton_Click()
'Restore or Maximize the form
'depending on the windowstate
If gFrm.WindowState = vbNormal Then
    ShowWindow gFrm.hWnd, ShowCommands.SW_gMAXIMIZE
Else
    ShowWindow gFrm.hWnd, ShowCommands.SW_gNORMAL
End If
'Make sure the button displays the
'correct picture when the form
'has changed
RestoreButton.TheButton.Refresh
End Sub

Private Sub sRestoreButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errOccurred
    
gFrm.SetFocus
Exit Sub

errOccurred:
ReleaseCapture
Beep
End Sub

