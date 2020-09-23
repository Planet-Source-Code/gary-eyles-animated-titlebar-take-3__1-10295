VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Animated Titlebar"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   Icon            =   "Titlebar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   274
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check4 
      Caption         =   "Animate buttons"
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show Modal"
      Height          =   495
      Left            =   240
      TabIndex        =   18
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Standard titlebar"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   1575
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   3360
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   10
      Max             =   255
      TickFrequency   =   10
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Animate titlebar"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Simple"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1080
      Top             =   3720
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Menu visible"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   4800
      Picture         =   "Titlebar.frx":164A
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   4560
      Picture         =   "Titlebar.frx":D68C
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.CommandButton Sysbut 
      Caption         =   "Close"
      Height          =   375
      Index           =   0
      Left            =   3600
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Sysbut 
      Caption         =   "Restore"
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   5
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Sysbut 
      Caption         =   "Minimize"
      Height          =   375
      Index           =   2
      Left            =   1920
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
   Begin VB.PictureBox TitleBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   120
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   297
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton sMinimize 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   6.75
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   30
         Width           =   240
      End
      Begin VB.CommandButton sRestore 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   6.75
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   30
         Width           =   240
      End
      Begin VB.CommandButton sClose 
         Caption         =   "r"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   6.75
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   30
         Width           =   240
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Have transparent system buttons, set value from 0 to 255."
      Height          =   855
      Left            =   3000
      TabIndex        =   16
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000003&
      Height          =   1815
      Left            =   1920
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "System button alpha transparency"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   3120
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "High"
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Low"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3720
      Width           =   495
   End
   Begin VB.Image Image4 
      Height          =   270
      Left            =   2040
      Picture         =   "Titlebar.frx":196CE
      Top             =   2280
      Width           =   810
   End
   Begin VB.Image Image3 
      Height          =   270
      Left            =   2040
      Picture         =   "Titlebar.frx":1A298
      Top             =   1920
      Width           =   810
   End
   Begin VB.Image Image2 
      Height          =   270
      Left            =   2040
      Picture         =   "Titlebar.frx":1AE62
      Top             =   1560
      Width           =   810
   End
   Begin VB.Image Image1 
      Height          =   270
      Left            =   2040
      Picture         =   "Titlebar.frx":1BA2C
      Top             =   1200
      Width           =   810
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IsInFocus As Boolean
Dim xad As Long
Dim TheAlpha As Long
Dim AlphaAdd As Long

Private Const constalpha = 10

Private WithEvents pTitlebar As TitlebarCustom
Attribute pTitlebar.VB_VarHelpID = -1

Private Sub Check1_Click()
Dim cc As Object
'Search's through all the objects finding
'all the menus, then makes them visible or
'hides them depending on if the check box
'is checked or not.
If Check1.Value = 1 Then
    For Each cc In Me
        If TypeOf cc Is Menu Then
            cc.Visible = True
        End If
    Next
Else
    For Each cc In Me
        If TypeOf cc Is Menu Then
            cc.Visible = False
        End If
    Next
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Timer1.Interval = 1
    Check4.Enabled = True
Else
    Timer1.Interval = 0
    Check4.Enabled = False
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
    Check2.Enabled = False
    Slider1.Value = 255
    pTitlebar.Alpha = 255
Else
    Check2.Enabled = True
End If
pTitlebar.Refresh
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
    Slider1.Enabled = False
Else
    Slider1.Enabled = True
End If
End Sub

Private Sub Command1_Click()
Dim snew As Form
Set snew = New Form2
Load snew
snew.Show 0, Me
End Sub

Private Sub Command2_Click()
Form2.Show 1, Me
End Sub

Private Sub Form_Load()
Set pTitlebar = New TitlebarCustom

pTitlebar.TitleBar Me, TitleBar, True
pTitlebar.SetButton pCloseButton, sClose
pTitlebar.SetButton pRestoreButton, sRestore
pTitlebar.SetButton pMinimizeButton, sMinimize
pTitlebar.HasAnIcon = True
pTitlebar.Alpha = 150

Slider1.Value = pTitlebar.Alpha
End Sub

Private Sub Form_Unload(Cancel As Integer)
pTitlebar.UnTitlebar

Dim ccFrm As Form
For Each ccFrm In Forms
     Unload ccFrm
Next
End Sub

Private Sub Slider1_Click()
pTitlebar.Alpha = Slider1.Value
pTitlebar.Refresh
End Sub

Private Sub Slider1_Scroll()
pTitlebar.Alpha = Slider1.Value
pTitlebar.Refresh
End Sub

Private Sub Sysbut_Click(index As Integer)
'Disables or Enables one of
'the titlebar buttons
If index = 0 Then
    If sClose.Enabled Then
        sClose.Enabled = False
    Else
        sClose.Enabled = True
    End If
ElseIf index = 1 Then
    If sRestore.Enabled Then
        sRestore.Enabled = False
    Else
        sRestore.Enabled = True
    End If
ElseIf index = 2 Then
    If sMinimize.Enabled Then
        sMinimize.Enabled = False
    Else
        sMinimize.Enabled = True
    End If
End If
End Sub

Private Sub Timer1_Timer()
'This animates the titlebar
'Simply delete if you don't
'wont it animated
xad = xad + 5
If xad > Picture1.ScaleHeight Then
    xad = 0
End If

TheAlpha = TheAlpha + AlphaAdd
If TheAlpha >= 255 Then
    AlphaAdd = -constalpha
    TheAlpha = 255
ElseIf TheAlpha <= 0 Then
    AlphaAdd = constalpha
    TheAlpha = 0
End If

If Check4.Value = 1 Then
    pTitlebar.Alpha = TheAlpha
Else
    pTitlebar.Alpha = Slider1.Value
End If

pTitlebar.Refresh
End Sub

Public Sub pTitlebar_DrawTitlebar()
If Check3.Value = 1 Then
    pTitlebar.DrawDefaultCaption True, True, True
        
    'Put more drawing command here if you want something
    'extra on the default titlebar. e.g. Like a picture.
    
    TitleBar.Refresh
    Exit Sub
End If

Dim TheIcon As Long
Dim xx, yy As Long
TheIcon = Me.Icon

With TitleBar
'Clear the titlebar
.Cls

'Depending whether the form is in focus
'or not, tile one of the following pictures
If pTitlebar.Focus Then
    'Form is in focus so tile the redish picture
    'from Picture1
    For xx = 0 To Int(.ScaleWidth / Picture1.ScaleHeight) + 1
        BitBlt .hDC, xx * Picture1.Width - xad, 0, Picture1.Width, Picture1.Height, Picture1.hDC, 0, 0, vbSrcCopy
    Next
Else
    'Form isn't in focus so tile the redish picture
    'from Picture2
    For xx = 0 To Int(.ScaleWidth / Picture2.ScaleHeight) + 1
    BitBlt .hDC, xx * Picture2.Width - xad, 0, Picture2.Width, Picture2.Height, Picture2.hDC, 0, 0, vbSrcCopy
    Next
End If

'Draw the forms icons in the top left
DrawIconEx .hDC, 1, 1, TheIcon, .ScaleHeight - 2, .ScaleHeight - 2, ByVal 0&, ByVal 0&, &H3
pTitlebar.DrawTextEx Me.Caption, .ScaleHeight + 5, 0, sMinimize.Left, .ScaleHeight

'Refresh the titlebar in order for the
'final result to show
.Refresh
End With
End Sub

