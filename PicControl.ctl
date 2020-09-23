VERSION 5.00
Begin VB.UserControl PicControl 
   Appearance      =   0  'Flat
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4110
   ScaleHeight     =   257
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   274
   Begin VB.Timer tmrInsert 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2640
      Top             =   4080
   End
   Begin VB.Timer tmrZoom 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2160
      Top             =   4080
   End
   Begin VB.PictureBox picTools 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   3000
      ScaleHeight     =   223
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   14
      Top             =   0
      Width           =   1095
      Begin VB.PictureBox picPosition 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         ScaleHeight     =   39
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   55
         TabIndex        =   25
         Top             =   120
         Width           =   855
         Begin VB.PictureBox picInsert 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            FillColor       =   &H00FF0000&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   0
            MouseIcon       =   "PicControl.ctx":0000
            MousePointer    =   99  'Custom
            ScaleHeight     =   17
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   17
            TabIndex        =   26
            Top             =   0
            Visible         =   0   'False
            Width           =   285
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "PicControl.ctx":0CCA
         Left            =   45
         List            =   "PicControl.ctx":0D0A
         TabIndex        =   24
         Top             =   1020
         Width           =   975
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   480
         Max             =   7
         TabIndex        =   17
         Top             =   3000
         Value           =   2
         Width           =   495
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "PicControl.ctx":0D69
         Left            =   45
         List            =   "PicControl.ctx":0D79
         TabIndex        =   16
         ToolTipText     =   "Magnification Size"
         Top             =   2280
         Width           =   975
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "PicControl.ctx":0D95
         Left            =   45
         List            =   "PicControl.ctx":0DA2
         TabIndex        =   15
         ToolTipText     =   "Lens Size"
         Top             =   1590
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Scroll Speed"
         ForeColor       =   &H00F3855A&
         Height          =   255
         Left            =   45
         TabIndex        =   23
         Top             =   810
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         ForeColor       =   &H00F3855A&
         Height          =   255
         Left            =   0
         TabIndex        =   21
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Picture Size"
         ForeColor       =   &H00F3855A&
         Height          =   255
         Left            =   45
         TabIndex        =   20
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Zoom 3x etc"
         ForeColor       =   &H00F3855A&
         Height          =   255
         Left            =   45
         TabIndex        =   19
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Lens Size"
         ForeColor       =   &H00F3855A&
         Height          =   255
         Left            =   45
         TabIndex        =   18
         Top             =   1395
         Width           =   975
      End
      Begin VB.Image Image2 
         Height          =   4035
         Left            =   -120
         Picture         =   "PicControl.ctx":0DBC
         Stretch         =   -1  'True
         Top             =   -480
         Width           =   1305
      End
   End
   Begin VB.Timer tmrScrollRight 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   1680
      Top             =   4080
   End
   Begin VB.Timer tmrScrollLeft 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   1200
      Top             =   4080
   End
   Begin VB.Timer tmrScrollDown 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   720
      Top             =   4080
   End
   Begin VB.Timer tmrScrollUp 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   240
      Top             =   4080
   End
   Begin VB.PictureBox picHolder 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   249
      TabIndex        =   1
      Top             =   3480
      Width           =   3735
      Begin VB.PictureBox picHolderInsert 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   480
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   184
         TabIndex        =   2
         Top             =   0
         Width           =   2760
         Begin VB.PictureBox picButton 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   30
            Picture         =   "PicControl.ctx":1209
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   12
            ToolTipText     =   "Left-botto corner"
            Top             =   30
            Width           =   255
         End
         Begin VB.PictureBox picButton 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   300
            Picture         =   "PicControl.ctx":15B9
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   11
            ToolTipText     =   "Right-bottom corner"
            Top             =   30
            Width           =   255
         End
         Begin VB.PictureBox picButton 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   570
            Picture         =   "PicControl.ctx":1969
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   10
            ToolTipText     =   "Centered"
            Top             =   30
            Width           =   255
         End
         Begin VB.PictureBox picButton 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   840
            Picture         =   "PicControl.ctx":1D13
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   9
            ToolTipText     =   "Upper-left corner"
            Top             =   30
            Width           =   255
         End
         Begin VB.PictureBox picButton 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   1110
            Picture         =   "PicControl.ctx":20C3
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   8
            ToolTipText     =   "Upper-right corner"
            Top             =   30
            Width           =   255
         End
         Begin VB.PictureBox picButton 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   1380
            Picture         =   "PicControl.ctx":2473
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   7
            ToolTipText     =   "Scroll Left"
            Top             =   30
            Width           =   255
         End
         Begin VB.PictureBox picButton 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   1650
            Picture         =   "PicControl.ctx":2826
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   6
            ToolTipText     =   "Scroll right"
            Top             =   30
            Width           =   255
         End
         Begin VB.PictureBox picButton 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   1920
            Picture         =   "PicControl.ctx":2BD7
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   5
            ToolTipText     =   "Scroll up"
            Top             =   30
            Width           =   255
         End
         Begin VB.PictureBox picButton 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   2190
            Picture         =   "PicControl.ctx":2F8C
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   4
            ToolTipText     =   "Scroll down"
            Top             =   30
            Width           =   255
         End
         Begin VB.PictureBox picButton 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   9
            Left            =   2460
            Picture         =   "PicControl.ctx":333F
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   3
            ToolTipText     =   "Enable magnifying glass "
            Top             =   30
            Width           =   255
         End
         Begin VB.Image Image1 
            Height          =   315
            Left            =   0
            Picture         =   "PicControl.ctx":36FD
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2745
         End
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   0
      ScaleHeight     =   223
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   199
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.PictureBox picMain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   0
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   57
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   855
         Begin VB.PictureBox picZoom 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            FillColor       =   &H00FF0000&
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   0
            MouseIcon       =   "PicControl.ctx":3B4A
            MousePointer    =   99  'Custom
            ScaleHeight     =   25
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   25
            TabIndex        =   22
            Top             =   0
            Visible         =   0   'False
            Width           =   405
         End
      End
   End
   Begin VB.Image picDown 
      Height          =   255
      Index           =   3
      Left            =   0
      Picture         =   "PicControl.ctx":4814
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image picDown 
      Height          =   255
      Index           =   9
      Left            =   0
      Picture         =   "PicControl.ctx":4BDB
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image picDown 
      Height          =   255
      Index           =   8
      Left            =   0
      Picture         =   "PicControl.ctx":4FA1
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image picDown 
      Height          =   255
      Index           =   7
      Left            =   0
      Picture         =   "PicControl.ctx":5361
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image picDown 
      Height          =   255
      Index           =   6
      Left            =   0
      Picture         =   "PicControl.ctx":5726
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image picDown 
      Height          =   255
      Index           =   5
      Left            =   0
      Picture         =   "PicControl.ctx":5AE3
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image picDown 
      Height          =   255
      Index           =   4
      Left            =   0
      Picture         =   "PicControl.ctx":5EA1
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image picDown 
      Height          =   255
      Index           =   2
      Left            =   0
      Picture         =   "PicControl.ctx":6266
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image picDown 
      Height          =   255
      Index           =   1
      Left            =   0
      Picture         =   "PicControl.ctx":661A
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image picDown 
      Height          =   255
      Index           =   0
      Left            =   0
      Picture         =   "PicControl.ctx":69D8
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image picUp 
      Height          =   255
      Index           =   9
      Left            =   2160
      Picture         =   "PicControl.ctx":6D9A
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image picUp 
      Height          =   255
      Index           =   8
      Left            =   720
      Picture         =   "PicControl.ctx":7158
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image picUp 
      Height          =   255
      Index           =   7
      Left            =   480
      Picture         =   "PicControl.ctx":750F
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image picUp 
      Height          =   255
      Index           =   6
      Left            =   240
      Picture         =   "PicControl.ctx":78C9
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image picUp 
      Height          =   255
      Index           =   5
      Left            =   0
      Picture         =   "PicControl.ctx":7C7E
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image picUp 
      Height          =   255
      Index           =   4
      Left            =   960
      Picture         =   "PicControl.ctx":8036
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image picUp 
      Height          =   255
      Index           =   3
      Left            =   1440
      Picture         =   "PicControl.ctx":83E9
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image picUp 
      Height          =   255
      Index           =   2
      Left            =   1680
      Picture         =   "PicControl.ctx":879C
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image picUp 
      Height          =   255
      Index           =   1
      Left            =   1200
      Picture         =   "PicControl.ctx":8B47
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image picUp 
      Height          =   255
      Index           =   0
      Left            =   1920
      Picture         =   "PicControl.ctx":8EFA
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "PicControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Types
Private Type PictureGate
    bWidth                        As Boolean
    bHeight                       As Boolean
End Type

Private Type POINTAPI                ' general use. Typically used for cursor location
    X                             As Long
    Y                             As Long
End Type

Private Type PosProps
    pL                            As Boolean
    pR                            As Boolean
    pB                            As Boolean
    pt                            As Boolean
End Type

Private Type Size
    sWidth                        As Integer
    sHeight                       As Integer
End Type

' Enums
Public Enum Appearance
    [peFlat] = 0
    [pe3D] = 1
End Enum
#If False Then
Private peFlat, pe3D
#End If

Public Enum AlignNB
    [peLeft] = 0
    [peRight] = 1
    [peCenter] = 2
End Enum
#If False Then
Private peLeft, peRight, peCenter
#End If

Public Enum BorderStyle
    [peNone] = 0
    [peFixed_Single] = 1
End Enum
#If False Then
Private peNone, peFixed_Single
#End If

Public Enum LenzSizeProps
    [peSmall] = 0
    [peMedium] = 1
    [peLarge] = 2
End Enum
#If False Then
Private peSmall, peMedium, peLarge
#End If

Public Enum MagnifyProperties
    [pe150%] = 1.5
    [pe200%] = 2
    [pe250%] = 2.5
    [pe300%] = 3
End Enum
#If False Then
Private pe150%, pe200%, pe250%, pe300%
#End If

'Constants
Private Const SrcCopy              As Long = &HCC0020
Private Const mDefScrollSpeed      As Integer = 50
Private Const mDefBorderStyle      As Integer = 1
Private Const mDefAppearance       As Integer = 1
Private Const mDefBackColor        As Long = &H8000000F
Private Const mDefAlignNB          As Integer = 2
Private Const mdefLS               As Integer = 0
Private Const mDefMag              As Integer = 2
Private Const SS                   As Integer = 250   ' minimum width & height

'Declarations
Private b                          As Picture
Private bw                         As Long
Private bh                         As Long
Private MagnifySize                As Single
Private LenSize                    As LenzSizeProps
Private mAlignNB                   As AlignNB         ' Controlbox alignment
Private SizeDiff                   As Size            ' between picFrame & picMain (diff for each pic)
Private mScrollSpeed               As Integer
Private bGate                      As Boolean
Private mBorderStyle               As BorderStyle
Private mAppearance                As Appearance
Private picGate                    As PictureGate
Private picMainpos                 As PosProps

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)  'Makes it sleep
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, _
       ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
       ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
       ByVal ySrc As Long, ByVal nSrcWidth As Long, _
       ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Property Get AlignNavBar() As AlignNB
AlignNavBar = mAlignNB
End Property
Public Property Let AlignNavBar(ByVal NewAlignNavBar As AlignNB)
mAlignNB = NewAlignNavBar
PropertyChanged "AlignNavBar"
UserControl_Resize
End Property

Public Property Get Appearance() As Appearance
Appearance = mAppearance
End Property
Public Property Let Appearance(ByVal NewAppearance As Appearance)
mAppearance = NewAppearance
PropertyChanged "Appearance"
picMain.Appearance = mAppearance
If Appearance = [pe3D] Then BorderStyle = [peFixed_Single]
End Property

Public Property Get BackColor() As OLE_COLOR
BackColor = picMain.BackColor
End Property
Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)
picMain.BackColor = NewBackColor
PropertyChanged "BackColor"
picFrame.BackColor = NewBackColor
picHolder.BackColor = NewBackColor
End Property

Public Property Get BorderStyle() As BorderStyle
BorderStyle = mBorderStyle
End Property
Public Property Let BorderStyle(ByVal NewBorderStyle As BorderStyle)
mBorderStyle = NewBorderStyle
picMain.BorderStyle = NewBorderStyle
PropertyChanged "BorderStyle"
End Property

Public Property Get LenzSize() As LenzSizeProps
LenzSize = LenSize
End Property
Public Property Let LenzSize(ByVal NewLenzSize As LenzSizeProps)
LenSize = NewLenzSize
PropertyChanged "LenzSize"
LenzSizer LenSize
End Property

Public Property Get Magnify() As MagnifyProperties
Magnify = MagnifySize
End Property
Public Property Let Magnify(ByVal NewMagnify As MagnifyProperties)
MagnifySize = NewMagnify
Combo3_SetText
PropertyChanged "Magnify"
If picZoom.Visible = True Then PaintZoomGlass
End Property

Public Property Get Picture() As Picture
Set Picture = picMain.Picture
End Property
'This property ables you to insert the controls picture to
'another picturebox outside the usercontrol. Elese you'll get
'an error
Public Property Let Picture(ByVal NewPicture As IPictureDisp)
Set Picture = NewPicture
End Property
Public Property Set Picture(ByVal NewPicture As Picture)
Set b = NewPicture
LockWindowUpdate picFrame.hwnd

' For main picture
PaintThePicture
PropertyChanged "Picture"
UserControl_Resize
If picZoom.Visible = True Then PaintZoomGlass

' For insert picture
With picInsert
    If picGate.bWidth Then
       .Left = picPosition.Width / 2 - .Width / 2
     Else
       .Left = 0
    End If
    If picGate.bHeight Then
       .Top = picPosition.Height / 2 - .Height / 2
     Else
       .Top = 0
    End If
End With
PaintThePictureCopy
PaintInsertGlass
' Check main pic position against insert position. Here we line them
' up cause left position is out by 1 (drop fractions in formula's)
picMain.Top = -((SizeDiff.sHeight / 100) * ((100 / (picPosition.Height - 2 - picInsert.Height)) * picInsert.Top))
picMain.Left = -((SizeDiff.sWidth / 100) * ((100 / (picPosition.Width - 2 - picInsert.Width)) * picInsert.Left)) + 1
HScroll1.Value = 2
LockWindowUpdate 0&
End Property

Public Property Get ScrollSpeed() As Integer
ScrollSpeed = mScrollSpeed
End Property
Public Property Let ScrollSpeed(ByVal NewScrollSpeed As Integer)
mScrollSpeed = NewScrollSpeed
PropertyChanged "ScrollSpeed"
    
'10 as minimum
If mScrollSpeed < 10 Then
    mScrollSpeed = 10
   '1000 as maximum
 ElseIf mScrollSpeed > 400 Then
    mScrollSpeed = 400
End If
Combo1.Text = mScrollSpeed
End Property

Public Property Get TipTextMagnifier() As String
TipTextMagnifier = picZoom.ToolTipText
End Property
Public Property Let TipTextMagnifier(ByVal NewTipTextMagnifier As String)
picZoom.ToolTipText = NewTipTextMagnifier
PropertyChanged "TipTextMagnifier"
End Property

Public Property Get TipTextPicView() As String
TipTextPicView = picMain.ToolTipText
End Property
Public Property Let TipTextPicView(ByVal NewTipTextPicView As String)
picMain.ToolTipText = NewTipTextPicView
PropertyChanged "TipTextPicView"
End Property

Private Sub UserControl_Initialize()
''''''''
End Sub

Private Sub UserControl_InitProperties()
AlignNavBar = mDefAlignNB
mScrollSpeed = 10
LenSize = peSmall
MagnifySize = [pe200%]
End Sub

Private Sub UserControl_Show()
Combo1.Text = mScrollSpeed
Combo2.Text = "Small"
Combo3.Text = "2.0x"
Combo3_SetText
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    Set picMain.Picture = .ReadProperty("Picture", Nothing)
    ScrollSpeed = .ReadProperty("ScrollSpeed", mDefScrollSpeed)
    BorderStyle = PropBag.ReadProperty("BorderStyle", mDefBorderStyle)
    Appearance = PropBag.ReadProperty("Appearance", mDefAppearance)
    BackColor = PropBag.ReadProperty("BackColor", UserControl.Parent.BackColor)
    picMain.ToolTipText = PropBag.ReadProperty("TipTextPicView", "")
    picZoom.ToolTipText = PropBag.ReadProperty("TipTextMagnifier", "")
    AlignNavBar = .ReadProperty("AlignNavBar", mDefAlignNB)
    LenzSize = PropBag.ReadProperty("LenzSize", mdefLS)
    Magnify = PropBag.ReadProperty("Magnify", mDefMag)
End With
End Sub

Private Sub UserControl_Resize()
Dim w As Integer
Dim h As Integer

On Error Resume Next

If picGate.bHeight Then picGate.bHeight = False
If picGate.bWidth Then picGate.bWidth = False

w = UserControl.ScaleWidth
h = UserControl.ScaleHeight

If w < SS Then
    UserControl.ScaleWidth = SS
    picFrame.Width = SS - 78
    picHolder.Width = SS
    w = SS
 ElseIf UserControl.ScaleHeight < SS Then
    UserControl.ScaleHeight = SS
    h = SS
End If

' Main frame
picFrame.Move 1, 1, w - 76, h - picHolder.Height - 3

' Main picture
Dim nL As Integer
Dim nT As Integer
' Center image if smaller than picFrame
With picMain
    If .Width < picFrame.Width Then
        nL = picFrame.Width / 2 - .Width / 2
        picGate.bWidth = True
       If .Height < picFrame.Height Then
          nT = picFrame.Height / 2 - .Height / 2
          picGate.bHeight = True
          .Move nL, nT
        Else
          .Move nL
       End If
     ElseIf .Height < picFrame.Height Then
       nT = picFrame.Height / 2 - .Height / 2
       picGate.bHeight = True
       .Top = nT
    End If
End With

'Controls
picTools.Move w - picTools.Width - 1, 1
' Control box holder
picHolder.Move 0, h - picHolder.Height - 1, w - 78

DoEvents
' Control box (inside control box holder)
With picHolderInsert
    Select Case mAlignNB
        Case peLeft
            .Left = 0
        Case peRight
            .Left = picHolder.Width - .Width
        Case peCenter
            .Left = picHolder.Width / 2 - .Width / 2
    End Select
End With

DoEvents
SizeDiff.sHeight = picMain.Height - picFrame.Height
SizeDiff.sWidth = picMain.Width - picFrame.Width
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "Picture", picMain.Picture, Nothing
    .WriteProperty "ScrollSpeed", mScrollSpeed, mDefScrollSpeed
    .WriteProperty "BorderStyle", mBorderStyle, mDefBorderStyle
    .WriteProperty "Appearance", mAppearance, mDefAppearance
    .WriteProperty "BackColor", picMain.BackColor, UserControl.Parent.BackColor
    .WriteProperty "TipTextPicView", picMain.ToolTipText, ""
    .WriteProperty "TipTextMagnifier", picZoom.ToolTipText, ""
    .WriteProperty "AlignNavBar", mAlignNB, mDefAlignNB
    .WriteProperty "LenzSize", LenSize, mdefLS
    .WriteProperty "Magnify", MagnifySize, mDefMag
End With
End Sub

Private Sub Combo1_Change()
Combo1_Click
End Sub
Private Sub Combo1_Click()
ScrollSpeed = Combo1.Text
End Sub

Private Sub Combo2_Change()
Combo2_Click
End Sub
Private Sub Combo2_Click()
Select Case Combo2.Text
   Case "Small"
      LenzSizer peSmall
   Case "Medium"
      LenzSizer peMedium
   Case "Large"
      LenzSizer peLarge
End Select
End Sub

Private Sub Combo3_Change()
Combo3_Click
End Sub
Private Sub Combo3_Click()
Select Case Combo3.Text
   Case "1.5x"
      MagnifySize = 1.5
   Case "2.0x"
      MagnifySize = 2
   Case "2.5x"
      MagnifySize = 2.5
   Case "3.0x"
      MagnifySize = 3
End Select
picZoom.ToolTipText = Combo3.Text & " magnification"
If picZoom.Visible = True Then PaintZoomGlass
End Sub
Private Sub Combo3_SetText()
With Combo3
    Select Case MagnifySize
       Case [pe150%]
          .Text = "1.5x"
       Case [pe200%]
          .Text = "2.0x"
       Case [pe250%]
          .Text = "2.5x"
       Case [pe250%]
          .Text = "3.0x"
    End Select
    picZoom.ToolTipText = .Text & " magnification"
End With
End Sub

Private Sub HScroll1_Change()
HScroll1_Click
End Sub
Private Sub HScroll1_Click()
If picMain.Picture = 0 Then Exit Sub
ReleaseCapture
DoEvents

Select Case HScroll1.Value
   Case Is = 0
      Label4.Caption = "50 %"
   Case Is = 1
      Label4.Caption = "75 %"
   Case Else
      Label4.Caption = (HScroll1.Value - 1) * 100 & " %"
End Select

LockWindowUpdate picFrame.hwnd
PaintThePicture
UserControl_Resize

picMain.ToolTipText = "Picture size: " & Label4.Caption

' If picutre size is less then picFrame size then center images
If picGate.bWidth Then
   picInsert.Left = picPosition.Width / 2 - picInsert.Width / 2
End If
If picGate.bHeight Then
   picInsert.Top = picPosition.Height / 2 - picInsert.Height / 2
End If
'reposition main picture
picMain.Top = -((SizeDiff.sHeight / 100) * ((100 / (picPosition.Height - 2 - picInsert.Height)) * picInsert.Top))
picMain.Left = -((SizeDiff.sWidth / 100) * ((100 / (picPosition.Width - 2 - picInsert.Width)) * picInsert.Left))
'reposition zoom lens
If picZoom.Visible = True Then
    picZoom.Visible = False
    PaintZoomGlass
End If
LockWindowUpdate 0&
End Sub

Private Sub LenzSizer(l As LenzSizeProps)
LenSize = l
With picZoom
   Select Case l
      Case peSmall
           .Width = 67
           .Height = 67
      Case peMedium
           .Width = 113
           .Height = 113
      Case peLarge
           .Width = 160
           .Height = 160
   End Select
   If .Visible = True Then PaintZoomGlass
End With
End Sub

Private Sub PaintThePicture()
Dim i As Single
Dim tH As Integer
Dim tW As Integer

On Error GoTo PaintErr

Select Case HScroll1.Value
   Case 0
      i = 0.5
   Case 1
      i = 0.75
   Case Else
      i = HScroll1.Value - 1
End Select

With picMain
    LockWindowUpdate picFrame.hwnd
    .Top = 0
    .Left = 0
    Set picMain.Picture = b
    .Width = .Width * i
    .Height = .Height * i
    bw = .ScaleX(b.Width, vbHimetric, .ScaleMode)
    bh = .ScaleY(b.Height, vbHimetric, .ScaleMode)
    picMain.PaintPicture b, 0, 0, .ScaleWidth, .ScaleHeight, 0, 0, bw, bh
    If .Visible = False Then .Visible = True
    tW = .ScaleX(b.Width, vbHimetric, .ScaleMode)
    tH = .ScaleY(b.Height, vbHimetric, .ScaleMode)
    picPosition.PaintPicture b, 0, 0, picPosition.ScaleWidth, picPosition.ScaleHeight, 0, 0, tW, tH
    If picInsert.Visible = False Then picInsert.Visible = True
End With
PaintErr:
LockWindowUpdate 0&
End Sub

Private Sub PaintThePictureCopy()
Dim tH As Integer
Dim tW As Integer

On Error GoTo PaintErr

With picPosition
    tW = .ScaleX(b.Width, vbHimetric, .ScaleMode)
    tH = .ScaleY(b.Height, vbHimetric, .ScaleMode)
    picPosition.PaintPicture b, 0, 0, picPosition.ScaleWidth, picPosition.ScaleHeight, 0, 0, tW, tH
    If picInsert.Visible = False Then picInsert.Visible = True
End With
PaintErr:
End Sub

Private Sub PaintZoomGlass()
Dim PercentPictureTop As Integer
Dim PercentPictureLeft As Integer
Dim ZoomPositionTop As Integer
Dim ZoomPositionLeft As Integer
Dim MagSize As Single
Dim ShowMsg As Boolean

On Error GoTo PaintErr

' Sanity checks to ensure zoom lens is smaller than picture size
If picZoom.Width > picMain.Width Then
   ShowMsg = True
End If
If picZoom.Height > picMain.Height Then
   ShowMsg = True
End If
If ShowMsg Then
   If LenSize = peMedium Or LenSize = peLarge Then
      MsgBox "Increase picture size," & vbCrLf & _
             "or decrease zoom lens size first!", vbExclamation + vbOKOnly, App.Title
    Else
      MsgBox "Increase picture size first!", vbExclamation + vbOKOnly, App.Title
   End If
   picZoom.Visible = False
   Exit Sub
End If

With picZoom
    ' Show zoom lens if not already visible
    If .Visible = False Then
        .Left = -picMain.Left + (picFrame.Width / 2 - .Width / 2)
        .Top = -picMain.Top + (picFrame.Height / 2 - .Height / 2)
        .Visible = True
    End If

    ' This will calculate the 'area' of the underlying picture (picMain)
    ' for the Zoom lenz to magnify.
    ' The formula works so the zoom lenz does not have to move anywhere
    ' offscreen or 'out of view' in order to zoom in on all of the underlying picture.
    ' Diagram demonstrates postion of 'area'
     
    '   '--'--'------------------------'--'--'
    '   'area '  b  '            '  b  ' area'
    '   '-----'  o  '            '  o  '-----'
    '   ' zoom   x  '            '  x  zoom  '
    '   '-----------'            '-----------'
    '   '           '-----------'            '
    '   '           'z '-----' b'            '
    '   '           'o 'Area!' o'            '
    '   '           'o '-----' x'            '
    '   '           '-----------'            '
    '   '-----------'            '-----------'
    '   ' zoom   b  '            '  b  zoom  '
    '   '-----'  o  '            '  o  '-----'
    '   'area '  x  '            '  x  ' area'
    '   '--'--'------------------------'--'--'

    ' Get picture positions for StretchBlt routine
    MagSize = .Height / MagnifySize
    PercentPictureTop = (100 / (picMain.Height - .Height)) * .Top
    ZoomPositionTop = .Top + (((.Height - MagSize) / 100) * PercentPictureTop)
    PercentPictureLeft = (100 / (picMain.Width - picZoom.Width)) * .Left
    ZoomPositionLeft = .Left + (((.Width - MagSize) / 100) * PercentPictureLeft)

    StretchBlt .hdc, _
               0, 0, .Width, .Height, _
               picMain.hdc, _
               ZoomPositionLeft, ZoomPositionTop, _
               .Width / MagnifySize, .Height / MagnifySize, SrcCopy
    .Refresh
End With
PaintErr:
End Sub

Private Sub PaintInsertGlass()
On Error GoTo PaintErr

With picInsert
    StretchBlt .hdc, _
            0, 0, .Width, .Height, _
            picPosition.hdc, _
            .Left, .Top, _
            .Width, .Height, SrcCopy
    .Refresh
End With
PaintErr:
End Sub

Private Sub picButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
picButton(Index).Picture = picDown(Index).Picture
' Sanity checks on timers
tmrZoom.Enabled = False
tmrInsert.Enabled = False

Select Case picButton(Index).Index
   Case 5
      tmrScrollLeft.Enabled = True
   Case 6
      tmrScrollRight.Enabled = True
   Case 7
      tmrScrollUp.Enabled = True
   Case 8
      tmrScrollDown.Enabled = True
End Select
End Sub

Private Sub picButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim pos As Size

picButton(Index).Picture = picUp(Index).Picture
    
Select Case picButton(Index).Index
   Case 0
      If picZoom.Visible = True Then
         If Not picGate.bHeight Then
            If Not picMainpos.pB Then picZoom.Top = picZoom.Top + (SizeDiff.sHeight + picMain.Top) '- picFrame.Height ' SizeDiff.sHeight
            picMainpos.pB = True
            picMainpos.pt = False
         End If
         If Not picGate.bWidth Then
            If Not picMainpos.pL Then picZoom.Left = picZoom.Left + picMain.Left ' picmain is negative
            picMainpos.pL = True
            picMainpos.pR = False
         End If
      End If
      If Not picGate.bWidth Then picMain.Left = 0
      If Not picGate.bHeight Then picMain.Top = -SizeDiff.sHeight
   
   Case 1
      If picZoom.Visible = True Then
         If Not picGate.bHeight Then
            If Not picMainpos.pB Then picZoom.Top = picZoom.Top + (SizeDiff.sHeight + picMain.Top)
            picMainpos.pB = True
            picMainpos.pt = False
         End If
         If Not picGate.bWidth Then
            If Not picMainpos.pR Then picZoom.Left = picZoom.Left + (SizeDiff.sWidth + picMain.Left)
            picMainpos.pR = True
            picMainpos.pL = False
         End If
      End If
      If Not picGate.bHeight Then picMain.Top = -SizeDiff.sHeight
      If Not picGate.bWidth Then picMain.Left = -SizeDiff.sWidth
   
   Case 2
      If picZoom.Visible = True Then
         If Not picGate.bHeight Then
            picZoom.Top = SizeDiff.sHeight / 2 + (picZoom.Top + picMain.Top)
            picMainpos.pt = False
            picMainpos.pB = False
         End If
         If Not picGate.bWidth Then
            picZoom.Left = SizeDiff.sWidth / 2 + (picZoom.Left + picMain.Left)
            picMainpos.pR = False
            picMainpos.pL = False
         End If
      End If
      If Not picGate.bHeight Then picMain.Top = -SizeDiff.sHeight / 2
      If Not picGate.bWidth Then picMain.Left = -SizeDiff.sWidth / 2
   
   Case 3
      If picZoom.Visible = True Then
         If Not picGate.bHeight Then
            If Not picMainpos.pt Then picZoom.Top = picZoom.Top + picMain.Top 'picmain is negative
            picMainpos.pt = True
            picMainpos.pB = False
         End If
         If Not picGate.bWidth Then
            If Not picMainpos.pL Then picZoom.Left = picZoom.Left + picMain.Left ' picmain is negative
            picMainpos.pL = True
            picMainpos.pR = False
         End If
      End If
      If Not picGate.bWidth Then picMain.Left = 0
      If Not picGate.bHeight Then picMain.Top = 0
   
   Case 4
      If picZoom.Visible = True Then
         If Not picGate.bHeight Then
            If Not picMainpos.pt Then picZoom.Top = picZoom.Top + picMain.Top 'picmain is negative
            picMainpos.pt = True
            picMainpos.pB = False
         End If
         If Not picGate.bWidth Then
            If Not picMainpos.pR Then picZoom.Left = picZoom.Left + (SizeDiff.sWidth + picMain.Left)
            picMainpos.pR = True
            picMainpos.pL = False
         End If
      End If
      If Not picGate.bHeight Then picMain.Top = 0
      If Not picGate.bWidth Then picMain.Left = -SizeDiff.sWidth
   
   Case 5
      tmrScrollLeft.Enabled = False
   Case 6
      tmrScrollRight.Enabled = False
   Case 7
      tmrScrollUp.Enabled = False
   Case 8
      tmrScrollDown.Enabled = False
   Case 9
      If picMain.Picture = 0 Then Exit Sub
      If picZoom.Visible = True Then picZoom.Visible = False
      PaintZoomGlass
End Select

If picZoom.Visible = True Then PaintZoomGlass
PositionInsertGlass
End Sub

Private Sub picInsert_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    tmrInsert.Enabled = True
End If
End Sub

Private Sub picInsert_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    ReleaseCapture  ' <--- Kills the MouseUp Event so no need for MouseUp
    SendMessage picInsert.hwnd, &HA1, 2, 0&
End If
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Sanity checks
tmrZoom.Enabled = False
tmrInsert.Enabled = False
End Sub

Private Sub picPosition_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tmrInsert.Enabled = False
End Sub

Private Sub picZoom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    tmrZoom.Enabled = True
    DoEvents
 Else
    picZoom.Visible = False
    tmrZoom.Enabled = False
End If
End Sub

Private Sub picZoom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    ReleaseCapture  ' <--- Kills the MouseUp Event so no need for MouseUp
    SendMessage picZoom.hwnd, &HA1, 2, 0&
End If
End Sub

Private Sub PositionInsertGlass()
Dim PercentPictureTop As Integer
Dim PercentPictureLeft As Integer

PercentPictureTop = (100 / SizeDiff.sHeight) * -picMain.Top
PercentPictureLeft = (100 / SizeDiff.sWidth) * -picMain.Left

With picInsert
   .Top = ((picPosition.Height - 2 - .Height) / 100) * PercentPictureTop
   .Left = ((picPosition.Width - 2 - .Width) / 100) * PercentPictureLeft
End With
End Sub

Private Sub tmrScrollLeft_Timer()
If picGate.bWidth Then Exit Sub
If picMain.Left = 0 Then Exit Sub
With picMain
    .Left = .Left + mScrollSpeed
    picZoom.Left = picZoom.Left - mScrollSpeed
    If .Left > 0 Then
       .Left = 0
       tmrScrollDown.Enabled = False
    End If
End With
If picZoom.Visible = True Then PaintZoomGlass
PositionInsertGlass
End Sub

Private Sub tmrScrollRight_Timer()
If picGate.bWidth Then Exit Sub
If picMain.Left = -SizeDiff.sWidth Then Exit Sub
With picMain
    .Left = .Left - mScrollSpeed
    picZoom.Left = picZoom.Left + mScrollSpeed
    If .Left < -SizeDiff.sWidth Then
       .Left = -SizeDiff.sWidth
       tmrScrollLeft.Enabled = False
    End If
End With
If picZoom.Visible = True Then PaintZoomGlass
PositionInsertGlass
End Sub

Private Sub tmrScrollUp_Timer()
If picGate.bHeight Then Exit Sub
If picMain.Top = 0 Then Exit Sub
With picMain
    .Top = .Top + mScrollSpeed
    picZoom.Top = picZoom.Top - mScrollSpeed
    If .Top > 0 Then
       .Top = 0
       tmrScrollDown.Enabled = False
    End If
End With
If picZoom.Visible = True Then PaintZoomGlass
PositionInsertGlass
End Sub

Private Sub tmrScrollDown_Timer()
If picGate.bHeight Then Exit Sub
If picMain.Top = -SizeDiff.sHeight Then Exit Sub
With picMain
    .Top = .Top - mScrollSpeed
    picZoom.Top = picZoom.Top + mScrollSpeed
    If .Top < -SizeDiff.sHeight Then
       .Top = -SizeDiff.sHeight
       tmrScrollUp.Enabled = False
    End If
End With
If picZoom.Visible = True Then PaintZoomGlass
PositionInsertGlass
End Sub

Private Sub tmrZoom_Timer()
Dim cursor As POINTAPI

GetCursorPos cursor
With picZoom
    Select Case .Top
        Case Is < 0
            cursor.Y = cursor.Y + -.Top
            SetCursorPos cursor.X, cursor.Y
        Case Is > picMain.Height - .Height
            cursor.Y = cursor.Y - (.Top - (picMain.Height - .Height))
            SetCursorPos cursor.X, cursor.Y
    End Select
    Select Case .Left
       Case Is < 0
          cursor.X = cursor.X + -.Left
          SetCursorPos cursor.X, cursor.Y
       Case Is > picMain.Width - .Width
          cursor.X = cursor.X - (.Left - (picMain.Width - .Width))
          SetCursorPos cursor.X, cursor.Y
    End Select
    DoEvents
    ' Sanity checks
    ' NOTE:
    ' By restting the cursor position this automatically repositions the picZoom.PictureBox
    ' If the cursor repositions followed by the mouse being released before the
    ' picturebox repositions itself under the cursor, the picturebox is left behind or
    ' 'partially offscreen'. Thus the cursor runs away...
    ' By checking the picturebox has repositioned here after SetCursorPos and DoEvents
    ' we can be sure that the picturebox is repositioned as well as the cursor!!!!
    Select Case .Top
        Case Is < 0
            .Top = .Top + -.Top
        Case Is > picMain.Height - .Height
            .Top = .Top - (.Top - (picMain.Height - .Height))
    End Select
    Select Case .Left
        Case Is < 0
          .Left = .Left + -.Left
        Case Is > picMain.Width - .Width
          .Left = .Left - (.Left - (picMain.Width - .Width))
    End Select
End With
PaintZoomGlass
End Sub

Private Sub tmrInsert_Timer()
Dim cursor As POINTAPI
Dim PercentPictureTop As Single
Dim PercentPictureLeft As Single

GetCursorPos cursor
With picInsert
    Select Case .Top
        Case Is < 0
            cursor.Y = cursor.Y + -.Top
            SetCursorPos cursor.X, cursor.Y
        Case Is > picPosition.Height - .Height - 2
            cursor.Y = cursor.Y - (.Top - (picPosition.Height - .Height - 2))
            SetCursorPos cursor.X, cursor.Y
    End Select
    Select Case .Left
       Case Is < 0
          cursor.X = cursor.X + -.Left
          SetCursorPos cursor.X, cursor.Y
       Case Is > picPosition.Width - .Width - 2
          cursor.X = cursor.X - (.Left - (picPosition.Width - .Width - 2))
          SetCursorPos cursor.X, cursor.Y
    End Select
    DoEvents
    ' Sanity checks
    ' NOTE:
    ' By restting the cursor position this automatically repositions the picZoom.PictureBox
    ' If the cursor repositions followed by the mouse being released before the
    ' picturebox repositions itself under the cursor, the picturebox is left behind or
    ' 'partially offscreen'. Thus the cursor runs away...
    ' By checking the picturebox has repositioned here after SetCursorPos and DoEvents
    ' we can be sure that the picturebox is repositioned as well as the cursor!!!!
    Select Case .Top
        Case Is < 0
            .Top = .Top + -.Top
        Case Is > picPosition.Height - .Height - 2
            .Top = .Top - (.Top - (picPosition.Height - .Height - 2))
    End Select
    Select Case .Left
        Case Is < 0
          .Left = .Left + -.Left
        Case Is > picPosition.Width - .Width - 2
          .Left = .Left - (.Left - (picPosition.Width - .Width - 2))
    End Select
    PaintInsertGlass
    ' reposition main picture
    PercentPictureTop = (100 / (picPosition.Height - 2 - .Height)) * .Top
    PercentPictureLeft = (100 / (picPosition.Width - 2 - .Width)) * .Left
End With

picMain.Top = -((SizeDiff.sHeight / 100) * PercentPictureTop)
picMain.Left = -((SizeDiff.sWidth / 100) * PercentPictureLeft)
End Sub
