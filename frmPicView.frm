VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPicView 
   Caption         =   "Image Viewer"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7050
   Icon            =   "frmPicView.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   324
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   470
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   3600
      Width           =   975
   End
   Begin Project1.PicControl PC1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8281
      ScrollSpeed     =   10
      Appearance      =   0
      TipTextPicView  =   "Picture size: 100%"
      TipTextMagnifier=   "1.5x magnification"
      Magnify         =   1.5
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmPicView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
On Error Resume Next

CommonDialog1.Filter = "*.jpg*.bmp*.gif"
CommonDialog1.ShowOpen
    
PC1.Picture = LoadPicture(CommonDialog1.FileName)
End Sub

Private Sub Form_Resize()
On Error Resume Next

PC1.Width = Me.ScaleWidth - 15
PC1.Height = Me.ScaleHeight - 15
Command1.Left = Me.ScaleWidth - Command1.Width - 13
End Sub
