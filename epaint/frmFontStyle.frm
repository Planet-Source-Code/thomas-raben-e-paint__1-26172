VERSION 5.00
Begin VB.Form frmFontStyle 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   50
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   113
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox StyleBg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   111
      TabIndex        =   0
      Top             =   0
      Width           =   1695
      Begin VB.Label lblStyle 
         BackColor       =   &H80000005&
         Caption         =   " Bold Italic"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   0
         TabIndex        =   4
         Top             =   540
         Width           =   1695
      End
      Begin VB.Label lblStyle 
         BackColor       =   &H80000005&
         Caption         =   " Italic"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblStyle 
         BackColor       =   &H80000005&
         Caption         =   " Bold"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   2
         Top             =   180
         Width           =   1695
      End
      Begin VB.Label lblStyle 
         BackColor       =   &H80000005&
         Caption         =   " Normal"
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmFontStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CurrentStyle As Integer

Private Sub Form_Load()
    Me.Hide
    DoEvents
    frmControls.lblFontStyle.Caption = Me.lblStyle(0).Caption
    
End Sub

Private Sub lblStyle_Click(Index As Integer)
    frmControls.ShowHideFontStyle
    frmControls.lblFontStyle.Caption = Me.lblStyle(Index).Caption
    frmControls.lblFontStyle.FontBold = Me.lblStyle(Index).FontBold
    frmControls.lblFontStyle.FontItalic = Me.lblStyle(Index).FontItalic
    
    On Error Resume Next
    frmMain.ActiveForm.Buffer.FontBold = Me.lblStyle(Index).FontBold
    frmMain.ActiveForm.TextInput.FontBold = Me.lblStyle(Index).FontBold
    frmMain.ActiveForm.lblTextSize.FontBold = Me.lblStyle(Index).FontBold
    frmMain.ActiveForm.Buffer.FontItalic = Me.lblStyle(Index).FontItalic
    frmMain.ActiveForm.TextInput.FontItalic = Me.lblStyle(Index).FontItalic
    frmMain.ActiveForm.lblTextSize.FontItalic = Me.lblStyle(Index).FontItalic
    frmMain.ActiveForm.TextInput.Width = frmMain.ActiveForm.lblTextSize.Width
    frmMain.ActiveForm.TextInput.Height = frmMain.ActiveForm.lblTextSize.Height
End Sub

Private Sub lblStyle_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If CurrentStyle <> Index Then
        Me.lblStyle(CurrentStyle).BackColor = vbWindowBackground
        Me.lblStyle(CurrentStyle).ForeColor = vbWindowText
        Me.lblStyle(Index).BackColor = vbActiveTitleBar
        Me.lblStyle(Index).ForeColor = vbActiveTitleBarText
        CurrentStyle = Index
    End If
End Sub
