VERSION 5.00
Begin VB.Form frmFonts 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   207
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   165
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox ListFonts 
      Height          =   255
      Left            =   180
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   2700
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.PictureBox FontBg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   60
      ScaleHeight     =   167
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   0
      Top             =   60
      Width           =   1995
      Begin VB.VScrollBar VScroll1 
         Height          =   1935
         Left            =   1680
         Max             =   0
         TabIndex        =   1
         Top             =   120
         Width           =   255
      End
      Begin VB.PictureBox FontScroll 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1275
         Left            =   0
         ScaleHeight     =   85
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   125
         TabIndex        =   2
         Top             =   0
         Width           =   1875
         Begin VB.Label LblSplit 
            BackColor       =   &H80000008&
            Caption         =   "Label1"
            Height          =   15
            Index           =   0
            Left            =   240
            TabIndex        =   7
            Top             =   780
            Width           =   1275
         End
         Begin VB.Label lblFont 
            BackColor       =   &H80000005&
            Caption         =   "Font"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   4
            Top             =   180
            Width           =   1635
         End
         Begin VB.Label lblFontName 
            BackColor       =   &H80000005&
            Caption         =   " Fontname"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   1455
         End
      End
   End
   Begin VB.Label lblFontTest 
      AutoSize        =   -1  'True
      Caption         =   "gGpOQsTl|"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      TabIndex        =   5
      Top             =   2700
      Visible         =   0   'False
      Width           =   1200
   End
End
Attribute VB_Name = "frmFonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CurrentFont As Integer

Private Sub Form_Load()
        Dim font As Variant
    Dim i As Integer
    Dim MaxHeight As Integer
    
    Me.Hide
    DoEvents
    
    CurrentFont = -1
    
    Me.FontBg.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    Me.VScroll1.Move Me.FontBg.ScaleWidth - Me.VScroll1.Width, 0, Me.VScroll1.Width, Me.FontBg.ScaleHeight
    Me.FontScroll.Width = Me.FontBg.ScaleWidth
    
    For i = 0 To Screen.FontCount - 1
        Me.ListFonts.AddItem Screen.Fonts(i)
    Next i
    
    For i = 0 To Me.ListFonts.ListCount - 1
        If i > 0 Then
            Load Me.lblFontName(i)
            Load Me.LblFont(i)
            Load Me.LblSplit(i)
        Else
            frmControls.LblFont.Caption = " " & Me.ListFonts.List(i)
        End If
        
        Me.lblFontName(i).Caption = " " & Me.ListFonts.List(i)
        Me.LblFont(i).Caption = " " & Me.ListFonts.List(i)
        Me.LblFont(i).font = Me.ListFonts.List(i)
        Me.lblFontTest.font = Me.ListFonts.List(i)
        Me.LblFont(i).FontSize = 12
        Me.lblFontTest.FontSize = 12
        Me.LblFont(i).Height = Me.lblFontTest.Height
        
        If i = 0 Then
            Me.lblFontName(i).Move 0, 0, Me.FontScroll.ScaleWidth
            Me.LblFont(i).Move 0, Me.lblFontName(i).Height, Me.FontScroll.ScaleWidth
            Me.LblSplit(i).Move 0, Me.LblFont(i).Top + Me.LblFont(i).Height, Me.FontScroll.ScaleWidth, 1
        Else
            Me.lblFontName(i).Move 0, Me.LblFont(i - 1).Top + Me.LblFont(i - 1).Height + 1, Me.FontScroll.ScaleWidth
            Me.LblFont(i).Move 0, Me.lblFontName(i).Height + Me.lblFontName(i).Top, Me.FontScroll.ScaleWidth
            Me.LblSplit(i).Move 0, Me.LblFont(i).Top + Me.LblFont(i).Height, Me.FontScroll.ScaleWidth, 1
            
            MaxHeight = Me.LblFont(i).Top + Me.LblFont(i).Height
            Me.FontScroll.Height = MaxHeight
        End If
        
        Me.LblFont(i).Visible = True
        Me.LblFont(i).ZOrder 0
        
        Me.lblFontName(i).Visible = True
        Me.lblFontName(i).ZOrder 0
        
        Me.LblSplit(i).Visible = True
        Me.LblSplit(i).ZOrder 0
    Next i
    
    If MaxHeight > Me.FontBg.Height Then
        Me.VScroll1.Max = Me.FontScroll.Height - Me.FontBg.ScaleHeight
        Me.VScroll1.SmallChange = Me.FontScroll.Height / i
        Me.VScroll1.LargeChange = Me.FontScroll.Height - (((Me.FontScroll.Height - Me.FontBg.ScaleHeight) / Me.FontScroll.Height) * Me.FontScroll.Height)
    
    Else
        Me.VScroll1.Visible = False
        Me.FontBg.Height = MaxHeight
        Me.Height = MaxHeight
    End If
    
End Sub

Private Sub LblFont_Click(Index As Integer)
    SelectFont Index
End Sub

Private Sub lblFont_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    MoveOverFont Index
    
End Sub

Private Sub MoveOverFont(Index As Integer)
    If CurrentFont <> Index Then
        If CurrentFont >= 0 Then
            Me.LblFont(CurrentFont).BackColor = vbWindowBackground
            Me.LblFont(CurrentFont).ForeColor = vbWindowText
            Me.lblFontName(CurrentFont).BackColor = vbWindowBackground
            Me.lblFontName(CurrentFont).ForeColor = vbWindowText
        End If
        Me.LblFont(Index).BackColor = vbActiveTitleBar
        Me.LblFont(Index).ForeColor = vbActiveTitleBarText
        Me.lblFontName(Index).BackColor = vbActiveTitleBar
        Me.lblFontName(Index).ForeColor = vbActiveTitleBarText
        CurrentFont = Index
    End If
End Sub

Private Sub SelectFont(Index As Integer)
    frmControls.LblFont.Caption = Me.lblFontName(Index).Caption
    frmControls.ShowHideFonts
    
    On Error Resume Next
    frmMain.ActiveForm.Buffer.font = Mid(frmControls.LblFont.Caption, 2)
    frmMain.ActiveForm.TextInput.font = Mid(frmControls.LblFont.Caption, 2)
    frmMain.ActiveForm.lblTextSize.font = Mid(frmControls.LblFont.Caption, 2)
    frmMain.ActiveForm.TextInput.Width = frmMain.ActiveForm.lblTextSize.Width
    frmMain.ActiveForm.TextInput.Height = frmMain.ActiveForm.lblTextSize.Height
    
End Sub

Private Sub lblFontName_Click(Index As Integer)
    SelectFont Index
    
End Sub

Private Sub lblFontName_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    MoveOverFont Index
End Sub

Private Sub VScroll1_Change()
    Me.FontScroll.SetFocus
    Me.FontScroll.Top = -Me.VScroll1.Value
End Sub

Private Sub VScroll1_GotFocus()
    Me.FontScroll.SetFocus
End Sub

Private Sub VScroll1_Scroll()
    Me.FontScroll.SetFocus
    Me.FontScroll.Top = -Me.VScroll1.Value
End Sub
