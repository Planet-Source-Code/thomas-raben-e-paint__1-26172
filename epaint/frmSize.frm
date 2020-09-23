VERSION 5.00
Begin VB.Form frmSize 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image Size..."
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   97
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   276
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Constrain Proportions"
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   1140
      Value           =   1  'Checked
      Width           =   2835
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   315
      Index           =   0
      Left            =   2940
      TabIndex        =   6
      Top             =   360
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   315
      Index           =   1
      Left            =   2940
      TabIndex        =   5
      Top             =   0
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Caption         =   "Image Size"
      Height          =   1035
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      Begin VB.PictureBox Constrain 
         BorderStyle     =   0  'None
         Height          =   675
         Left            =   1800
         Picture         =   "frmSize.frx":0000
         ScaleHeight     =   45
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   45
         TabIndex        =   8
         Top             =   240
         Width           =   675
      End
      Begin VB.TextBox txtWidth 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   780
         TabIndex        =   2
         Text            =   "400"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtHeight 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   780
         TabIndex        =   1
         Text            =   "300"
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Width:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Top             =   300
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Height:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   3
         Top             =   660
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    Dim W As Integer
    
    If Me.Check1.Value = 1 Then
        Me.Constrain.Visible = True
        
        If Me.txtWidth.Text = "" Then
            W = 0
        Else
            W = Me.txtWidth.Text
        End If
        If Me.Check1.Value = 1 Then Me.txtHeight.Text = CInt(frmMain.ActiveForm.Buffer.ScaleHeight * (CInt(W) / frmMain.ActiveForm.Buffer.ScaleWidth))
    Else
        Me.Constrain.Visible = False
    End If
    
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim OW As Integer, OH As Integer
    
    If Index = 1 Then
        OW = frmMain.ActiveForm.Buffer.Width
        OH = frmMain.ActiveForm.Buffer.Height
        frmMain.ActiveForm.Buffer.Width = Me.txtWidth + 2
        frmMain.ActiveForm.Buffer.Height = Me.txtHeight + 2
        
        StretchBlt frmMain.ActiveForm.Buffer.hdc, 0, 0, frmMain.ActiveForm.Buffer.ScaleWidth, frmMain.ActiveForm.Buffer.ScaleHeight, frmMain.ActiveForm.Buffer.hdc, 0, 0, OW, OH, vbSrcCopy
        
        frmMain.ActiveForm.PaintArea.Width = (frmMain.ActiveForm.Buffer.ScaleWidth * (frmMain.ActiveForm.GetZoomFactor / 100)) + 2
        frmMain.ActiveForm.PaintArea.Height = (frmMain.ActiveForm.Buffer.ScaleHeight * (frmMain.ActiveForm.GetZoomFactor / 100)) + 2
        
        UpdateArea frmMain.ActiveForm.Buffer, 0, 0, frmMain.ActiveForm.GetZoomFactor
    End If
    
    Unload Me

    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    frmMain.Enabled = True
    frmMain.SetFocus
    Cancel = True
    
End Sub

Private Sub txtHeight_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim H As Integer
    If Me.txtHeight.Text = "" Then
        H = 0
    Else
        H = Me.txtHeight.Text
    End If
    
    If Me.Check1.Value = 1 Then Me.txtWidth.Text = CInt(frmMain.ActiveForm.Buffer.ScaleWidth * (CInt(H) / frmMain.ActiveForm.Buffer.ScaleHeight))
End Sub

Private Sub txtWidth_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim W As Integer
    If Me.txtWidth.Text = "" Then
        W = 0
    Else
        W = Me.txtWidth.Text
    End If
    
    If Me.Check1.Value = 1 Then Me.txtHeight.Text = CInt(frmMain.ActiveForm.Buffer.ScaleHeight * (CInt(W) / frmMain.ActiveForm.Buffer.ScaleWidth))
End Sub
