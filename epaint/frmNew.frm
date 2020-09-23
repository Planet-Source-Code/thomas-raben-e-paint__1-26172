VERSION 5.00
Begin VB.Form frmNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Image"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   203
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   331
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   315
      Index           =   1
      Left            =   3780
      TabIndex        =   12
      Top             =   2340
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   315
      Index           =   0
      Left            =   3780
      TabIndex        =   11
      Top             =   2700
      Width           =   1155
   End
   Begin VB.Frame Frame2 
      Caption         =   "Background"
      Height          =   1335
      Left            =   60
      TabIndex        =   7
      Top             =   1680
      Width           =   3615
      Begin VB.OptionButton Option1 
         Caption         =   "Background"
         Height          =   315
         Index           =   2
         Left            =   180
         TabIndex        =   10
         Top             =   900
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Foreground"
         Height          =   315
         Index           =   1
         Left            =   180
         TabIndex        =   9
         Top             =   600
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         Caption         =   "White"
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   300
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Image"
      Height          =   1155
      Left            =   60
      TabIndex        =   2
      Top             =   480
      Width           =   3615
      Begin VB.TextBox txtHeight 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   780
         TabIndex        =   6
         Text            =   "300"
         Top             =   660
         Width           =   975
      End
      Begin VB.TextBox txtWidth 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   780
         TabIndex        =   5
         Text            =   "400"
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Height:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   4
         Top             =   720
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Width:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   720
      TabIndex        =   1
      Text            =   "Untitled"
      Top             =   60
      Width           =   2955
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
    frmMain.Enabled = True
    
    If Index = 1 Then
        Dim f As New frmPaint
        f.PaintArea.Width = Me.txtWidth.Text + 2
        f.PaintArea.Height = Me.txtHeight.Text + 2
        
        f.Buffer.Width = Me.txtWidth.Text + 2
        f.Buffer.Height = Me.txtHeight.Text + 2
        
        f.Caption = Me.txtName.Text & " - 100%"
        f.Tag = Me.txtName.Text
        f.Show
        
        If Me.Option1(0).Value = True Then
            f.PaintArea.BackColor = vbWhite
            f.Buffer.BackColor = vbWhite
        ElseIf Me.Option1(1).Value = True Then
            f.PaintArea.BackColor = frmMain.SelColor(0).BackColor
            f.Buffer.BackColor = frmMain.SelColor(0).BackColor
        Else
            f.PaintArea.BackColor = frmMain.SelColor(1).BackColor
            f.Buffer.BackColor = frmMain.SelColor(1).BackColor
        End If
        
        frmMain.LeftBar.SetFocus
    End If
    
    Me.Hide
End Sub

Private Sub Form_Load()
    Me.Option1(0).Value = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    Cancel = True
    
End Sub

