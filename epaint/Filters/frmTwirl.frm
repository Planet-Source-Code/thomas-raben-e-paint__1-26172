VERSION 5.00
Begin VB.Form frmTwirl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Twirl..."
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   3420
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Source 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   2400
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   9
      Top             =   1260
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Dest 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   2400
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox ProgressB 
      Height          =   255
      Left            =   60
      ScaleHeight     =   195
      ScaleWidth      =   2175
      TabIndex        =   6
      Top             =   2340
      Visible         =   0   'False
      Width           =   2235
      Begin VB.PictureBox Progress 
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   0
         ScaleHeight     =   195
         ScaleWidth      =   375
         TabIndex        =   7
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   315
      Index           =   1
      Left            =   2460
      TabIndex        =   5
      Top             =   420
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   315
      Index           =   0
      Left            =   2460
      TabIndex        =   4
      Top             =   60
      Width           =   915
   End
   Begin VB.PictureBox TwirlPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   2235
      Left            =   60
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   1
      Top             =   60
      Width           =   2235
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   50
      Left            =   60
      Max             =   200
      TabIndex        =   0
      Top             =   2340
      Value           =   100
      Width           =   2235
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Twirl Value:"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   2700
      Width           =   1215
   End
End
Attribute VB_Name = "frmTwirl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Dim Parms As String

Private Sub Command1_Click(Index As Integer)
    If Index = 0 Then
        Render_Twirl Me.HScroll1.Value - (Me.HScroll1.Max / 2)
    End If
    
    Unload Me
    End
    
End Sub

Private Sub Form_Load()
    Parms = Command$
    If Parms = "" Then End
    
    Me.Source.Picture = LoadPicture(Parms)
    Me.Dest.Picture = LoadPicture(Parms)
    
    Draw_Twirl Me.HScroll1.Value - (Me.HScroll1.Max / 2)
    Me.Label1.Caption = Me.HScroll1.Value - (Me.HScroll1.Max / 2)
    
    
    
End Sub

Private Sub HScroll1_Change()
    Draw_Twirl Me.HScroll1.Value - (Me.HScroll1.Max / 2)
    Me.Label1.Caption = Me.HScroll1.Value - (Me.HScroll1.Max / 2)
    Me.TwirlPic.SetFocus
End Sub

Private Sub HScroll1_Scroll()
    Draw_Twirl Me.HScroll1.Value - (Me.HScroll1.Max / 2)
    Me.Label1.Caption = Me.HScroll1.Value - (Me.HScroll1.Max / 2)
    Me.TwirlPic.SetFocus
End Sub


'DRAW TWIRL
Public Sub Draw_Twirl(Angle As Double)
    Dim Rad As Double
    Dim A As Double
    Dim B As Double
    Dim x As Double
    Dim y As Double
    
    x = Me.TwirlPic.ScaleWidth / 2
    y = Me.TwirlPic.ScaleHeight / 2
    
    B = Angle / 10000
    
    Me.TwirlPic.Cls
    
    For Rad = 100 To 0 Step -0.1
        A = A + B
        SetPixelV Me.TwirlPic.hdc, x + Cos(A) * Rad, y + Sin(A) * Rad, 0
        SetPixelV Me.TwirlPic.hdc, x - Cos(A) * Rad, y - Sin(A) * Rad, 0
        SetPixelV Me.TwirlPic.hdc, x - Sin(A) * Rad, y + Cos(A) * Rad, 0
        SetPixelV Me.TwirlPic.hdc, x + Sin(A) * Rad, y - Cos(A) * Rad, 0
    Next Rad
    
End Sub

Public Sub Render_Twirl(Angle As Double)
    Dim Rad As Double
    Dim A As Double
    Dim B As Double
    Dim x As Double
    Dim y As Double
    Dim R As Double
    Dim C As Long
    Dim W As Double
    Dim Max As Integer
    
    Dim OS_Y As Integer
    
    Dim done() As Boolean
    
    ReDim done(Me.Source.ScaleWidth, Me.Source.ScaleHeight)
    
    On Error Resume Next
    
    Const PI = 3.1415
    
    x = Me.Source.ScaleWidth / 2
    y = Me.Source.ScaleHeight / 2
    
    If y > x Then
        Max = y - x
    Else
        Max = 0
    End If
    
    If Me.Source.ScaleHeight > Me.Source.ScaleWidth Then
        B = Angle / ((Me.Source.ScaleHeight / 2) * 100)
        W = (Me.Source.ScaleHeight / 2)
    Else
        B = Angle / ((Me.Source.ScaleWidth / 2) * 100)
        W = (Me.Source.ScaleWidth / 2)
    End If
    
    Me.Enabled = False
    Me.HScroll1.Enabled = False
    Me.Command1(0).Enabled = False
    Me.Command1(1).Enabled = False
    Me.ProgressB.Visible = True
    Me.Progress.Width = 1
    
    For Rad = y - Max To 0 Step -0.1
        A = A + B
        For R = 0 To PI * 2 Step W / (W * 100)
            C = GetPixel(Me.Source.hdc, (x + Cos(R) * Rad), (y + Sin(R) * Rad))
            If done((x + Cos(A + R) * Rad), (y + Sin(A + R) * Rad)) = False Then
                SetPixelV Me.Dest.hdc, (x + Cos(A + R) * Rad), (y + Sin(A + R) * Rad), C
                done((x + Cos(A + R) * Rad), (y + Sin(A + R) * Rad)) = True
            End If
        Next R
        Me.Progress.Width = Me.ProgressB.ScaleWidth / 100 * (100 - (Rad / y * 100))
        DoEvents
    Next Rad
    
    Me.Dest.Refresh
    
    Me.Enabled = True
    Me.HScroll1.Enabled = True
    Me.Command1(0).Enabled = True
    Me.Command1(1).Enabled = True
    Me.ProgressB.Visible = False
    
    Me.Source.Refresh

    SavePicture Me.Dest.Image, Parms
End Sub
