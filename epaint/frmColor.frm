VERSION 5.00
Begin VB.Form frmColor 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color Picker"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   284
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   369
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   915
      Left            =   4440
      ScaleHeight     =   855
      ScaleWidth      =   975
      TabIndex        =   5
      Top             =   300
      Width           =   1035
      Begin VB.PictureBox OldColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   0
         ScaleHeight     =   435
         ScaleWidth      =   975
         TabIndex        =   7
         Top             =   420
         Width           =   975
      End
      Begin VB.PictureBox NewColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   0
         ScaleHeight     =   435
         ScaleWidth      =   975
         TabIndex        =   6
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   315
      Index           =   0
      Left            =   4440
      TabIndex        =   3
      Top             =   3900
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   315
      Index           =   1
      Left            =   4440
      TabIndex        =   2
      Top             =   3540
      Width           =   1035
   End
   Begin VB.PictureBox ColorMap 
      AutoRedraw      =   -1  'True
      Height          =   3885
      Left            =   4080
      ScaleHeight     =   255
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   1
      Top             =   300
      Width           =   255
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         X1              =   -4
         X2              =   32
         Y1              =   32
         Y2              =   32
      End
   End
   Begin VB.PictureBox ColorSel 
      AutoRedraw      =   -1  'True
      Height          =   3885
      Left            =   60
      ScaleHeight     =   255
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   300
      Width           =   3885
      Begin VB.Shape SelObject 
         BorderColor     =   &H00FFFFFF&
         Height          =   165
         Index           =   1
         Left            =   1035
         Shape           =   3  'Circle
         Top             =   1695
         Width           =   165
      End
      Begin VB.Shape SelObject 
         BorderColor     =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   1020
         Shape           =   3  'Circle
         Top             =   1680
         Width           =   195
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Blue:"
      Height          =   195
      Index           =   2
      Left            =   4440
      TabIndex        =   13
      Top             =   1980
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Green:"
      Height          =   195
      Index           =   1
      Left            =   4440
      TabIndex        =   12
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Red:"
      Height          =   195
      Index           =   0
      Left            =   4440
      TabIndex        =   11
      Top             =   1380
      Width           =   615
   End
   Begin VB.Label lblRGB 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   10
      Top             =   1920
      Width           =   435
   End
   Begin VB.Label lblRGB 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   9
      Top             =   1620
      Width           =   435
   End
   Begin VB.Label lblRGB 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   8
      Top             =   1320
      Width           =   435
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Select color:"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   885
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MapPos As Integer
Dim MouseDown As Boolean
Dim SelMouseDown As Boolean


Private Sub ColorMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseDown = True
    
End Sub

Private Sub ColorMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tY As Single
    
    If MouseDown = True Then
        tY = Y
        
        If tY < 0 Then tY = 0
        If tY > Me.ColorMap.ScaleHeight - 1 Then tY = Me.ColorMap.ScaleHeight - 1
        
        Me.Line1.X1 = 0
        Me.Line1.X2 = 17
        Me.Line1.Y1 = tY
        Me.Line1.Y2 = tY
    End If
End Sub

Private Sub ColorMap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tY As Single
    
    tY = Y
    
    If tY < 0 Then tY = 0
    If tY > Me.ColorMap.ScaleHeight - 1 Then tY = Me.ColorMap.ScaleHeight - 1
    
    Me.Line1.X1 = 0
    Me.Line1.X2 = 17
    Me.Line1.Y1 = tY
    Me.Line1.Y2 = tY
    
    Me.Enabled = False
    Me.ColorMap.Enabled = False
    Me.ColorSel.Enabled = False
    Me.Command1(0).Enabled = False
    Me.Command1(1).Enabled = False
    
    MouseDown = False
    DrawGradient (GetPixel(Me.ColorMap.hdc, 1, Me.Line1.Y1))
    SelectColor Me.SelObject(0).Left + 6, Me.SelObject(0).Top + 6
    
    Me.Enabled = True
    Me.ColorMap.Enabled = True
    Me.ColorSel.Enabled = True
    Me.Command1(0).Enabled = True
    Me.Command1(1).Enabled = True
    
End Sub

Private Sub ColorSel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tX As Single, tY As Single
    
    tX = X
    tY = Y
    
    If tX < 0 Then tX = 0
    If tY < 0 Then tY = 0
    
    If tX > Me.ColorSel.ScaleWidth - 1 Then tX = Me.ColorSel.ScaleWidth - 1
    If tY > Me.ColorSel.ScaleHeight - 1 Then tY = Me.ColorSel.ScaleHeight - 1
    
    SelMouseDown = True
    SelectColor tX, tY
    
End Sub

Private Sub ColorSel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tX As Single, tY As Single
    
    If SelMouseDown = True Then
        tX = X
        tY = Y
        
        If tX < 0 Then tX = 0
        If tY < 0 Then tY = 0
        
        If tX > Me.ColorSel.ScaleWidth - 1 Then tX = Me.ColorSel.ScaleWidth - 1
        If tY > Me.ColorSel.ScaleHeight - 1 Then tY = Me.ColorSel.ScaleHeight - 1
        
        SelectColor tX, tY
    End If
End Sub

Private Sub ColorSel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tX As Single, tY As Single
    
    tX = X
    tY = Y
    
    If tX < 0 Then tX = 0
    If tY < 0 Then tY = 0
    
    If tX > Me.ColorSel.ScaleWidth - 1 Then tX = Me.ColorSel.ScaleWidth - 1
    If tY > Me.ColorSel.ScaleHeight - 1 Then tY = Me.ColorSel.ScaleHeight - 1
    
    SelMouseDown = False
    SelectColor tX, tY
    
End Sub

Private Sub SelectColor(X As Single, Y As Single)
    Dim r As Long, g As Long, b As Long
    
    
    Me.SelObject(0).Move X - 6, Y - 6
    Me.SelObject(1).Move X - 5, Y - 5
    
    b = Me.ColorSel.Point(X, Y) \ 65536
    g = (Me.ColorSel.Point(X, Y) - b * 65536) \ 256
    r = Me.ColorSel.Point(X, Y) - b * 65536 - g * 256

    
    Me.lblRGB(0).Caption = r
    Me.lblRGB(1).Caption = g
    Me.lblRGB(2).Caption = b
    
    If X < 0 Then X = 0
    If Y < 0 Then Y = 0
    
    
    Me.NewColor.BackColor = Me.ColorSel.Point(X, Y)
    
    On Error Resume Next
    frmMain.ActiveForm.Buffer.ForeColor = Me.NewColor.BackColor
    frmMain.ActiveForm.TextInput.ForeColor = Me.NewColor.BackColor
    
End Sub

Private Sub Command1_Click(Index As Integer)
    Me.Hide
    
    If Index = 1 Then
        frmMain.SelColor(SelColorIndex).BackColor = Me.NewColor.BackColor
        
        frmMain.ColorBlend(0).BackColor = frmMain.SelColor(0).BackColor
        frmMain.ColorBlend(1).BackColor = frmMain.SelColor(1).BackColor
        frmMain.SetColorBars
        
        DrawPreviewGradient
        
        On Error Resume Next
        frmMain.ActiveForm.Buffer.ForeColor = frmMain.SelColor(0).BackColor
        frmMain.ActiveForm.TextInput.ForeColor = frmMain.SelColor(0).BackColor
        frmMain.ActiveForm.TextInput.BackColor = frmMain.SelColor(1).BackColor
    End If
    
    frmMain.Enabled = True
    frmMain.SetFocus
    
End Sub

Private Sub DrawGradient(MyColor As Long)
    Dim C(2) As Byte
    Dim X As Integer
    Dim Y As Integer
    
    Dim r As Double, g As Double, b As Double
    Dim Dr As Double, Dg As Double, Db As Double
    Dim Sr As Double, Sg As Double, Sb As Double
    
    Dim r2 As Double, g2 As Double, b2 As Double
    Dim Sr2 As Double, Sg2 As Double, Sb2 As Double
    
    
    Db = MyColor \ 65536
    Dg = (MyColor - Db * 65536) \ 256
    Dr = MyColor - Db * 65536 - Dg * 256
    
    r = 255
    g = 255
    b = 255
    
    Sr = (r - Dr) / (Me.ColorSel.ScaleWidth - 1)
    Sg = (g - Dg) / (Me.ColorSel.ScaleWidth - 1)
    Sb = (b - Db) / (Me.ColorSel.ScaleWidth - 1)
        
    For X = 0 To Me.ColorSel.ScaleWidth - 1
        
        r2 = r
        g2 = g
        b2 = b
        
        Sr2 = r2 / (Me.ColorSel.ScaleHeight - 1)
        Sg2 = g2 / (Me.ColorSel.ScaleHeight - 1)
        Sb2 = b2 / (Me.ColorSel.ScaleHeight - 1)
        
        For Y = 0 To Me.ColorSel.ScaleHeight - 1
            SetPixelV Me.ColorSel.hdc, X, Y, RGB(r2, g2, b2)
             
            r2 = r2 - Sr2
            g2 = g2 - Sg2
            b2 = b2 - Sb2
        Next Y
        
        r = r - Sr
        g = g - Sg
        b = b - Sb
    Next X
    
    Me.ColorSel.Refresh
    
End Sub

Private Sub DrawAll()
    Dim i As Integer
    
    For i = 0 To Me.ColorMap.ScaleHeight / 6
        Me.ColorMap.Line (0, i)-(17, i), RGB(255, 0, i * (255 / (Me.ColorMap.ScaleHeight / 6))), BF
    Next i
    
    For i = 0 To Me.ColorMap.ScaleHeight / 6
        Me.ColorMap.Line (0, i + ((Me.ColorMap.ScaleHeight / 6) * 1))-(17, (i + 2) + ((Me.ColorMap.ScaleHeight / 6) * 1)), RGB(255 - i * (255 / (Me.ColorMap.ScaleHeight / 6)), 0, 255), BF
    Next i
    
    For i = 0 To Me.ColorMap.ScaleHeight / 6
        Me.ColorMap.Line (0, i + ((Me.ColorMap.ScaleHeight / 6) * 2))-(17, (i + 2) + ((Me.ColorMap.ScaleHeight / 6) * 2)), RGB(0, i * (255 / (Me.ColorMap.ScaleHeight / 6)), 255), BF
    Next i
    
    For i = 0 To Me.ColorMap.ScaleHeight / 6
        Me.ColorMap.Line (0, i + ((Me.ColorMap.ScaleHeight / 6) * 3))-(17, (i + 2) + ((Me.ColorMap.ScaleHeight / 6) * 3)), RGB(0, 255, 255 - i * (255 / (Me.ColorMap.ScaleHeight / 6))), BF
    Next i

    For i = 0 To Me.ColorMap.ScaleHeight / 6
        Me.ColorMap.Line (0, i + ((Me.ColorMap.ScaleHeight / 6) * 4))-(17, (i + 2) + ((Me.ColorMap.ScaleHeight / 6) * 4)), RGB(i * (255 / (Me.ColorMap.ScaleHeight / 6)), 255, 0), BF
    Next i
    
    For i = 0 To Me.ColorMap.ScaleHeight / 6
        Me.ColorMap.Line (0, i + ((Me.ColorMap.ScaleHeight / 6) * 5))-(17, (i + 2) + ((Me.ColorMap.ScaleHeight / 6) * 5)), RGB(255, 255 - i * (255 / (Me.ColorMap.ScaleHeight / 6)), 0), BF
    Next i
End Sub

Private Sub Form_Load()
    Dim r As Long, g As Long, b As Long
    
    DrawAll
    
    Me.SelObject(0).Left = -20
    Me.SelObject(1).Left = -20
    
    Me.Line1.X1 = 0
    Me.Line1.X2 = 17
    Me.Line1.Y1 = 0
    Me.Line1.Y2 = 0
    
    Me.NewColor.BackColor = frmMain.SelColor(SelColorIndex).BackColor
    
    b = Me.NewColor.BackColor \ 65536
    g = (Me.NewColor.BackColor - b * 65536) \ 256
    r = Me.NewColor.BackColor - b * 65536 - g * 256

    Me.lblRGB(0).Caption = r
    Me.lblRGB(1).Caption = g
    Me.lblRGB(2).Caption = b
    
    
    DrawGradient (GetPixel(Me.ColorMap.hdc, 1, Me.Line1.Y1))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    frmMain.Enabled = True
    frmMain.SetFocus
    
    Cancel = True
End Sub
