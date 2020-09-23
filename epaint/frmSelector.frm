VERSION 5.00
Begin VB.Form frmSelector 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   29
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   133
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox ScrollBlock 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   360
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   1
      Top             =   60
      Width           =   135
   End
   Begin VB.PictureBox ScrollBar 
      BackColor       =   &H80000000&
      Height          =   135
      Left            =   60
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   0
      Top             =   150
      Width           =   1875
   End
End
Attribute VB_Name = "frmSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MouseDown As Boolean
Dim Xpos As Integer



Private Sub Form_Load()
    Me.Hide
End Sub

Private Sub ScrollBlock_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Xpos = X
    MouseDown = True
    
End Sub

Private Sub ScrollBlock_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim TmpX As Integer
    
    If MouseDown = True Then
        TmpX = Me.ScrollBlock.Left + X - Xpos
        If TmpX < Me.ScrollBar.Left Then TmpX = Me.ScrollBar.Left
        If TmpX > Me.ScrollBar.Left + Me.ScrollBar.Width - Me.ScrollBlock.Width Then TmpX = Me.ScrollBar.Left + Me.ScrollBar.Width - Me.ScrollBlock.Width
        Me.ScrollBlock.Left = TmpX
        
        CalcPos
    End If
End Sub

Private Sub ScrollBlock_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseDown = False
End Sub

Public Sub CalcPos()
    Dim percent As Double

    percent = ((Me.ScrollBlock.Left - Me.ScrollBar.Left) / (Me.ScrollBar.Width - (Me.ScrollBlock.Width)) * (SelMethod(Me.Tag).Max - SelMethod(Me.Tag).Min)) + SelMethod(Me.Tag).Min
    frmControls.lblSelector(Me.Tag).Caption = CInt(percent) & frmControls.lblSelector(Me.Tag).Tag
    SelMethod(Me.Tag).Current = percent
    
    If Me.Tag = 7 Then
        On Error Resume Next
        frmMain.ActiveForm.Buffer.FontSize = SelMethod(7).Current
        frmMain.ActiveForm.TextInput.FontSize = SelMethod(7).Current * (frmMain.ActiveForm.GetZoomFactor / 100)
        frmMain.ActiveForm.lblTextSize.FontSize = SelMethod(7).Current * (frmMain.ActiveForm.GetZoomFactor / 100)
        frmMain.ActiveForm.TextInput.Width = frmMain.ActiveForm.lblTextSize.Width
        frmMain.ActiveForm.TextInput.Height = frmMain.ActiveForm.lblTextSize.Height
    End If
End Sub
