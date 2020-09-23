VERSION 5.00
Begin VB.Form frmPaint 
   Caption         =   "Untitled"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   312
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   561
   Begin VB.PictureBox Undo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   5340
      ScaleHeight     =   135
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   183
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.PictureBox SelectedBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   2460
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   177
      TabIndex        =   6
      Top             =   3300
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.PictureBox BufferSelected 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   2460
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   177
      TabIndex        =   5
      Top             =   2340
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.PictureBox Buffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   2460
      ScaleHeight     =   135
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   183
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.PictureBox Back 
      BorderStyle     =   0  'None
      FillStyle       =   7  'Diagonal Cross
      Height          =   2475
      Left            =   0
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   137
      TabIndex        =   2
      Top             =   0
      Width           =   2055
      Begin VB.PictureBox PaintArea 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         DrawWidth       =   10
         ForeColor       =   &H80000008&
         Height          =   2535
         Left            =   -60
         ScaleHeight     =   167
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   167
         TabIndex        =   3
         Top             =   540
         Width           =   2535
         Begin VB.TextBox TextInput 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   7
            Top             =   60
            Visible         =   0   'False
            Width           =   75
         End
         Begin VB.Shape DrawCircle 
            DrawMode        =   6  'Mask Pen Not
            Height          =   975
            Left            =   780
            Shape           =   2  'Oval
            Top             =   720
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Shape DrawBox 
            DrawMode        =   6  'Mask Pen Not
            Height          =   1335
            Left            =   360
            Top             =   480
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Line Temp2 
            BorderColor     =   &H0000FF00&
            Visible         =   0   'False
            X1              =   40
            X2              =   84
            Y1              =   80
            Y2              =   12
         End
         Begin VB.Line Temp 
            BorderColor     =   &H000000FF&
            Visible         =   0   'False
            X1              =   96
            X2              =   40
            Y1              =   40
            Y2              =   112
         End
         Begin VB.Shape CropArea 
            BorderStyle     =   3  'Dot
            DrawMode        =   6  'Mask Pen Not
            Height          =   1095
            Left            =   540
            Top             =   180
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Shape SelectArea 
            BorderStyle     =   3  'Dot
            DrawMode        =   6  'Mask Pen Not
            Height          =   1095
            Left            =   720
            Top             =   480
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Line FollowLine 
            DrawMode        =   6  'Mask Pen Not
            Visible         =   0   'False
            X1              =   32
            X2              =   136
            Y1              =   60
            Y2              =   96
         End
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      Max             =   0
      TabIndex        =   1
      Top             =   2460
      Width           =   2055
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2475
      Left            =   2040
      Max             =   0
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lblTextSize 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   840
      TabIndex        =   8
      Top             =   3240
      Width           =   45
   End
End
Attribute VB_Name = "frmPaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MouseDown As Boolean
Dim MouseStart As POINTAPI
Dim ZoomFactor As Double
Dim Title As String
Dim MyX As Single, MyY As Single
Dim Dirty As Boolean
Dim ControlDown As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 17 Then ControlDown = True
    
    If CurrentButton <> 12 And ControlDown = False Then
        Select Case KeyCode
            Case Asc("C")
                SelectTool 0
                
            Case Asc("S")
                SelectTool 5
                
            Case Asc("R")
                SelectTool 7
                
            Case Asc("Z")
                SelectTool 2
                
            Case Asc("A")
                SelectTool 6
                
            Case Asc("P")
                SelectTool 1
                
            Case Asc("G")
                SelectTool 4
                
            Case Asc("F")
                SelectTool 3
                
            Case Asc("E")
                SelectTool 10
                
            Case Asc("B")
                SelectTool 9
                
            Case Asc("W")
                SelectTool 13
                
            Case Asc("T")
                SelectTool 12
                
            Case Asc("L")
                SelectTool 11
                
            Case Asc("Y")
                SelectTool 8
                
        End Select
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 17 Then ControlDown = False
    
End Sub

Private Sub Form_Load()
    ZoomFactor = 100
    Dirty = False
   
    Form_Resize
    
End Sub

Private Sub Form_Paint()
    Form_Resize
    
End Sub



Private Sub Form_Resize()
    'just in case...
    On Error Resume Next
    'resize the paint form
    Me.PaintArea.AutoRedraw = False
    Me.HScroll1.Move 0, Me.ScaleHeight - Me.HScroll1.Height, Me.ScaleWidth - Me.VScroll1.Width, Me.HScroll1.Height
    Me.VScroll1.Move Me.ScaleWidth - Me.VScroll1.Width, 0, Me.VScroll1.Width, Me.ScaleHeight - Me.HScroll1.Height
    Me.Back.Move 0, 0, Me.ScaleWidth - Me.VScroll1.Width, Me.ScaleHeight - Me.HScroll1.Height
    
    If Me.PaintArea.Width < Me.Back.ScaleWidth Then
        Me.PaintArea.Left = (Me.Back.ScaleWidth - Me.PaintArea.Width) / 2
        Me.HScroll1.Min = 0
        Me.HScroll1.Max = 0
    Else
        Me.PaintArea.Left = -Me.HScroll1.Value
        Me.HScroll1.SmallChange = 1
        Me.HScroll1.Max = Me.PaintArea.Width - Me.Back.ScaleWidth
        Me.HScroll1.LargeChange = Me.PaintArea.Width - (((Me.PaintArea.Width - Me.Back.ScaleWidth) / Me.PaintArea.Width) * Me.PaintArea.Width)
    End If
    
    If Me.PaintArea.Height < Me.Back.ScaleHeight Then
        Me.PaintArea.Top = (Me.Back.ScaleHeight - Me.PaintArea.Height) / 2
        Me.VScroll1.Min = 0
        Me.VScroll1.Max = 0
    Else
        Me.PaintArea.Top = -Me.VScroll1.Value
        Me.VScroll1.Max = Me.PaintArea.Height - Me.Back.ScaleHeight
        Me.VScroll1.SmallChange = 1
        Me.VScroll1.LargeChange = Me.PaintArea.Height - (((Me.PaintArea.Height - Me.Back.ScaleHeight) / Me.PaintArea.Height) * Me.PaintArea.Height)
    End If
    Me.PaintArea.AutoRedraw = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim FileName As String
    
    Dim X As Integer
    If Dirty = True Then
        X = MsgBox("The image has been changed." & vbCrLf & vbCrLf & "Do you wish to save the changes?", vbExclamation + vbYesNo, "The image has been changed")
        
        If X = 6 Then
            If frmMain.ActiveForm.Buffer.Tag <> "" Then
                SavePicture frmMain.ActiveForm.Buffer.Image, frmMain.ActiveForm.Buffer.Tag
                SetDirtyFalse
            Else
                FileName = GetSaveName("Save As...")
                 
                If FileName <> "" Then
                    SavePicture frmMain.ActiveForm.Buffer.Image, FileName
                    frmMain.ActiveForm.Buffer.Tag = FileName
                    frmMain.ActiveForm.Caption = FileName & " - " & frmMain.ActiveForm.GetZoomFactor & "%"
                    SetDirtyFalse
                Else
                    X = MsgBox("Save failed." & vbCrLf & vbCrLf & "Close anyway?", vbCritical + vbYesNo, "Save failed.")
                    If X <> 6 Then Cancel = True
                End If
            End If
        End If
    End If
    
End Sub

Public Function SetDirtyFalse()
    Dirty = False
    
End Function

Private Sub HScroll1_GotFocus()
    Me.PaintArea.Refresh
    Me.PaintArea.SetFocus
End Sub

Private Sub HScroll1_Scroll()
    'scroll the picture...
    Me.PaintArea.SetFocus
    Me.PaintArea.Left = -Me.HScroll1.Value
    RealignZoom -1, -1
End Sub

Private Sub PaintArea_GotFocus()
    frmControls.lblZoom.Caption = ZoomFactor & " %"
    Me.PaintArea.Refresh
    
End Sub

Private Sub PaintArea_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        'prepare for undo...
        If Me.SelectArea.Visible = False Then
            frmMain.ActiveForm.Undo.Width = frmMain.ActiveForm.Buffer.Width
            frmMain.ActiveForm.Undo.Height = frmMain.ActiveForm.Buffer.Height
            BitBlt frmMain.ActiveForm.Undo.hdc, 0, 0, frmMain.ActiveForm.Buffer.ScaleWidth, frmMain.ActiveForm.Buffer.ScaleHeight, frmMain.ActiveForm.Buffer.hdc, 0, 0, vbSrcCopy
        Else
            frmMain.ActiveForm.Undo.Width = frmMain.ActiveForm.BufferSelected.Width
            frmMain.ActiveForm.Undo.Height = frmMain.ActiveForm.BufferSelected.Height
            BitBlt frmMain.ActiveForm.Undo.hdc, 0, 0, frmMain.ActiveForm.BufferSelected.ScaleWidth, frmMain.ActiveForm.BufferSelected.ScaleHeight, frmMain.ActiveForm.BufferSelected.hdc, 0, 0, vbSrcCopy
        End If
        
        MouseDown = True
        Buffer.CurrentX = ((X - ZoomFactor / 200) * (100 / ZoomFactor))
        Buffer.CurrentY = ((Y - ZoomFactor / 200) * (100 / ZoomFactor))
        BufferSelected.CurrentX = Buffer.CurrentX - Me.BufferSelected.Left
        BufferSelected.CurrentY = Buffer.CurrentY - Me.BufferSelected.Top
        
        MouseStart.X = ((X - ZoomFactor / 200) * (100 / ZoomFactor))
        MouseStart.Y = ((Y - ZoomFactor / 200) * (100 / ZoomFactor))
        Draw Buffer, (X - ZoomFactor / 200) / ZoomFactor * 100, (Y - ZoomFactor / 200) / ZoomFactor * 100, ZoomFactor, MouseDown
    Else
        MouseDown = False
    End If
End Sub

Private Sub PaintArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MouseDown = True Then
        Draw Buffer, (X - ZoomFactor / 200) / ZoomFactor * 100, (Y - ZoomFactor / 200) / ZoomFactor * 100, ZoomFactor, MouseDown

        If ZoomFactor <> 100 Then
            frmMain.lblArea.Caption = CInt((Abs(MouseStart.X - (X - ZoomFactor / 200) * (100 / ZoomFactor)))) & " x " & CInt((Abs(MouseStart.Y - (Y - ZoomFactor / 200) * (100 / ZoomFactor))))
        Else
            frmMain.lblArea.Caption = CInt(Abs(MouseStart.X - X)) & " x " & CInt(Abs(MouseStart.Y - Y))
        End If
    End If
    
    If ZoomFactor <> 100 Then
        frmMain.lblCoords = CInt((X - ZoomFactor / 200) / ZoomFactor * 100) & ", " & CInt((Y - ZoomFactor / 200) / ZoomFactor * 100)
    Else
        frmMain.lblCoords = CInt(X) & ", " & CInt(Y)
    End If
    
    MyX = (X - ZoomFactor / 200) / ZoomFactor * 100
    MyY = (Y - ZoomFactor / 200) / ZoomFactor * 100
    
    'select the correct mousecursor
    Select Case CurrentButton
        Case 0
            If Me.PaintArea.MousePointer <> 0 Then Me.PaintArea.MousePointer = 0
        
        Case Else
            If Me.PaintArea.MouseIcon <> frmControls.MyCursor(CurrentButton).MouseIcon Then
                Me.PaintArea.MouseIcon = frmControls.MyCursor(CurrentButton).MouseIcon
                Me.PaintArea.MousePointer = 99
            End If
    End Select
    
    If CurrentButton = 8 Then
        ShowColorUnderMouse (X - ZoomFactor / 200) / ZoomFactor * 100, (Y - ZoomFactor / 200) / ZoomFactor * 100
    End If
    
End Sub


Public Function GetZoomFactor()
    GetZoomFactor = ZoomFactor
End Function

Private Sub PaintArea_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim TmpW As Integer, TmpH As Integer
    
    TmpW = Me.PaintArea.ScaleWidth
    TmpH = Me.PaintArea.ScaleHeight
    
    Dirty = True
    
    MouseDown = False
    If Button = 1 Then Draw Buffer, (X - ZoomFactor / 200) / ZoomFactor * 100, (Y - ZoomFactor / 200) / ZoomFactor * 100, ZoomFactor, MouseDown
    
    If Button = 1 Then
        If CurrentButton = 2 And ZoomFactor < 1000 Then
            Me.PaintArea.AutoRedraw = False
            ZoomFactor = ZoomFactor + 100
            Me.PaintArea.Width = (Me.Buffer.ScaleWidth * (ZoomFactor / 100)) + 2
            Me.PaintArea.Height = (Me.Buffer.ScaleHeight * (ZoomFactor / 100)) + 2
            Me.Caption = Me.Tag & " - " & ZoomFactor & "%"
            RealignZoom CDbl(X / TmpW), CDbl(Y / TmpH)
        End If
    Else
        If CurrentButton = 2 And ZoomFactor > 100 Then
            Me.PaintArea.AutoRedraw = False
            ZoomFactor = ZoomFactor - 100
            Me.PaintArea.Width = (Me.Buffer.ScaleWidth * (ZoomFactor / 100)) + 2
            Me.PaintArea.Height = (Me.Buffer.ScaleHeight * (ZoomFactor / 100)) + 2
            Me.Caption = Me.Tag & " - " & ZoomFactor & "%"
            RealignZoom CDbl(X / TmpW), CDbl(Y / TmpH)
        End If
    End If

    
    frmControls.lblZoom.Caption = ZoomFactor & " %"
    
End Sub


Public Sub RealignZoom(X As Double, Y As Double)
    
    Me.Enabled = False
    'Me.PaintArea.Cls
    Me.PaintArea.AutoRedraw = False
    Form_Resize
    
    If X >= 0 Then
        Me.VScroll1.Value = X * Me.VScroll1.Max
    End If
    If Y >= 0 Then
        Me.HScroll1.Value = Y * Me.HScroll1.Max
    End If
    Form_Resize

    If Me.CropArea.Visible = True Then
        Me.CropArea.Move Me.BufferSelected.Left * (ZoomFactor / 100), Me.BufferSelected.Top * (ZoomFactor / 100), Me.BufferSelected.Width * (ZoomFactor / 100), Me.BufferSelected.Height * (ZoomFactor / 100)
    End If
    
    If Me.SelectArea.Visible = True Then
        Me.SelectArea.Move Me.BufferSelected.Left * (ZoomFactor / 100), Me.BufferSelected.Top * (ZoomFactor / 100), Me.BufferSelected.Width * (ZoomFactor / 100), Me.BufferSelected.Height * (ZoomFactor / 100)
        OriginalSelX = (Me.SelectArea.Left / 100) * 100
        OriginalSelY = (Me.SelectArea.Top / 100) * 100
    End If
    
    Me.PaintArea.AutoRedraw = True
    
    If ZoomFactor = 100 Then
        BitBlt PaintArea.hdc, 0, 0, PaintArea.ScaleWidth, PaintArea.ScaleHeight, Buffer.hdc, 0, 0, vbSrcCopy
        PaintArea.Refresh
    Else
        StretchBlt PaintArea.hdc, HScroll1.Value, VScroll1.Value, Back.ScaleWidth, Back.ScaleHeight, _
                   Buffer.hdc, (HScroll1.Value / ZoomFactor) * 100, (VScroll1.Value / ZoomFactor) * 100, (Back.ScaleWidth / ZoomFactor) * 100, (Back.ScaleHeight / ZoomFactor) * 100, vbSrcCopy
        
        PaintArea.Refresh
    End If
    Me.Enabled = True
    
End Sub


Private Sub TextInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.TextInput.Visible = True Then
        Me.lblTextSize.Caption = Me.TextInput.Text & "M"
        Me.TextInput.Width = Me.lblTextSize.Width
        Me.TextInput.Height = Me.lblTextSize.Height
    End If
End Sub

Private Sub TextInput_KeyUp(KeyCode As Integer, Shift As Integer)
    If Me.TextInput.Visible = True Then
        Me.lblTextSize.Caption = Me.TextInput.Text & "M"
        Me.TextInput.Width = Me.lblTextSize.Width
        Me.TextInput.Height = Me.lblTextSize.Height
    End If
End Sub

Private Sub VScroll1_GotFocus()
    Me.PaintArea.Refresh
    Me.PaintArea.SetFocus
End Sub

Private Sub VScroll1_Scroll()
    'scroll the picture...
    Me.PaintArea.SetFocus
    Me.PaintArea.Top = -Me.VScroll1.Value
    RealignZoom -1, -1
End Sub


Private Sub ShowColorUnderMouse(X As Single, Y As Single)
    Dim NewColor As Long
    Dim r As Long, g As Long, b As Long
    
    If X >= 0 And Y >= 0 And X < Me.Buffer.ScaleWidth And Y < Me.Buffer.ScaleHeight Then
        On Error Resume Next
        NewColor = GetPixel(Me.Buffer.hdc, X, Y)
        frmControls.ColorPick.BackColor = NewColor
        
        b = NewColor \ 65536
        g = (NewColor - b * 65536) \ 256
        r = NewColor - b * 65536 - g * 256
        
        frmControls.lblRed.Caption = r
        frmControls.lblGreen.Caption = g
        frmControls.lblBlue.Caption = b
    End If
    
End Sub
