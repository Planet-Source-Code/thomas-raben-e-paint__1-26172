Attribute VB_Name = "GuiApi"
Option Explicit

Const PROCESS_ALL_ACCESS& = &H1F0FFF
Const STILL_ACTIVE& = &H103&
Const INFINITE& = &HFFFF

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type SelectorMethod
    Min As Integer
    Max As Integer
    Current As Double
End Type

Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Public Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Public CheckedButton(20) As Boolean
Public CheckedButtonRect(20) As Boolean
Public CheckedButtonBrush(20) As Boolean
Public CheckedButtonLight(20) As Boolean

Public CurrentButton As Integer
Public CurrentButtonRect As Integer
Public CurrentButtonBrush As Integer
Public CurrentButtonLight As Integer

Public SelColorIndex As Integer

Public ClipBoardGotData As Boolean

Public SelMethod(10) As SelectorMethod

Public Sub main()
    Dim X As Long
    Dim SPW As Long
    Dim SPH As Long
    
    SPW = 401
    SPH = 301
    
    frmSPlash.Show
    X = SetWindowPos(frmSPlash.hwnd, -1, (Screen.Width / 15 - SPW) / 2, (Screen.Height / 15 - SPH) / 2, SPW, SPH, 0)
    frmSPlash.SetFocus
    DoEvents
    frmMain.Show
    
End Sub

Public Sub LoadFilters()
    Dim Path As String
    Dim i As Integer
    Dim Title As String
    
    Path = App.Path
    
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    Path = Path & "filters"
    
    frmMain.lstFilters.Path = Path
    
    For i = 0 To frmMain.lstFilters.ListCount - 1
        If i <> 0 Then
            Load frmMain.Filter(i)
        End If
        
        frmMain.Filter(i).Visible = True
        
        frmMain.Filter(i).Tag = frmMain.lstFilters.List(i)
        
        Title = Mid(frmMain.lstFilters.List(i), 1, InStr(1, frmMain.lstFilters.List(i), ".exe") - 1)
        Title = UCase(Left(Title, 1)) & Mid(Title, 2)
        frmMain.Filter(i).Caption = Title
    Next i
    
End Sub

Public Sub ExecFilter(Filter As String)
    Dim X%
    Dim Path As String
    Dim Parms As String

    Path = App.Path
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    Path = Path & "temp\filter.bmp"

    If frmMain.ActiveForm.SelectArea.Visible = True Then
        SavePicture frmMain.ActiveForm.BufferSelected.Image, Path
        Parms = " " & Path
    Else
        SavePicture frmMain.ActiveForm.Buffer.Image, Path
        Parms = " " & Path
    End If
    
    Path = App.Path
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    Path = Path & "filters\" & Filter & Parms
    
    frmMain.Enabled = False
    
    X% = Shell(Path, vbNormalFocus)
    
    fWait X%
    
    If frmMain.ActiveForm.SelectArea.Visible = True Then
        frmMain.ActiveForm.BufferSelected.Picture = LoadPicture(Trim(Parms))
    Else
        frmMain.ActiveForm.Buffer.Picture = LoadPicture(Trim(Parms))
    End If
    
    Kill Trim(Parms)
    
    frmMain.Enabled = True
    frmMain.SetFocus
    
    UpdateArea frmMain.ActiveForm.Buffer, 0, 0, frmMain.ActiveForm.GetZoomFactor

End Sub

Function fWait(ByVal lProgID As Long) As Long
    Dim lExitCode As Long, hdlProg As Long
    hdlProg = OpenProcess(PROCESS_ALL_ACCESS, False, lProgID)
    GetExitCodeProcess hdlProg, lExitCode

    Do While lExitCode = STILL_ACTIVE&
        DoEvents
        GetExitCodeProcess hdlProg, lExitCode
    Loop
    
    CloseHandle hdlProg
    fWait = lExitCode
End Function




Private Sub SetSelMethods()
    'Airbrush Pressure
    SelMethod(0).Min = 0
    SelMethod(0).Max = 100
    SelMethod(0).Current = 50
    frmControls.lblSelector(0).Caption = "50" & frmControls.lblSelector(0).Tag
    
    'Light darken Pressure
    SelMethod(1).Min = 0
    SelMethod(1).Max = 100
    SelMethod(1).Current = 50
    frmControls.lblSelector(1).Caption = "50" & frmControls.lblSelector(1).Tag
    
    'Airbrush Size
    SelMethod(2).Min = 1
    SelMethod(2).Max = 100
    SelMethod(2).Current = 10
    frmControls.lblSelector(2).Caption = "10" & frmControls.lblSelector(2).Tag
    
    'Light darken Size
    SelMethod(3).Min = 1
    SelMethod(3).Max = 100
    SelMethod(3).Current = 10
    frmControls.lblSelector(3).Caption = "10" & frmControls.lblSelector(3).Tag
    
    'Brush Size
    SelMethod(4).Min = 1
    SelMethod(4).Max = 100
    SelMethod(4).Current = 1
    frmControls.lblSelector(4).Caption = "1" & frmControls.lblSelector(4).Tag
    
    'Rect Size
    SelMethod(5).Min = 1
    SelMethod(5).Max = 100
    SelMethod(5).Current = 1
    frmControls.lblSelector(5).Caption = "1" & frmControls.lblSelector(5).Tag
    
    'circle Size
    SelMethod(6).Min = 1
    SelMethod(6).Max = 100
    SelMethod(6).Current = 1
    frmControls.lblSelector(6).Caption = "1" & frmControls.lblSelector(6).Tag
    
    'Font Size
    SelMethod(7).Min = 1
    SelMethod(7).Max = 100
    SelMethod(7).Current = 10
    frmControls.lblSelector(7).Caption = "10" & frmControls.lblSelector(7).Tag
    
    'Blur Size
    SelMethod(8).Min = 1
    SelMethod(8).Max = 100
    SelMethod(8).Current = 10
    frmControls.lblSelector(8).Caption = "10" & frmControls.lblSelector(8).Tag
    
End Sub

Public Sub CheckFlatButtons()
    Dim i As Integer
    Dim Rec As RECT, Point As POINTAPI
    
    GetCursorPos Point
    
    For i = 0 To frmMain.Cmd.UBound
        GetWindowRect frmMain.Cmd(i).hwnd, Rec
        
        If Point.X < Rec.Left Or Point.X > Rec.Right Or Point.Y < Rec.Top Or Point.Y > Rec.Bottom Then
            If CheckedButton(i) = False Then
                frmMain.Cmd(i).BackColor = vb3DFace
                frmMain.Cmd(i).Line (0, 0)-(frmMain.Cmd(i).ScaleWidth - 1, frmMain.Cmd(i).ScaleHeight - 1), vb3DFace, B
            End If
        Else
            If CheckedButton(i) = False Then
                frmMain.Cmd(i).BackColor = vb3DFace
                frmMain.Cmd(i).Line (0, 0)-(frmMain.Cmd(i).ScaleWidth - 1, 0), vb3DHighlight
                frmMain.Cmd(i).Line (0, 0)-(0, frmMain.Cmd(i).ScaleHeight - 1), vb3DHighlight
                frmMain.Cmd(i).Line (frmMain.Cmd(i).ScaleWidth - 1, 1)-(frmMain.Cmd(i).ScaleWidth - 1, frmMain.Cmd(i).ScaleHeight), vb3DShadow
                frmMain.Cmd(i).Line (0, frmMain.Cmd(i).ScaleHeight - 1)-(frmMain.Cmd(i).ScaleWidth - 1, frmMain.Cmd(i).ScaleHeight - 1), vb3DShadow
            End If
        End If
    Next i
End Sub

Public Sub CheckFlatButtonsRect()
    Dim i As Integer
    Dim Rec As RECT, Point As POINTAPI
    
    GetCursorPos Point
    
    For i = 0 To frmControls.Cmd.UBound
        GetWindowRect frmControls.Cmd(i).hwnd, Rec
        
        If Point.X < Rec.Left Or Point.X > Rec.Right Or Point.Y < Rec.Top Or Point.Y > Rec.Bottom Then
            If CheckedButtonRect(i) = False Then
                frmControls.Cmd(i).BackColor = vb3DFace
                frmControls.Cmd(i).Line (0, 0)-(frmControls.Cmd(i).ScaleWidth - 1, frmControls.Cmd(i).ScaleHeight - 1), vb3DFace, B
            End If
        Else
            If CheckedButtonRect(i) = False Then
                frmControls.Cmd(i).BackColor = vb3DFace
                frmControls.Cmd(i).Line (0, 0)-(frmControls.Cmd(i).ScaleWidth - 1, 0), vb3DHighlight
                frmControls.Cmd(i).Line (0, 0)-(0, frmControls.Cmd(i).ScaleHeight - 1), vb3DHighlight
                frmControls.Cmd(i).Line (frmControls.Cmd(i).ScaleWidth - 1, 1)-(frmControls.Cmd(i).ScaleWidth - 1, frmControls.Cmd(i).ScaleHeight), vb3DShadow
                frmControls.Cmd(i).Line (0, frmControls.Cmd(i).ScaleHeight - 1)-(frmControls.Cmd(i).ScaleWidth - 1, frmControls.Cmd(i).ScaleHeight - 1), vb3DShadow
            End If
        End If
    Next i
End Sub

Public Sub CheckFlatButtonsCircle()
    Dim i As Integer
    Dim Rec As RECT, Point As POINTAPI
    
    GetCursorPos Point
    
    For i = 0 To frmControls.Cmd2.UBound
        GetWindowRect frmControls.Cmd2(i).hwnd, Rec
        
        If Point.X < Rec.Left Or Point.X > Rec.Right Or Point.Y < Rec.Top Or Point.Y > Rec.Bottom Then
            If CheckedButtonRect(i) = False Then
                frmControls.Cmd2(i).BackColor = vb3DFace
                frmControls.Cmd2(i).Line (0, 0)-(frmControls.Cmd2(i).ScaleWidth - 1, frmControls.Cmd2(i).ScaleHeight - 1), vb3DFace, B
            End If
        Else
            If CheckedButtonRect(i) = False Then
                frmControls.Cmd2(i).BackColor = vb3DFace
                frmControls.Cmd2(i).Line (0, 0)-(frmControls.Cmd2(i).ScaleWidth - 1, 0), vb3DHighlight
                frmControls.Cmd2(i).Line (0, 0)-(0, frmControls.Cmd2(i).ScaleHeight - 1), vb3DHighlight
                frmControls.Cmd2(i).Line (frmControls.Cmd2(i).ScaleWidth - 1, 1)-(frmControls.Cmd2(i).ScaleWidth - 1, frmControls.Cmd2(i).ScaleHeight), vb3DShadow
                frmControls.Cmd2(i).Line (0, frmControls.Cmd2(i).ScaleHeight - 1)-(frmControls.Cmd2(i).ScaleWidth - 1, frmControls.Cmd2(i).ScaleHeight - 1), vb3DShadow
            End If
        End If
    Next i
End Sub

Public Sub CheckFlatButtonsBrush()
    Dim i As Integer
    Dim Rec As RECT, Point As POINTAPI
    
    GetCursorPos Point
    
    For i = 0 To frmControls.Cmd3.UBound
        GetWindowRect frmControls.Cmd3(i).hwnd, Rec
        
        If Point.X < Rec.Left Or Point.X > Rec.Right Or Point.Y < Rec.Top Or Point.Y > Rec.Bottom Then
            If CheckedButtonBrush(i) = False Then
                frmControls.Cmd3(i).BackColor = vb3DFace
                frmControls.Cmd3(i).Line (0, 0)-(frmControls.Cmd3(i).ScaleWidth - 1, frmControls.Cmd3(i).ScaleHeight - 1), vb3DFace, B
            End If
        Else
            If CheckedButtonBrush(i) = False Then
                frmControls.Cmd3(i).BackColor = vb3DFace
                frmControls.Cmd3(i).Line (0, 0)-(frmControls.Cmd3(i).ScaleWidth - 1, 0), vb3DHighlight
                frmControls.Cmd3(i).Line (0, 0)-(0, frmControls.Cmd3(i).ScaleHeight - 1), vb3DHighlight
                frmControls.Cmd3(i).Line (frmControls.Cmd3(i).ScaleWidth - 1, 1)-(frmControls.Cmd3(i).ScaleWidth - 1, frmControls.Cmd3(i).ScaleHeight), vb3DShadow
                frmControls.Cmd3(i).Line (0, frmControls.Cmd3(i).ScaleHeight - 1)-(frmControls.Cmd3(i).ScaleWidth - 1, frmControls.Cmd3(i).ScaleHeight - 1), vb3DShadow
            End If
        End If
    Next i
End Sub

Public Sub CheckFlatButtonsLight()
    Dim i As Integer
    Dim Rec As RECT, Point As POINTAPI
    
    GetCursorPos Point
    
    For i = 0 To frmControls.Cmd4.UBound
        GetWindowRect frmControls.Cmd4(i).hwnd, Rec
        
        If Point.X < Rec.Left Or Point.X > Rec.Right Or Point.Y < Rec.Top Or Point.Y > Rec.Bottom Then
            If CheckedButtonLight(i) = False Then
                frmControls.Cmd4(i).BackColor = vb3DFace
                frmControls.Cmd4(i).Line (0, 0)-(frmControls.Cmd4(i).ScaleWidth - 1, frmControls.Cmd4(i).ScaleHeight - 1), vb3DFace, B
            End If
        Else
            If CheckedButtonLight(i) = False Then
                frmControls.Cmd4(i).BackColor = vb3DFace
                frmControls.Cmd4(i).Line (0, 0)-(frmControls.Cmd4(i).ScaleWidth - 1, 0), vb3DHighlight
                frmControls.Cmd4(i).Line (0, 0)-(0, frmControls.Cmd4(i).ScaleHeight - 1), vb3DHighlight
                frmControls.Cmd4(i).Line (frmControls.Cmd4(i).ScaleWidth - 1, 1)-(frmControls.Cmd4(i).ScaleWidth - 1, frmControls.Cmd4(i).ScaleHeight), vb3DShadow
                frmControls.Cmd4(i).Line (0, frmControls.Cmd4(i).ScaleHeight - 1)-(frmControls.Cmd4(i).ScaleWidth - 1, frmControls.Cmd4(i).ScaleHeight - 1), vb3DShadow
            End If
        End If
    Next i
End Sub

Public Sub SelectTool(i As Integer)
    Dim BarX As Long
    
    frmControls.DrawToolbar(CurrentButton).Visible = False
    frmControls.DrawToolbar(i).Visible = True
    
    frmMain.Cmd(CurrentButton).BackColor = vb3DFace
    frmMain.Cmd(CurrentButton).Line (0, 0)-(frmMain.Cmd(CurrentButton).ScaleWidth - 1, frmMain.Cmd(CurrentButton).ScaleHeight - 1), vb3DFace, B
    CheckedButton(CurrentButton) = False
    
    CheckedButton(i) = True
    CurrentButton = i
    
    
    frmMain.Cmd(i).BackColor = vbScrollBars
    frmMain.Cmd(i).Line (0, 0)-(frmMain.Cmd(i).ScaleWidth - 1, 0), vb3DShadow
    frmMain.Cmd(i).Line (0, 0)-(0, frmMain.Cmd(i).ScaleHeight - 1), vb3DShadow
    frmMain.Cmd(i).Line (frmMain.Cmd(i).ScaleWidth - 1, 1)-(frmMain.Cmd(i).ScaleWidth - 1, frmMain.Cmd(i).ScaleHeight), vb3DHighlight
    frmMain.Cmd(i).Line (0, frmMain.Cmd(i).ScaleHeight - 1)-(frmMain.Cmd(i).ScaleWidth - 1, frmMain.Cmd(i).ScaleHeight - 1), vb3DHighlight
    
    
    On Error Resume Next
    Select Case CurrentButton
        Case 0
            frmMain.ActiveForm.PaintArea.MousePointer = 0
        
        Case Else
            frmMain.ActiveForm.PaintArea.MouseIcon = frmControls.MyCursor(CurrentButton).MouseIcon
            frmMain.ActiveForm.PaintArea.MousePointer = 99
    End Select
    
    On Error Resume Next
    
    BarX = frmMain.TopBar.ScaleWidth - frmMain.CoordsInfo.Width
    
    If BarX < frmControls.DrawToolbar(CurrentButton).Left + frmControls.DrawToolbar(CurrentButton).Width Then
        BarX = frmControls.DrawToolbar(CurrentButton).Left + frmControls.DrawToolbar(CurrentButton).Width
    End If
    
    If CurrentButton <> 12 Then
        frmMain.ActiveForm.TextInput.Text = ""
        frmMain.ActiveForm.lblTextSize.Caption = "M"
        frmMain.ActiveForm.TextInput.Visible = False
    End If
    
    frmMain.CoordsInfo.Left = BarX
    frmControls.UnClickButton
    
End Sub


Public Sub SelectToolRect(i As Integer)
   
    frmControls.Cmd(CurrentButtonRect).BackColor = vb3DFace
    frmControls.Cmd(CurrentButtonRect).Line (0, 0)-(frmControls.Cmd(CurrentButtonRect).ScaleWidth - 1, frmControls.Cmd(CurrentButtonRect).ScaleHeight - 1), vb3DFace, B
    
    CheckedButtonRect(CurrentButtonRect) = False
    CheckedButtonRect(i) = True
    CurrentButtonRect = i
    
    
    frmControls.Cmd(i).BackColor = vbScrollBars
    frmControls.Cmd(i).Line (0, 0)-(frmControls.Cmd(i).ScaleWidth - 1, 0), vb3DShadow
    frmControls.Cmd(i).Line (0, 0)-(0, frmControls.Cmd(i).ScaleHeight - 1), vb3DShadow
    frmControls.Cmd(i).Line (frmControls.Cmd(i).ScaleWidth - 1, 1)-(frmControls.Cmd(i).ScaleWidth - 1, frmControls.Cmd(i).ScaleHeight), vb3DHighlight
    frmControls.Cmd(i).Line (0, frmControls.Cmd(i).ScaleHeight - 1)-(frmControls.Cmd(i).ScaleWidth - 1, frmControls.Cmd(i).ScaleHeight - 1), vb3DHighlight

    frmControls.Cmd2(i).BackColor = vbScrollBars
    frmControls.Cmd2(i).Line (0, 0)-(frmControls.Cmd(i).ScaleWidth - 1, 0), vb3DShadow
    frmControls.Cmd2(i).Line (0, 0)-(0, frmControls.Cmd(i).ScaleHeight - 1), vb3DShadow
    frmControls.Cmd2(i).Line (frmControls.Cmd(i).ScaleWidth - 1, 1)-(frmControls.Cmd(i).ScaleWidth - 1, frmControls.Cmd(i).ScaleHeight), vb3DHighlight
    frmControls.Cmd2(i).Line (0, frmControls.Cmd(i).ScaleHeight - 1)-(frmControls.Cmd(i).ScaleWidth - 1, frmControls.Cmd(i).ScaleHeight - 1), vb3DHighlight
    
End Sub

Public Sub SelectToolBrush(i As Integer)
   
    frmControls.Cmd3(CurrentButtonBrush).BackColor = vb3DFace
    frmControls.Cmd3(CurrentButtonBrush).Line (0, 0)-(frmControls.Cmd3(CurrentButtonBrush).ScaleWidth - 1, frmControls.Cmd3(CurrentButtonBrush).ScaleHeight - 1), vb3DFace, B
    
    CheckedButtonBrush(CurrentButtonBrush) = False
    CheckedButtonBrush(i) = True
    CurrentButtonBrush = i
    
    
    frmControls.Cmd3(i).BackColor = vbScrollBars
    frmControls.Cmd3(i).Line (0, 0)-(frmControls.Cmd3(i).ScaleWidth - 1, 0), vb3DShadow
    frmControls.Cmd3(i).Line (0, 0)-(0, frmControls.Cmd3(i).ScaleHeight - 1), vb3DShadow
    frmControls.Cmd3(i).Line (frmControls.Cmd3(i).ScaleWidth - 1, 1)-(frmControls.Cmd3(i).ScaleWidth - 1, frmControls.Cmd3(i).ScaleHeight), vb3DHighlight
    frmControls.Cmd3(i).Line (0, frmControls.Cmd3(i).ScaleHeight - 1)-(frmControls.Cmd3(i).ScaleWidth - 1, frmControls.Cmd3(i).ScaleHeight - 1), vb3DHighlight
    
End Sub

Public Sub SelectToolLight(i As Integer)
   
    frmControls.Cmd4(CurrentButtonLight).BackColor = vb3DFace
    frmControls.Cmd4(CurrentButtonLight).Line (0, 0)-(frmControls.Cmd4(CurrentButtonRect).ScaleWidth - 1, frmControls.Cmd4(CurrentButtonRect).ScaleHeight - 1), vb3DFace, B
    
    CheckedButtonLight(CurrentButtonLight) = False
    CheckedButtonLight(i) = True
    CurrentButtonLight = i
    
    
    frmControls.Cmd4(i).BackColor = vbScrollBars
    frmControls.Cmd4(i).Line (0, 0)-(frmControls.Cmd4(i).ScaleWidth - 1, 0), vb3DShadow
    frmControls.Cmd4(i).Line (0, 0)-(0, frmControls.Cmd4(i).ScaleHeight - 1), vb3DShadow
    frmControls.Cmd4(i).Line (frmControls.Cmd4(i).ScaleWidth - 1, 1)-(frmControls.Cmd4(i).ScaleWidth - 1, frmControls.Cmd4(i).ScaleHeight), vb3DHighlight
    frmControls.Cmd4(i).Line (0, frmControls.Cmd4(i).ScaleHeight - 1)-(frmControls.Cmd4(i).ScaleWidth - 1, frmControls.Cmd4(i).ScaleHeight - 1), vb3DHighlight
    
End Sub

Public Sub Init()
    Dim i As Integer
        
    frmSPlash.lblStatus.Caption = "Loading Colors..."
    DoEvents
    SetParent frmSelector.hwnd, frmMain.hwnd
    
    frmSPlash.lblStatus.Caption = "Loading Fonts..."
    DoEvents
    SetParent frmFonts.hwnd, frmMain.hwnd
    
    frmSPlash.lblStatus.Caption = "Loading Fontstyles..."
    DoEvents
    SetParent frmFontStyle.hwnd, frmMain.hwnd
    
    frmSPlash.lblStatus.Caption = "Loading Filters..."
    DoEvents
    LoadFilters

    frmSelector.Hide
    frmFonts.Hide
    frmFontStyle.Hide
    
    SetSelMethods
    
    frmSPlash.lblStatus.Caption = "preparing GUI..."
    DoEvents
    
    frmControls.CmdFont.BackColor = vb3DFace
    frmControls.CmdFont.Line (0, 0)-(frmControls.CmdFont.ScaleWidth, 0), vb3DHighlight
    frmControls.CmdFont.Line (0, 0)-(0, frmControls.CmdFont.ScaleHeight), vb3DHighlight
    frmControls.CmdFont.Line (frmControls.CmdFont.ScaleWidth - 1, 1)-(frmControls.CmdFont.ScaleWidth - 1, frmControls.CmdFont.ScaleHeight), vb3DShadow
    frmControls.CmdFont.Line (1, frmControls.CmdFont.ScaleHeight - 1)-(frmControls.CmdFont.ScaleWidth, frmControls.CmdFont.ScaleHeight - 1), vb3DShadow
    
    frmControls.CmdFontStyle.BackColor = vb3DFace
    frmControls.CmdFontStyle.Line (0, 0)-(frmControls.CmdFontStyle.ScaleWidth, 0), vb3DHighlight
    frmControls.CmdFontStyle.Line (0, 0)-(0, frmControls.CmdFontStyle.ScaleHeight), vb3DHighlight
    frmControls.CmdFontStyle.Line (frmControls.CmdFontStyle.ScaleWidth - 1, 1)-(frmControls.CmdFontStyle.ScaleWidth - 1, frmControls.CmdFontStyle.ScaleHeight), vb3DShadow
    frmControls.CmdFontStyle.Line (1, frmControls.CmdFontStyle.ScaleHeight - 1)-(frmControls.CmdFontStyle.ScaleWidth, frmControls.CmdFontStyle.ScaleHeight - 1), vb3DShadow
    
    For i = 0 To frmControls.CmdButton.UBound
        frmControls.CmdButton(i).BackColor = vb3DFace
        frmControls.CmdButton(i).Line (0, 0)-(frmControls.CmdButton(i).ScaleWidth, 0), vb3DHighlight
        frmControls.CmdButton(i).Line (0, 0)-(0, frmControls.CmdButton(i).ScaleHeight), vb3DHighlight
        frmControls.CmdButton(i).Line (frmControls.CmdButton(i).ScaleWidth - 1, 1)-(frmControls.CmdButton(i).ScaleWidth - 1, frmControls.CmdButton(i).ScaleHeight), vb3DShadow
        frmControls.CmdButton(i).Line (1, frmControls.CmdButton(i).ScaleHeight - 1)-(frmControls.CmdButton(i).ScaleWidth, frmControls.CmdButton(i).ScaleHeight - 1), vb3DShadow
    Next i
    
    frmSelector.Line (0, 0)-(frmSelector.ScaleWidth - 1, frmSelector.ScaleHeight - 1), vbWindowFrame, B
    frmSelector.Line (0, 0)-(frmSelector.ScaleWidth - 1, 0), vb3DHighlight
    frmSelector.Line (0, 0)-(0, frmSelector.ScaleHeight - 1), vb3DHighlight
    
    frmSelector.Line (2, frmSelector.ScaleHeight - 2)-(frmSelector.ScaleWidth - 1, frmSelector.ScaleHeight - 2), vb3DShadow
    frmSelector.Line (frmSelector.ScaleWidth - 2, 2)-(frmSelector.ScaleWidth - 2, frmSelector.ScaleHeight - 2), vb3DShadow
    
    frmSelector.ScrollBlock.Line (0, 0)-(frmSelector.ScrollBlock.ScaleWidth - 1, frmSelector.ScrollBlock.ScaleHeight - 1), vbWindowFrame, B
    frmSelector.ScrollBlock.Line (0, 0)-(frmSelector.ScrollBlock.ScaleWidth - 1, 0), vb3DHighlight
    frmSelector.ScrollBlock.Line (0, 0)-(0, frmSelector.ScrollBlock.ScaleHeight - 1), vb3DHighlight
    
    frmSelector.ScrollBlock.Line (2, frmSelector.ScrollBlock.ScaleHeight - 2)-(frmSelector.ScrollBlock.ScaleWidth - 1, frmSelector.ScrollBlock.ScaleHeight - 2), vb3DShadow
    frmSelector.ScrollBlock.Line (frmSelector.ScrollBlock.ScaleWidth - 2, 2)-(frmSelector.ScrollBlock.ScaleWidth - 2, frmSelector.ScrollBlock.ScaleHeight - 2), vb3DShadow
    
    
    frmMain.SelColor(0).BackColor = RGB(0, 0, 0)
    frmMain.SelColor(1).BackColor = RGB(255, 255, 255)
    
    DrawPreviewGradient
    
    For i = 0 To frmControls.CmdDropDown.UBound
        frmControls.DrawButton i, False
    Next i
    
    frmMain.ColorsBg.Line (0, 0)-(frmMain.ColorsBg.ScaleWidth, 0), vb3DShadow
    frmMain.ColorsBg.Line (0, 1)-(frmMain.ColorsBg.ScaleWidth, 1), vb3DHighlight
    
    frmMain.TopBar.Line (0, 0)-(frmMain.TopBar.ScaleWidth, 0), vb3DShadow
    frmMain.TopBar.Line (0, 1)-(frmMain.TopBar.ScaleWidth, 1), vb3DHighlight
    
    frmMain.CoordsInfo.Line (0, 0)-(0, frmMain.CoordsInfo.ScaleHeight), vb3DShadow
    frmMain.CoordsInfo.Line (1, 0)-(1, frmMain.CoordsInfo.ScaleHeight), vb3DHighlight
    
    For i = 0 To frmMain.Cmd.UBound
        frmControls.DrawToolbar(i).Move 5, 3
        SetParent frmControls.DrawToolbar(i).hwnd, frmMain.TopBar.hwnd
        frmControls.ToolIcon(i).Line (frmControls.ToolIcon(i).ScaleWidth - 2, 0)-(frmControls.ToolIcon(i).ScaleWidth - 2, frmControls.ToolIcon(i).ScaleHeight), vb3DShadow
        frmControls.ToolIcon(i).Line (frmControls.ToolIcon(i).ScaleWidth - 1, 0)-(frmControls.ToolIcon(i).ScaleWidth - 1, frmControls.ToolIcon(i).ScaleHeight), vb3DHighlight
        
    Next i
    
    For i = 0 To frmMain.SplitH.UBound
        frmMain.SplitH(i).Line (0, 0)-(frmMain.SplitH(i).ScaleWidth, 0), vb3DShadow
        frmMain.SplitH(i).Line (0, 1)-(frmMain.SplitH(i).ScaleWidth, 1), vb3DHighlight
    Next i
    
End Sub

Public Sub LoadSwatches(File As String)
    Dim Path As String
    Dim tempstr As String
    Dim i As Integer
    Dim X As Integer
    Dim Y As Integer
    
    Path = App.Path
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    Path = Path & "swatches\" & File
    
    Open Path For Input As #1
        Do While Not EOF(1)
            Line Input #1, tempstr
            If tempstr <> "" Then
                If i <> 0 Then
                    Load frmMain.Swatch(i)
                End If
                
                frmMain.Swatch(i).BackColor = CLng(Mid(tempstr, 1, InStr(1, tempstr, " ") - 1))
                frmMain.Swatch(i).ToolTipText = Trim(Mid(tempstr, InStr(1, tempstr, " ") + 1))
                
                frmMain.Swatch(i).Move X, Y
                frmMain.Swatch(i).ZOrder 0
                frmMain.Swatch(i).Visible = True
                
                X = X + frmMain.Swatch(i).Width - 1
                If (X + frmMain.Swatch(i).Width - 1) >= frmMain.SwatchScroll.Width Then
                    X = 0
                    Y = Y + frmMain.Swatch(i).Height - 1
                    
                End If
                frmMain.SwatchScroll.Height = (Y + frmMain.Swatch(0).Height * 3)
                i = i + 1
            End If
        Loop
        
    Close #1
    
    If frmMain.SwatchScroll.Height > frmMain.SwatchesBg.ScaleHeight Then
        frmMain.ScrollSwatch.Min = 0
        frmMain.ScrollSwatch.Max = frmMain.SwatchScroll.Height - frmMain.SwatchesBg.ScaleHeight
        frmMain.ScrollSwatch.SmallChange = frmMain.SwatchScroll.Height / (frmMain.Swatch(0).Height - 1)
        frmMain.ScrollSwatch.LargeChange = frmMain.SwatchScroll.Height - (((frmMain.SwatchScroll.Height - frmMain.SwatchScroll.ScaleHeight) / frmMain.SwatchScroll.Height) * frmMain.SwatchScroll.Height)
        frmMain.ScrollSwatch.Enabled = True
    Else
        frmMain.ScrollSwatch.Enabled = False
    End If
    
End Sub


Public Sub SaveSwatches(File As String)
    Dim Path As String
    Dim tempstr As String
    Dim i As Integer
    
    Path = App.Path
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    Path = Path & "swatches\" & File
    
    For i = 0 To frmMain.Swatch.UBound
        tempstr = tempstr & frmMain.Swatch(i).BackColor & " " & frmMain.Swatch(i).ToolTipText & vbCrLf
    Next i
    
    Kill Path
    Open Path For Output As #1
        Print #1, tempstr
    Close #1
    
End Sub


Public Sub DrawPreviewGradient()
    Dim C1(2) As Byte
    Dim C2(2) As Byte
    
    Dim W As Integer, H As Integer
    Dim i As Integer
    
    Dim RS As Double, GS As Double, BS As Double
    Dim r As Double, g As Double, b As Double
    
    Dim Red1 As Double, Blue1 As Double, Green1 As Double
    Dim Red2 As Double, Blue2 As Double, Green2 As Double
    
    W = frmMain.Gradient.ScaleWidth
    H = frmMain.Gradient.ScaleHeight

    b = frmMain.SelColor(0).BackColor \ 65536
    g = (frmMain.SelColor(0).BackColor - b * 65536) \ 256
    r = frmMain.SelColor(0).BackColor - b * 65536 - g * 256

    
    Red1 = r
    Green1 = g
    Blue1 = b
    
    Blue2 = frmMain.SelColor(1).BackColor \ 65536
    Green2 = (frmMain.SelColor(1).BackColor - Blue2 * 65536) \ 256
    Red2 = frmMain.SelColor(1).BackColor - Blue2 * 65536 - Green2 * 256
    
    'On Error Resume Next
    
    If Red1 <> Red2 Then
        RS = ((Red1 - Red2) / W)
    Else
        RS = 0
    End If
    
    If Green1 <> Green2 Then
        GS = ((Green1 - Green2) / W)
    Else
        GS = 0
    End If
    
    If Blue1 <> Blue2 Then
        BS = ((Blue1 - Blue2) / W)
    Else
        BS = 0
    End If
    
    For i = 0 To W
        r = r - RS
        g = g - GS
        b = b - BS
        
        If r < 0 Then r = 0
        If r > 255 Then r = 255
        If g < 0 Then g = 0
        If g > 255 Then g = 255
        If b < 0 Then b = 0
        If b > 255 Then b = 255
        
        frmMain.Gradient.Line (i, 0)-(i, 20), RGB(r, g, b), B
    Next i
    
    
End Sub


Public Sub DrawGradient(Dest As PictureBox, Color1 As Long, Color2 As Long)
    Dim C1(2) As Byte
    Dim C2(2) As Byte
    
    Dim W As Integer, H As Integer
    Dim i As Integer
    
    Dim RS As Double, GS As Double, BS As Double
    Dim r As Double, g As Double, b As Double
    
    Dim Red1 As Double, Blue1 As Double, Green1 As Double
    Dim Red2 As Double, Blue2 As Double, Green2 As Double
    
    W = Dest.ScaleWidth
    H = Dest.ScaleHeight

    b = Color1 \ 65536
    g = (Color1 - b * 65536) \ 256
    r = Color1 - b * 65536 - g * 256

    
    Red1 = r
    Green1 = g
    Blue1 = b
    
    Blue2 = Color2 \ 65536
    Green2 = (Color2 - Blue2 * 65536) \ 256
    Red2 = Color2 - Blue2 * 65536 - Green2 * 256
    
    'On Error Resume Next
    
    If Red1 <> Red2 Then
        RS = ((Red1 - Red2) / W)
    Else
        RS = 0
    End If
    
    If Green1 <> Green2 Then
        GS = ((Green1 - Green2) / W)
    Else
        GS = 0
    End If
    
    If Blue1 <> Blue2 Then
        BS = ((Blue1 - Blue2) / W)
    Else
        BS = 0
    End If
    
    For i = 0 To W
        r = r - RS
        g = g - GS
        b = b - BS
        
        If r < 0 Then r = 0
        If r > 255 Then r = 255
        If g < 0 Then g = 0
        If g > 255 Then g = 255
        If b < 0 Then b = 0
        If b > 255 Then b = 255
        
        Dest.Line (i, 0)-(i, 20), RGB(r, g, b), B
    Next i
    
    
End Sub
