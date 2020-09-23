Attribute VB_Name = "SimplePaints"
Option Explicit

Public OriginalSelX As Integer
Public OriginalSelY As Integer

Public PrevX As Integer
Public PrevY As Integer

Public Sub Draw(Area As PictureBox, x As Integer, Y As Integer, ZoomFactor As Double, MouseDown As Boolean)
    Dim X1 As Long, X2 As Long
    Dim Y1 As Long, Y2 As Long
    Dim i As Integer
    Dim NewColor As Long, OldColor As Long
    Dim r1 As Double, g1 As Double, b1 As Double
    Dim r2 As Double, g2 As Double, b2 As Double
    
    Dim Distance(3) As Double
    Dim MultiPlyer As Double
    Dim Angle As Double
    Dim GradientC As Double
    Dim Tx1 As Long, Tx2 As Long
    Dim Ty1 As Long, Ty2 As Long
    Dim MaxGradient As Double
    
    Dim TempRadius As Double
    Dim Done() As Boolean
    Dim Path As String
    
    
    'select the corrent color for the operation.
    Area.ForeColor = frmMain.SelColor(0).BackColor
    Area.Parent.BufferSelected.ForeColor = frmMain.SelColor(0).BackColor
    
    Select Case CurrentButton
        'Move selected area
        Case 0
            If Area.Parent.SelectArea.Visible = True Then
                If MouseDown = True Then
                    Area.Parent.SelectArea.Left = x * CInt(ZoomFactor / 100) - (Area.CurrentX * CInt(ZoomFactor / 100) - OriginalSelX)
                    Area.Parent.SelectArea.Top = Y * CInt(ZoomFactor / 100) - (Area.CurrentY * CInt(ZoomFactor / 100) - OriginalSelY)
                    
                    Area.Parent.SelectedBack.Refresh
                    BitBlt Area.hdc, Area.Parent.BufferSelected.Left, Area.Parent.BufferSelected.Top, Area.Parent.BufferSelected.Width, Area.Parent.BufferSelected.Height, Area.Parent.SelectedBack.hdc, 0, 0, vbSrcCopy
                    Area.Parent.PaintArea.Refresh

                    Area.Parent.BufferSelected.Move Area.Parent.SelectArea.Left * (100 / ZoomFactor), Area.Parent.SelectArea.Top * (100 / ZoomFactor)
                    
                    Area.Parent.SelectedBack.Left = Area.Parent.SelectArea.Left * (100 / ZoomFactor)
                    Area.Parent.SelectedBack.Top = Area.Parent.SelectArea.Top * (100 / ZoomFactor)
                    
                    BitBlt Area.Parent.SelectedBack.hdc, 0, 0, Area.Parent.BufferSelected.Width, Area.Parent.BufferSelected.Height, Area.Parent.Buffer.hdc, Area.Parent.BufferSelected.Left, Area.Parent.BufferSelected.Top, vbSrcCopy
                    
                    Area.Parent.SelectedBack.Refresh
                    Area.Parent.SelectedBack.CurrentX = Area.CurrentX - Area.Parent.SelectArea.Left * (100 / ZoomFactor)
                    Area.Parent.SelectedBack.CurrentY = Area.CurrentY - Area.Parent.SelectArea.Top * (100 / ZoomFactor)
                    
                    Area.Refresh
                    UpdateArea Area, x, Y, ZoomFactor
                Else

                    
                    Area.Parent.SelectedBack.Refresh
                    BitBlt Area.hdc, Area.Parent.BufferSelected.Left, Area.Parent.BufferSelected.Top, Area.Parent.BufferSelected.Width, Area.Parent.BufferSelected.Height, Area.Parent.SelectedBack.hdc, 0, 0, vbSrcCopy
                    Area.Parent.PaintArea.Refresh

                    
                    OriginalSelX = (Area.Parent.SelectArea.Left / 100) * 100
                    OriginalSelY = (Area.Parent.SelectArea.Top / 100) * 100

                    Area.Parent.BufferSelected.Move OriginalSelX * (100 / ZoomFactor), OriginalSelY * (100 / ZoomFactor)
                    
                    Area.Parent.SelectedBack.Left = Area.Parent.SelectArea.Left * (100 / ZoomFactor)
                    Area.Parent.SelectedBack.Top = Area.Parent.SelectArea.Top * (100 / ZoomFactor)
                    
                    BitBlt Area.Parent.SelectedBack.hdc, 0, 0, Area.Parent.BufferSelected.Width, Area.Parent.BufferSelected.Height, Area.Parent.Buffer.hdc, Area.Parent.BufferSelected.Left, Area.Parent.BufferSelected.Top, vbSrcCopy
                    
                    Area.Parent.SelectedBack.Refresh
                    Area.Parent.SelectedBack.CurrentX = Area.CurrentX - Area.Parent.SelectArea.Left * (100 / ZoomFactor)
                    Area.Parent.SelectedBack.CurrentY = Area.CurrentY - Area.Parent.SelectArea.Top * (100 / ZoomFactor)
                    
                    Area.Refresh
                    UpdateArea Area, x, Y, ZoomFactor
                End If
            End If
        'simple line.
        Case 1
            If CurrentButtonBrush = 0 Then
                Area.DrawWidth = SelMethod(4).Current
                Area.Parent.BufferSelected.DrawWidth = SelMethod(4).Current
                
                If (Area.CurrentX) <= ((x) + 0.5) And (Area.CurrentX) >= ((x) - 0.5) And (Area.CurrentY) <= ((Y) + 0.5) And (Area.CurrentY) >= ((Y) - 0.5) Then
                    If Area.Parent.SelectArea.Visible = False Then
                        If MouseDown = True Then Area.PSet (x, Y)
                    Else
                        If MouseDown = True Then
                            Area.Parent.BufferSelected.PSet (x - Area.Parent.BufferSelected.Left, Y - Area.Parent.BufferSelected.Top)
                        End If
                        
                        Area.Parent.BufferSelected.Refresh
                    End If
                Else
                    If Area.Parent.SelectArea.Visible = False Then
                        Area.Line -(x, Y)
                    Else
                        Area.Parent.BufferSelected.Line -(x - Area.Parent.BufferSelected.Left, Y - Area.Parent.BufferSelected.Top)
                        Area.Parent.BufferSelected.Refresh
                    End If
                End If
                Area.Refresh
                
                UpdateArea Area, x, Y, ZoomFactor
            ElseIf CurrentButtonBrush = 1 Then
                If MouseDown = True Then
                    Area.Parent.FollowLine.X1 = Area.CurrentX * (ZoomFactor / 100)
                    Area.Parent.FollowLine.X2 = x * (ZoomFactor / 100)
                    
                    Area.Parent.FollowLine.Y1 = Area.CurrentY * (ZoomFactor / 100)
                    Area.Parent.FollowLine.Y2 = Y * (ZoomFactor / 100)
                    Area.Parent.FollowLine.Visible = True
                Else
                    Area.Parent.FollowLine.Visible = False
                    Area.DrawWidth = SelMethod(4).Current
                    Area.Parent.BufferSelected.DrawWidth = SelMethod(4).Current
                    
                    If (Area.CurrentX) <= ((x) + 0.5) And (Area.CurrentX) >= ((x) - 0.5) And (Area.CurrentY) <= ((Y) + 0.5) And (Area.CurrentY) >= ((Y) - 0.5) Then
                        If Area.Parent.SelectArea.Visible = False Then
                            Area.PSet (x, Y)
                        Else
                            Area.Parent.BufferSelected.PSet (x - Area.Parent.BufferSelected.Left, Y - Area.Parent.BufferSelected.Top)
                            Area.Parent.BufferSelected.Refresh
                        End If
                    Else
                        If Area.Parent.SelectArea.Visible = False Then
                            Area.Line (Area.CurrentX, Area.CurrentY)-(x, Y)
                        Else
                            Area.Parent.BufferSelected.Line (Area.Parent.BufferSelected.CurrentX, Area.Parent.BufferSelected.CurrentY)-(x - Area.Parent.BufferSelected.Left, Y - Area.Parent.BufferSelected.Top)
                            Area.Parent.BufferSelected.Refresh
                        End If
                    End If
                    Area.Refresh
                    
                    UpdateArea Area, x, Y, ZoomFactor
                End If
            End If
            
        'floodfill
        Case 3
            Area.FillColor = frmMain.SelColor(0).BackColor
            Area.Parent.BufferSelected.FillColor = frmMain.SelColor(0).BackColor
            
            If Area.Parent.SelectArea.Visible = False Then
                ExtFloodFill Area.hdc, x, Y, GetPixel(Area.hdc, x, Y), 1
            Else
                ExtFloodFill Area.Parent.BufferSelected.hdc, x - Area.Parent.BufferSelected.Left, Y - Area.Parent.BufferSelected.Top, GetPixel(Area.Parent.BufferSelected.hdc, x - Area.Parent.BufferSelected.Left, Y - Area.Parent.BufferSelected.Top), 1
                Area.Parent.BufferSelected.Refresh
            End If
            
            Area.Refresh
                    
            UpdateArea Area, x, Y, ZoomFactor
        
        'select area
        Case 5
            Area.Parent.CropArea.Visible = False
            
            If MouseDown = True Then
                If x > Area.CurrentX Then
                    Area.Parent.SelectArea.Left = Area.CurrentX * (ZoomFactor / 100)
                    Area.Parent.SelectArea.Width = (x * (ZoomFactor / 100)) - (Area.CurrentX * (ZoomFactor / 100))
                Else
                    Area.Parent.SelectArea.Left = x * (ZoomFactor / 100)
                    Area.Parent.SelectArea.Width = (Area.CurrentX * (ZoomFactor / 100)) - (x * (ZoomFactor / 100))
                End If
                
                If Y > Area.CurrentY Then
                    Area.Parent.SelectArea.Top = Area.CurrentY * (ZoomFactor / 100)
                    Area.Parent.SelectArea.Height = (Y * (ZoomFactor / 100)) - (Area.CurrentY * (ZoomFactor / 100))
                Else
                    Area.Parent.SelectArea.Top = Y * (ZoomFactor / 100)
                    Area.Parent.SelectArea.Height = (Area.CurrentY * (ZoomFactor / 100)) - (Y * (ZoomFactor / 100))
                End If

                If Area.Parent.SelectArea.Width > Area.Parent.PaintArea.ScaleWidth - Area.Parent.SelectArea.Left Then Area.Parent.SelectArea.Width = Area.Parent.PaintArea.ScaleWidth - Area.Parent.SelectArea.Left
                If Area.Parent.SelectArea.Height > Area.Parent.PaintArea.ScaleHeight - Area.Parent.SelectArea.Top Then Area.Parent.SelectArea.Height = Area.Parent.PaintArea.ScaleHeight - Area.Parent.SelectArea.Top
                If Area.Parent.SelectArea.Left < 0 Then
                    Area.Parent.SelectArea.Width = Area.Parent.SelectArea.Width + Area.Parent.SelectArea.Left
                    Area.Parent.SelectArea.Left = 0
                End If
                If Area.Parent.SelectArea.Top < 0 Then
                    Area.Parent.SelectArea.Height = Area.Parent.SelectArea.Height + Area.Parent.SelectArea.Top
                    Area.Parent.SelectArea.Top = 0
                End If

                Area.Parent.SelectArea.Visible = True
            Else
                If CInt(Area.CurrentX) = CInt(x) And CInt(Area.CurrentY) = CInt(Y) Then
                    Area.Parent.SelectArea.Visible = False
                ElseIf (Area.CurrentX) <= ((x) + 1) And (Area.CurrentX) >= ((x) - 1) And (Area.CurrentY) <= ((Y) + 1) And (Area.CurrentY) >= ((Y) - 1) Then
                    Area.Parent.SelectArea.Visible = False
                Else
                    Area.Parent.BufferSelected.Cls
                    Area.Parent.BufferSelected.Picture = Nothing
                    
                    
                    Area.Parent.BufferSelected.Left = Area.Parent.SelectArea.Left * (100 / ZoomFactor)
                    Area.Parent.BufferSelected.Top = Area.Parent.SelectArea.Top * (100 / ZoomFactor)
                    Area.Parent.BufferSelected.Width = Area.Parent.SelectArea.Width * (100 / ZoomFactor)
                    Area.Parent.BufferSelected.Height = Area.Parent.SelectArea.Height * (100 / ZoomFactor)
                    BitBlt Area.Parent.BufferSelected.hdc, 0, 0, Area.Parent.BufferSelected.Width, Area.Parent.BufferSelected.Height, Area.Parent.Buffer.hdc, Area.Parent.SelectArea.Left * (100 / ZoomFactor), Area.Parent.SelectArea.Top * (100 / ZoomFactor), vbSrcCopy
                    Area.Parent.BufferSelected.Refresh
                    Area.Parent.BufferSelected.CurrentX = Area.CurrentX - Area.Parent.SelectArea.Left * (100 / ZoomFactor)
                    Area.Parent.BufferSelected.CurrentY = Area.CurrentY - Area.Parent.SelectArea.Top * (100 / ZoomFactor)
                    
                    Area.Parent.SelectedBack.Left = Area.Parent.SelectArea.Left * (100 / ZoomFactor)
                    Area.Parent.SelectedBack.Top = Area.Parent.SelectArea.Top * (100 / ZoomFactor)
                    Area.Parent.SelectedBack.Width = Area.Parent.SelectArea.Width * (100 / ZoomFactor)
                    Area.Parent.SelectedBack.Height = Area.Parent.SelectArea.Height * (100 / ZoomFactor)
                    Area.Parent.SelectedBack.BackColor = frmMain.SelColor(1).BackColor
                    
                    Area.Parent.SelectedBack.Refresh
                    Area.Parent.SelectedBack.CurrentX = Area.CurrentX - Area.Parent.SelectArea.Left * (100 / ZoomFactor)
                    Area.Parent.SelectedBack.CurrentY = Area.CurrentY - Area.Parent.SelectArea.Top * (100 / ZoomFactor)
                    
                    
                    OriginalSelX = Area.Parent.SelectArea.Left
                    OriginalSelY = Area.Parent.SelectArea.Top
                End If
            End If
            
        'Air brush.
        Case 6
            b1 = frmMain.SelColor(0).BackColor \ 65536
            g1 = (frmMain.SelColor(0).BackColor - b1 * 65536) \ 256
            r1 = frmMain.SelColor(0).BackColor - b1 * 65536 - g1 * 256
            
            If Area.Parent.SelectArea.Visible = False Then
                If MouseDown = True Then
                    DrawAirBrush Area.hdc, CInt(x), CInt(Y), CByte(r1), CByte(g1), CByte(b1), CLng(SelMethod(2).Current), CLng(SelMethod(2).Current) / 2
                End If
            Else
                If MouseDown = True Then
                    DrawAirBrush Area.Parent.BufferSelected.hdc, CInt(x) - Area.Parent.BufferSelected.Left, CInt(Y) - Area.Parent.BufferSelected.Top, CByte(r1), CByte(g1), CByte(b1), CLng(SelMethod(2).Current), CLng(SelMethod(2).Current) / 2
                End If
            End If
            
            Area.Refresh
            UpdateArea Area, x, Y, ZoomFactor
            
            
            
        'select area for cropping
        Case 7
            Area.Parent.SelectArea.Visible = False
            
            If MouseDown = True Then
                If x > Area.CurrentX Then
                    Area.Parent.CropArea.Left = Area.CurrentX * (ZoomFactor / 100)
                    Area.Parent.CropArea.Width = (x * (ZoomFactor / 100)) - (Area.CurrentX * (ZoomFactor / 100))
                Else
                    Area.Parent.CropArea.Left = x * (ZoomFactor / 100)
                    Area.Parent.CropArea.Width = (Area.CurrentX * (ZoomFactor / 100)) - (x * (ZoomFactor / 100))
                End If
                
                If Y > Area.CurrentY Then
                    Area.Parent.CropArea.Top = Area.CurrentY * (ZoomFactor / 100)
                    Area.Parent.CropArea.Height = (Y * (ZoomFactor / 100)) - (Area.CurrentY * (ZoomFactor / 100))
                Else
                    Area.Parent.CropArea.Top = Y * (ZoomFactor / 100)
                    Area.Parent.CropArea.Height = (Area.CurrentY * (ZoomFactor / 100)) - (Y * (ZoomFactor / 100))
                End If
                
                If Area.Parent.CropArea.Width > Area.Parent.PaintArea.ScaleWidth - Area.Parent.CropArea.Left Then Area.Parent.CropArea.Width = Area.Parent.PaintArea.ScaleWidth - Area.Parent.CropArea.Left
                If Area.Parent.CropArea.Height > Area.Parent.PaintArea.ScaleHeight - Area.Parent.CropArea.Top Then Area.Parent.CropArea.Height = Area.Parent.PaintArea.ScaleHeight - Area.Parent.CropArea.Top
                If Area.Parent.CropArea.Left < 0 Then
                    Area.Parent.CropArea.Width = Area.Parent.CropArea.Width + Area.Parent.CropArea.Left
                    Area.Parent.CropArea.Left = 0
                End If
                If Area.Parent.CropArea.Top < 0 Then
                    Area.Parent.CropArea.Height = Area.Parent.CropArea.Height + Area.Parent.CropArea.Top
                    Area.Parent.CropArea.Top = 0
                End If
                
                 Area.Parent.CropArea.Visible = True
            Else
                If (Area.CurrentX) <= ((x) + 1) And (Area.CurrentX) >= ((x) - 1) And (Area.CurrentY) <= ((Y) + 1) And (Area.CurrentY) >= ((Y) - 1) Then
                    Area.Parent.CropArea.Visible = False
                Else
                    X1 = MsgBox("Are you sure you wish to crop?", vbExclamation + vbYesNo, "Crop?")
                    If X1 = 6 Then
                        Area.Parent.BufferSelected.Picture = LoadPicture()

                        Area.Parent.BufferSelected.Left = Area.Parent.CropArea.Left * (100 / ZoomFactor)
                        Area.Parent.BufferSelected.Top = Area.Parent.CropArea.Top * (100 / ZoomFactor)
                        Area.Parent.BufferSelected.Width = Area.Parent.CropArea.Width * (100 / ZoomFactor)
                        Area.Parent.BufferSelected.Height = Area.Parent.CropArea.Height * (100 / ZoomFactor)

                        BitBlt Area.Parent.BufferSelected.hdc, 0, 0, Area.Parent.BufferSelected.Width, Area.Parent.BufferSelected.Height, Area.Parent.Buffer.hdc, Area.Parent.CropArea.Left * (100 / ZoomFactor), Area.Parent.CropArea.Top * (100 / ZoomFactor), vbSrcCopy
                        Area.Parent.BufferSelected.Refresh
                        
                        
                        frmMain.ActiveForm.PaintArea.Width = Area.Parent.CropArea.Width + 2
                        frmMain.ActiveForm.PaintArea.Height = Area.Parent.CropArea.Height + 2
                        frmMain.ActiveForm.Buffer.Cls
                        frmMain.ActiveForm.Buffer.Width = Area.Parent.CropArea.Width * (100 / ZoomFactor) + 2
                        frmMain.ActiveForm.Buffer.Height = Area.Parent.CropArea.Height * (100 / ZoomFactor) + 2
                        
                        BitBlt frmMain.ActiveForm.Buffer.hdc, 0, 0, Area.Parent.BufferSelected.Width, Area.Parent.BufferSelected.Height, Area.Parent.BufferSelected.hdc, 0, 0, vbSrcCopy
                        
                        UpdateArea Area, x, Y, ZoomFactor
                        Area.Parent.CropArea.Visible = False
                        frmMain.ActiveForm.RealignZoom CDbl(x), CDbl(Y)
                    Else
                        Area.Parent.CropArea.Visible = False
                    End If
                End If
            End If
        
        'Get color
        Case 8
            If x >= 0 And Y >= 0 And Y < Area.ScaleHeight And x < Area.ScaleWidth Then
                On Error Resume Next
                frmMain.SelColor(frmMain.GetSelectedColor).BackColor = GetPixel(Area.hdc, x, Y)
                frmMain.ColorBlend(0).BackColor = frmMain.SelColor(0).BackColor
                frmMain.ColorBlend(1).BackColor = frmMain.SelColor(1).BackColor
                frmMain.SetColorBars
                DrawPreviewGradient
                
                On Error Resume Next
                frmMain.ActiveForm.Buffer.ForeColor = frmMain.SelColor(0).BackColor
                frmMain.ActiveForm.TextInput.ForeColor = frmMain.SelColor(0).BackColor
                frmMain.ActiveForm.TextInput.BackColor = frmMain.SelColor(1).BackColor
            End If
            
        'Gradient fill
        Case 4
            If MouseDown = True Then
                Area.Parent.FollowLine.X1 = Area.CurrentX * (ZoomFactor / 100)
                Area.Parent.FollowLine.X2 = x * (ZoomFactor / 100)
                
                Area.Parent.FollowLine.Y1 = Area.CurrentY * (ZoomFactor / 100)
                Area.Parent.FollowLine.Y2 = Y * (ZoomFactor / 100)
                Area.Parent.FollowLine.Visible = True
                
                Angle = GetAngleFromCoords(CInt(Area.Parent.FollowLine.X1) * (ZoomFactor / 100), CInt(Area.Parent.FollowLine.Y1) * (ZoomFactor / 100), CInt(Area.Parent.FollowLine.X2) * (ZoomFactor / 100), CInt(Area.Parent.FollowLine.Y2) * (ZoomFactor / 100))
                frmControls.lblAngle.Alignment = 1
                frmControls.lblAngle.Caption = CInt(Angle) & " ° "
                
                Distance(0) = (Sqr(((Area.CurrentX - 0) ^ 2) + ((Area.CurrentY - 0) ^ 2)))
                Distance(1) = (Sqr(((Area.CurrentX - Area.ScaleWidth) ^ 2) + ((Area.CurrentY - 0) ^ 2)))
                Distance(2) = (Sqr(((Area.CurrentX - 0) ^ 2) + ((Area.CurrentY - Area.ScaleHeight) ^ 2)))
                Distance(3) = (Sqr(((Area.CurrentX - Area.ScaleWidth) ^ 2) + ((Area.CurrentY - Area.ScaleHeight) ^ 2)))
                
                For i = 0 To 3
                    If MultiPlyer < Distance(i) Then MultiPlyer = Distance(i)
                Next i
                
                'MultiPlyer = 100
                
                Area.Parent.Temp.X1 = Area.CurrentX * (ZoomFactor / 100) - Sin(Angle * (3.14159265358979 / 180)) * MultiPlyer
                Area.Parent.Temp.Y1 = Area.CurrentY * (ZoomFactor / 100) - Cos(Angle * (3.14159265358979 / 180)) * MultiPlyer
                
                Area.Parent.Temp.X2 = Area.CurrentX * (ZoomFactor / 100) + Sin(Angle * (3.14159265358979 / 180)) * MultiPlyer
                Area.Parent.Temp.Y2 = Area.CurrentY * (ZoomFactor / 100) + Cos(Angle * (3.14159265358979 / 180)) * MultiPlyer
                
                Area.Parent.Temp2.X1 = x * (ZoomFactor / 100) - Sin(Angle * (3.14159265358979 / 180)) * MultiPlyer
                Area.Parent.Temp2.Y1 = Y * (ZoomFactor / 100) - Cos(Angle * (3.14159265358979 / 180)) * MultiPlyer
                
                Area.Parent.Temp2.X2 = x * (ZoomFactor / 100) + Sin(Angle * (3.14159265358979 / 180)) * MultiPlyer
                Area.Parent.Temp2.Y2 = Y * (ZoomFactor / 100) + Cos(Angle * (3.14159265358979 / 180)) * MultiPlyer
            Else
                frmControls.lblAngle.Alignment = 2
                frmControls.lblAngle.Caption = "N/A"
                Area.DrawWidth = 3
                Area.Parent.BufferSelected.DrawWidth = 3
                Angle = GetAngleFromCoords(CInt(Area.Parent.FollowLine.X1) * (ZoomFactor / 100), CInt(Area.Parent.FollowLine.Y1) * (ZoomFactor / 100), CInt(Area.Parent.FollowLine.X2) * (ZoomFactor / 100), CInt(Area.Parent.FollowLine.Y2) * (ZoomFactor / 100))
                Area.Parent.FollowLine.Visible = False
                
                MaxGradient = (Sqr(((Area.Parent.Temp.X1 * (100 / ZoomFactor) - Area.Parent.Temp2.X1 * (100 / ZoomFactor)) ^ 2) + ((Area.Parent.Temp.Y1 * (100 / ZoomFactor) - Area.Parent.Temp2.Y1 * (100 / ZoomFactor)) ^ 2)))
                
                For GradientC = 0 To MaxGradient Step 0.5
                    
                    X1 = Area.Parent.Temp.X1 * (100 / ZoomFactor) + Cos(-Angle * (3.14159265358979 / 180)) * GradientC
                    Y1 = Area.Parent.Temp.Y1 * (100 / ZoomFactor) + Sin(-Angle * (3.14159265358979 / 180)) * GradientC
                    
                    X2 = Area.Parent.Temp.X2 * (100 / ZoomFactor) + Cos(-Angle * (3.14159265358979 / 180)) * GradientC
                    Y2 = Area.Parent.Temp.Y2 * (100 / ZoomFactor) + Sin(-Angle * (3.14159265358979 / 180)) * GradientC

                    If Tx1 = X1 And Ty1 = Y1 Then
                    
                    ElseIf X1 < -4 And X2 < -4 Then
                    
                    ElseIf X1 > Area.ScaleWidth + 4 And X2 > Area.ScaleWidth + 4 Then
                    
                    Else
                    
                        If Area.Parent.SelectArea.Visible = False Then
                            Area.Line (X1, Y1)-(X2, Y2), GetGradientColor(CLng(MaxGradient), CLng(GradientC))
                        Else
                            Area.Parent.BufferSelected.Line (X1 - Area.Parent.BufferSelected.Left, Y1 - Area.Parent.BufferSelected.Top)-(X2 - Area.Parent.BufferSelected.Left, Y2 - Area.Parent.BufferSelected.Top), GetGradientColor(CLng(MaxGradient), CLng(GradientC))
                            Area.Parent.BufferSelected.Refresh
                        End If
                        
                        Tx1 = X1
                        Ty1 = Y1
                    End If
                    
                Next GradientC
                
                Area.Refresh
                UpdateArea Area, x, Y, ZoomFactor
            End If
            
        'Draw Square
        Case 9
            If MouseDown = True Then
                If x > Area.CurrentX Then
                    Area.Parent.DrawBox.Left = Area.CurrentX * (ZoomFactor / 100)
                    Area.Parent.DrawBox.Width = (x * (ZoomFactor / 100)) - (Area.CurrentX * (ZoomFactor / 100))
                Else
                    Area.Parent.DrawBox.Left = x * (ZoomFactor / 100)
                    Area.Parent.DrawBox.Width = (Area.CurrentX * (ZoomFactor / 100)) - (x * (ZoomFactor / 100))
                End If
                
                If Y > Area.CurrentY Then
                    Area.Parent.DrawBox.Top = Area.CurrentY * (ZoomFactor / 100)
                    Area.Parent.DrawBox.Height = (Y * (ZoomFactor / 100)) - (Area.CurrentY * (ZoomFactor / 100))
                Else
                    Area.Parent.DrawBox.Top = Y * (ZoomFactor / 100)
                    Area.Parent.DrawBox.Height = (Area.CurrentY * (ZoomFactor / 100)) - (Y * (ZoomFactor / 100))
                End If
                
                Area.Parent.DrawBox.Visible = True
            Else
                Area.Parent.DrawBox.Visible = False
                
                Select Case CurrentButtonRect
                    Case 0
                        Area.FillStyle = 1
                        Area.DrawWidth = SelMethod(5).Current
                        Area.ForeColor = frmMain.SelColor(0).BackColor
                        
                        Area.Parent.BufferSelected.FillStyle = 1
                        Area.Parent.BufferSelected.DrawWidth = SelMethod(5).Current
                        Area.Parent.BufferSelected.ForeColor = frmMain.SelColor(0).BackColor
                    Case 1
                        Area.FillStyle = 0
                        Area.DrawWidth = SelMethod(5).Current
                        Area.ForeColor = frmMain.SelColor(0).BackColor
                        Area.FillColor = frmMain.SelColor(1).BackColor
                        
                        Area.Parent.BufferSelected.FillStyle = 0
                        Area.Parent.BufferSelected.DrawWidth = SelMethod(5).Current
                        Area.Parent.BufferSelected.ForeColor = frmMain.SelColor(0).BackColor
                        Area.Parent.BufferSelected.FillColor = frmMain.SelColor(1).BackColor
                    Case 2
                        Area.FillStyle = 0
                        Area.DrawWidth = 1
                        Area.ForeColor = frmMain.SelColor(0).BackColor
                        Area.FillColor = frmMain.SelColor(0).BackColor
                        
                        Area.Parent.BufferSelected.FillStyle = 0
                        Area.Parent.BufferSelected.DrawWidth = 1
                        Area.Parent.BufferSelected.ForeColor = frmMain.SelColor(0).BackColor
                        Area.Parent.BufferSelected.FillColor = frmMain.SelColor(0).BackColor
                End Select
                
                If Area.Parent.SelectArea.Visible = False Then
                    Area.Line (Area.CurrentX, Area.CurrentY)-(x, Y), frmMain.SelColor(0).BackColor, B
                    Area.Refresh
                Else
                    Area.Parent.BufferSelected.Line (Area.CurrentX - Area.Parent.BufferSelected.Left, Area.CurrentY - Area.Parent.BufferSelected.Top)-(x - Area.Parent.BufferSelected.Left, Y - Area.Parent.BufferSelected.Top), frmMain.SelColor(0).BackColor, B
                    Area.Parent.BufferSelected.Refresh
                End If
                
                UpdateArea Area, x, Y, ZoomFactor
            End If
            
        'Draw Circle
        Case 10
            If MouseDown = True Then
                If x > Area.CurrentX Then
                    Area.Parent.DrawCircle.Left = Area.CurrentX * (ZoomFactor / 100)
                    Area.Parent.DrawCircle.Width = (x * (ZoomFactor / 100)) - (Area.CurrentX * (ZoomFactor / 100))
                Else
                    Area.Parent.DrawCircle.Left = x * (ZoomFactor / 100)
                    Area.Parent.DrawCircle.Width = (Area.CurrentX * (ZoomFactor / 100)) - (x * (ZoomFactor / 100))
                End If
                
                If Y > Area.CurrentY Then
                    Area.Parent.DrawCircle.Top = Area.CurrentY * (ZoomFactor / 100)
                    Area.Parent.DrawCircle.Height = (Y * (ZoomFactor / 100)) - (Area.CurrentY * (ZoomFactor / 100))
                Else
                    Area.Parent.DrawCircle.Top = Y * (ZoomFactor / 100)
                    Area.Parent.DrawCircle.Height = (Area.CurrentY * (ZoomFactor / 100)) - (Y * (ZoomFactor / 100))
                End If
                
                Area.Parent.DrawCircle.Visible = True
            Else
                Area.Parent.DrawCircle.Visible = False
                
                Select Case CurrentButtonRect
                    Case 0
                        Area.FillStyle = 1
                        Area.DrawWidth = SelMethod(6).Current
                        Area.ForeColor = frmMain.SelColor(0).BackColor
                        
                        Area.Parent.BufferSelected.FillStyle = 1
                        Area.Parent.BufferSelected.DrawWidth = SelMethod(6).Current
                        Area.Parent.BufferSelected.ForeColor = frmMain.SelColor(0).BackColor
                    Case 1
                        Area.FillStyle = 0
                        Area.DrawWidth = SelMethod(6).Current
                        Area.ForeColor = frmMain.SelColor(0).BackColor
                        Area.FillColor = frmMain.SelColor(1).BackColor
                        
                        Area.Parent.BufferSelected.FillStyle = 0
                        Area.Parent.BufferSelected.DrawWidth = SelMethod(6).Current
                        Area.Parent.BufferSelected.ForeColor = frmMain.SelColor(0).BackColor
                        Area.Parent.BufferSelected.FillColor = frmMain.SelColor(1).BackColor
                    Case 2
                        Area.FillStyle = 0
                        Area.DrawWidth = 1
                        Area.ForeColor = frmMain.SelColor(0).BackColor
                        Area.FillColor = frmMain.SelColor(0).BackColor
                        
                        Area.Parent.BufferSelected.FillStyle = 0
                        Area.Parent.BufferSelected.DrawWidth = 1
                        Area.Parent.BufferSelected.ForeColor = frmMain.SelColor(0).BackColor
                        Area.Parent.BufferSelected.FillColor = frmMain.SelColor(0).BackColor
                End Select
                
                If Area.Parent.SelectArea.Visible = False Then
                    RoundRect Area.hdc, _
                                Area.Parent.DrawCircle.Left * (100 / ZoomFactor), _
                                Area.Parent.DrawCircle.Top * (100 / ZoomFactor), _
                                Area.Parent.DrawCircle.Left * (100 / ZoomFactor) + Area.Parent.DrawCircle.Width * (100 / ZoomFactor), _
                                Area.Parent.DrawCircle.Top * (100 / ZoomFactor) + Area.Parent.DrawCircle.Height * (100 / ZoomFactor), _
                                Area.Parent.DrawCircle.Left * (100 / ZoomFactor) + Area.Parent.DrawCircle.Width * (100 / ZoomFactor), _
                                Area.Parent.DrawCircle.Top * (100 / ZoomFactor) + Area.Parent.DrawCircle.Height * (100 / ZoomFactor)
                    Area.Refresh
                Else
                    RoundRect Area.Parent.BufferSelected.hdc, _
                                (Area.Parent.DrawCircle.Left * (100 / ZoomFactor)) - Area.Parent.BufferSelected.Left, _
                                (Area.Parent.DrawCircle.Top * (100 / ZoomFactor)) - Area.Parent.BufferSelected.Top, _
                                (Area.Parent.DrawCircle.Left * (100 / ZoomFactor) + Area.Parent.DrawCircle.Width * (100 / ZoomFactor)) - Area.Parent.BufferSelected.Left, _
                                (Area.Parent.DrawCircle.Top * (100 / ZoomFactor) + Area.Parent.DrawCircle.Height * (100 / ZoomFactor)) - Area.Parent.BufferSelected.Top, _
                                Area.Parent.DrawCircle.Left * (100 / ZoomFactor) + Area.Parent.DrawCircle.Width * (100 / ZoomFactor), _
                                Area.Parent.DrawCircle.Top * (100 / ZoomFactor) + Area.Parent.DrawCircle.Height * (100 / ZoomFactor)
                    Area.Parent.BufferSelected.Refresh
                End If
                UpdateArea Area, x, Y, ZoomFactor
            End If
            
            
        'Lighten or darken.
        Case 11
            b1 = frmMain.SelColor(0).BackColor \ 65536
            g1 = (frmMain.SelColor(0).BackColor - b1 * 65536) \ 256
            r1 = frmMain.SelColor(0).BackColor - b1 * 65536 - g1 * 256
            
            If Area.Parent.SelectArea.Visible = False Then
                If MouseDown = True Then
                    If CurrentButtonLight = 0 Then
                        DrawLightDark Area.hdc, CInt(x), CInt(Y), CByte(r1), CByte(g1), CByte(b1), CLng(SelMethod(3).Current), CLng(SelMethod(3).Current) / 2, False
                    Else
                        DrawLightDark Area.hdc, CInt(x), CInt(Y), CByte(r1), CByte(g1), CByte(b1), CLng(SelMethod(3).Current), CLng(SelMethod(3).Current) / 2, True
                    End If
                End If
            Else
                If MouseDown = True Then
                    If CurrentButtonLight = 0 Then
                        DrawLightDark Area.Parent.BufferSelected.hdc, CInt(x) - Area.Parent.BufferSelected.Left, CInt(Y) - Area.Parent.BufferSelected.Top, CByte(r1), CByte(g1), CByte(b1), CLng(SelMethod(3).Current), CLng(SelMethod(3).Current) / 2, False
                    Else
                        DrawLightDark Area.Parent.BufferSelected.hdc, CInt(x) - Area.Parent.BufferSelected.Left, CInt(Y) - Area.Parent.BufferSelected.Top, CByte(r1), CByte(g1), CByte(b1), CLng(SelMethod(3).Current), CLng(SelMethod(3).Current) / 2, True
                    End If
                End If
            End If
            
            Area.Refresh
            UpdateArea Area, x, Y, ZoomFactor
        
        'textbox...
        Case 12
            'Area.Parent.TextInput.Text = ""
            Area.Parent.SelectArea.Visible = False
            Area.font = Mid(frmControls.lblFont.Caption, 2)
            Area.Parent.TextInput.font = Mid(frmControls.lblFont.Caption, 2)
            Area.Parent.lblTextSize.font = Mid(frmControls.lblFont.Caption, 2)
            
            Area.FontSize = SelMethod(7).Current
            Area.Parent.TextInput.FontSize = SelMethod(7).Current * (ZoomFactor / 100)
            Area.ForeColor = frmMain.SelColor(0).BackColor
            Area.Parent.TextInput.ForeColor = frmMain.SelColor(0).BackColor
            Area.Parent.TextInput.BackColor = frmMain.SelColor(1).BackColor
            Area.Parent.lblTextSize.FontSize = SelMethod(7).Current * (ZoomFactor / 100)
            Area.Parent.lblTextSize.Caption = Area.Parent.TextInput.Text & "M"
            Area.Parent.TextInput.Move x * (ZoomFactor / 100), Y * (ZoomFactor / 100), Area.Parent.lblTextSize.Width, Area.Parent.lblTextSize.Height
            Area.Parent.TextInput.Visible = True
            Area.Parent.TextInput.SetFocus
            
            
        'Blur
        Case 13
            If Area.Parent.SelectArea.Visible = False Then
                If MouseDown = True Then
                    DrawBlur Area.hdc, CInt(x), CInt(Y), CLng(SelMethod(8).Current), CLng(SelMethod(8).Current) / 2
                End If
            Else
                If MouseDown = True Then
                    DrawBlur Area.Parent.BufferSelected.hdc, CInt(x) - Area.Parent.BufferSelected.Left, CInt(Y) - Area.Parent.BufferSelected.Top, CLng(SelMethod(8).Current), CLng(SelMethod(8).Current) / 2
                End If
            End If
            
            Area.Refresh
            UpdateArea Area, x, Y, ZoomFactor
    End Select
End Sub

'update to screen...
Public Sub UpdateArea(Area As PictureBox, x As Integer, Y As Integer, ZoomFactor As Double)
    If Area.Parent.SelectArea.Visible = True Then
        BitBlt Area.hdc, Area.Parent.BufferSelected.Left, Area.Parent.BufferSelected.Top, Area.Parent.BufferSelected.Width, Area.Parent.BufferSelected.Height, Area.Parent.BufferSelected.hdc, 0, 0, vbSrcCopy
    End If
    
    If ZoomFactor = 100 Then
        BitBlt Area.Parent.PaintArea.hdc, 0, 0, Area.Parent.PaintArea.ScaleWidth, Area.Parent.PaintArea.ScaleHeight, Area.hdc, 0, 0, vbSrcCopy
        Area.Parent.PaintArea.Refresh
    Else
        StretchBlt Area.Parent.PaintArea.hdc, Area.Parent.HScroll1.Value, Area.Parent.VScroll1.Value, Area.Parent.Back.ScaleWidth, Area.Parent.Back.ScaleHeight, _
                   Area.hdc, Area.Parent.HScroll1.Value / ZoomFactor * 100, Area.Parent.VScroll1.Value / ZoomFactor * 100, Area.Parent.Back.ScaleWidth / ZoomFactor * 100, Area.Parent.Back.ScaleHeight / ZoomFactor * 100, vbSrcCopy
        
        Area.Parent.PaintArea.Refresh
    End If
End Sub

Public Function GetAngleFromCoords(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double)
    Dim Angle As Double

    If X2 <> X1 Then
        Angle = Abs((Atn((Y2 - Y1) / (X2 - X1))) * (180 / 3.14159265358979))
    Else
        Angle = 90
    End If
    
    If X1 > X2 And Y1 > Y2 Then
        Angle = 90 + (90 - Angle)
    ElseIf X1 > X2 And Y1 < Y2 Then
        Angle = Angle + 180
    ElseIf X1 < X2 And Y1 < Y2 Then
        Angle = 270 + (90 - Angle)
    
    ElseIf X1 > X2 And Y1 = Y2 Then
        Angle = 180
    ElseIf X1 < X2 And Y1 = Y2 Then
        Angle = 0
    ElseIf X1 = X2 And Y1 > Y2 Then
        Angle = 90
    ElseIf X1 = X2 And Y1 < Y2 Then
        Angle = 270
    End If
    
    GetAngleFromCoords = Angle
    
End Function

Public Function GetGradientColor(Max As Long, Position As Long) As Long
    Dim C1(2) As Byte
    Dim C2(2) As Byte

    Dim i As Integer
    
    Dim RS As Double, GS As Double, BS As Double
    Dim r As Double, g As Double, b As Double
    
    Dim Red1 As Double, Blue1 As Double, Green1 As Double
    Dim Red2 As Double, Blue2 As Double, Green2 As Double

    If Max <= 0 Then
        GetGradientColor = frmMain.SelColor(0).BackColor
        Exit Function
    End If

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
        RS = ((Red1 - Red2) / Max)
    Else
        RS = 0
    End If
    
    If Green1 <> Green2 Then
        GS = ((Green1 - Green2) / Max)
    Else
        GS = 0
    End If
    
    If Blue1 <> Blue2 Then
        BS = ((Blue1 - Blue2) / Max)
    Else
        BS = 0
    End If
    

    r = r - RS * Position
    g = g - GS * Position
    b = b - BS * Position
    
    If r < 0 Then r = 0
    If r > 255 Then r = 255
    If g < 0 Then g = 0
    If g > 255 Then g = 255
    If b < 0 Then b = 0
    If b > 255 Then b = 255
    
    GetGradientColor = RGB(r, g, b)

End Function


Public Sub DrawAirBrush(Target As Long, x As Single, Y As Single, RedB As Byte, GreenB As Byte, BlueB As Byte, Radius As Long, NumberOfSteps As Long)
    Dim cx As Long 'X counter
    Dim cy As Long 'Y counter
    Dim TempColor As Long
    Dim TempRadius As Integer
    Dim i As Integer
    Dim Red As Long
    Dim Green As Long
    Dim Blue As Long
    
    Dim Done() As Boolean
    ReDim Done(-Radius To Radius, -Radius To Radius)
    
    For i = 1 To NumberOfSteps
        TempRadius = Radius / NumberOfSteps * i
        
        For cx = -TempRadius To TempRadius
            For cy = -TempRadius To TempRadius
                If Not Done(cx, cy) Then
                    If (cx * cx) + (cy * cy) <= TempRadius * TempRadius Then
                        TempColor = GetPixel(Target, cx + x, cy + Y)
                        
                        SetPixelV Target, cx + x, cy + Y, GetAirColor(TempColor, NumberOfSteps * (11 - SelMethod(0).Current / 10), NumberOfSteps * (10 - SelMethod(0).Current / 10) + CLng(i))
                        
                        Done(cx, cy) = True
                    End If
                End If
            Next cy
        Next cx
    Next i
End Sub

Public Function GetAirColor(Color As Long, Max As Long, Position As Long) As Long
    Dim C1(2) As Byte
    Dim C2(2) As Byte

    Dim i As Integer
    
    Dim RS As Double, GS As Double, BS As Double
    Dim r As Double, g As Double, b As Double
    
    Dim Red1 As Double, Blue1 As Double, Green1 As Double
    Dim Red2 As Double, Blue2 As Double, Green2 As Double

    If Max <= 0 Then
        GetAirColor = Color
        Exit Function
    End If


    b = frmMain.SelColor(0).BackColor \ 65536
    g = (frmMain.SelColor(0).BackColor - b * 65536) \ 256
    r = frmMain.SelColor(0).BackColor - b * 65536 - g * 256

    
    Red1 = r
    Green1 = g
    Blue1 = b
    
    Blue2 = Color \ 65536
    Green2 = (Color - Blue2 * 65536) \ 256
    Red2 = Color - Blue2 * 65536 - Green2 * 256
    
    If Red1 <> Red2 Then
        RS = ((Red1 - Red2) / Max)
    Else
        RS = 0
    End If
    
    If Green1 <> Green2 Then
        GS = ((Green1 - Green2) / Max)
    Else
        GS = 0
    End If
    
    If Blue1 <> Blue2 Then
        BS = ((Blue1 - Blue2) / Max)
    Else
        BS = 0
    End If
    

    r = r - RS * (Position)
    g = g - GS * (Position)
    b = b - BS * (Position)
    
    If r < 0 Then r = 0
    If r > 255 Then r = 255
    If g < 0 Then g = 0
    If g > 255 Then g = 255
    If b < 0 Then b = 0
    If b > 255 Then b = 255
       
    GetAirColor = RGB(CInt(r), CInt(g), CInt(b))

End Function


Public Sub DrawLightDark(Target As Long, x As Single, Y As Single, RedB As Byte, GreenB As Byte, BlueB As Byte, Radius As Long, NumberOfSteps As Long, Darken As Boolean)
    Dim cx As Long 'X counter
    Dim cy As Long 'Y counter
    Dim TempColor As Long
    Dim TempRadius As Integer
    Dim i As Integer
    Dim Red As Long
    Dim Green As Long
    Dim Blue As Long
    
    Dim Done() As Boolean
    ReDim Done(-Radius To Radius, -Radius To Radius)
    
    For i = 1 To NumberOfSteps
        TempRadius = Radius / NumberOfSteps * i
        
        For cx = -TempRadius To TempRadius
            For cy = -TempRadius To TempRadius
                If Not Done(cx, cy) Then
                    If (cx * cx) + (cy * cy) <= TempRadius * TempRadius Then
                        TempColor = GetPixel(Target, cx + x, cy + Y)
                        
                        SetPixelV Target, cx + x, cy + Y, GetLightDarkColor(TempColor, NumberOfSteps * (101 - SelMethod(1).Current), NumberOfSteps * (100 - SelMethod(1).Current) + CLng(i), Darken)
                        
                        Done(cx, cy) = True
                    End If
                End If
            Next cy
        Next cx
    Next i
End Sub


Public Function GetLightDarkColor(Color As Long, Max As Long, Position As Long, Darken As Boolean) As Long
    Dim C1(2) As Byte
    Dim C2(2) As Byte

    Dim i As Integer
    
    Dim RS As Double, GS As Double, BS As Double
    Dim r As Double, g As Double, b As Double
    
    Dim Red1 As Double, Blue1 As Double, Green1 As Double
    Dim Red2 As Double, Blue2 As Double, Green2 As Double

    If Max <= 0 Then
        GetLightDarkColor = Color
        Exit Function
    End If

    If Darken = False Then
        b = 255
        g = 255
        r = 255
    Else
        b = 0
        g = 0
        r = 0
    End If
    
    Red1 = r
    Green1 = g
    Blue1 = b
    
    Blue2 = Color \ 65536
    Green2 = (Color - Blue2 * 65536) \ 256
    Red2 = Color - Blue2 * 65536 - Green2 * 256
    
    If Red1 <> Red2 Then
        RS = ((Red1 - Red2) / Max)
    Else
        RS = 0
    End If
    
    If Green1 <> Green2 Then
        GS = ((Green1 - Green2) / Max)
    Else
        GS = 0
    End If
    
    If Blue1 <> Blue2 Then
        BS = ((Blue1 - Blue2) / Max)
    Else
        BS = 0
    End If
    

    r = r - RS * (Position)
    g = g - GS * (Position)
    b = b - BS * (Position)
    
    If r < 0 Then r = 0
    If r > 255 Then r = 255
    If g < 0 Then g = 0
    If g > 255 Then g = 255
    If b < 0 Then b = 0
    If b > 255 Then b = 255
    
    GetLightDarkColor = RGB(CInt(r), CInt(g), CInt(b))

End Function


Public Sub DrawBlur(Target As Long, x As Single, Y As Single, Radius As Long, NumberOfSteps As Long)
    Dim cx As Long 'X counter
    Dim cy As Long 'Y counter
    Dim TempColor(8) As Long

    
    Dim TempRadius As Integer
    Dim i As Integer
    Dim u As Integer
    Dim Red(3) As Long
    Dim Green(3) As Long
    Dim Blue(3) As Long
    
    Dim Color As Long
    
    Dim Done() As Boolean
    ReDim Done(-Radius To Radius, -Radius To Radius)
    
    For i = 1 To NumberOfSteps
        TempRadius = Radius / NumberOfSteps * i
        
        For cx = -TempRadius To TempRadius 'Step 2
            For cy = -TempRadius To TempRadius 'Step 2
                If Not Done(cx, cy) Then
                    If (cx * cx) + (cy * cy) <= TempRadius * TempRadius Then
                        TempColor(0) = GetPixel(Target, cx + x, cy + Y)
                        TempColor(1) = GetPixel(Target, cx + x, cy + Y - 1)
                        TempColor(2) = GetPixel(Target, cx + x - 1, cy + Y - 1)
                        TempColor(3) = GetPixel(Target, cx + x - 1, cy + Y)
                        
                        For u = 0 To 3
                            Blue(u) = TempColor(u) \ 65536
                            Green(u) = (TempColor(u) - Blue(u) * 65536) \ 256
                            Red(u) = TempColor(u) - Blue(u) * 65536 - Green(u) * 256
                            If Red(u) < 0 Then Red(u) = 0
                            If Red(u) > 255 Then Red(u) = 255
                            If Green(u) < 0 Then Green(u) = 0
                            If Green(u) > 255 Then Green(u) = 255
                            If Blue(u) < 0 Then Blue(u) = 0
                            If Blue(u) > 255 Then Blue(u) = 255
                        Next u
                        
                        Color = RGB((Red(0) + Red(1) + Red(2) + Red(3)) / 4, (Green(0) + Green(1) + Green(2) + Green(3)) / 4, (Blue(0) + Blue(1) + Blue(2) + Blue(3)) / 4)
                        
                        SetPixelV Target, cx + x, cy + Y, Color
                        SetPixelV Target, cx + x, cy + Y - 1, Color
                        SetPixelV Target, cx + x - 1, cy + Y - 1, Color
                        SetPixelV Target, cx + x - 1, cy + Y, Color
                        
                        Done(cx, cy) = True
                    End If
                End If
            Next cy
        Next cx

    Next i
    
    PrevX = x
    PrevY = Y
End Sub

Public Function CalcGreyScale(ByVal Colr As Long) As Integer
    Dim r As Long, g As Long, b As Long
    
    r = Colr Mod 256
    Colr = Colr \ 256
    g = Colr Mod 256
    Colr = Colr \ 256
    b = Colr Mod 256
    
    CalcGreyScale = 76 * r / 255 + 150 * g / 255 + 28 * b / 255

End Function
