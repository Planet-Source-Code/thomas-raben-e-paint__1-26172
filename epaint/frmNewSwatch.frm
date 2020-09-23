VERSION 5.00
Begin VB.Form frmNewSwatch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Swatch..."
   ClientHeight    =   1125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   3825
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   -420
      TabIndex        =   3
      Top             =   600
      Width           =   4695
      Begin VB.CommandButton Command1 
         Caption         =   "Cancel"
         Height          =   315
         Index           =   1
         Left            =   2460
         TabIndex        =   5
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   315
         Index           =   0
         Left            =   3360
         TabIndex        =   4
         Top             =   180
         Width           =   855
      End
   End
   Begin VB.TextBox txtSwatch 
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   300
      Width           =   3075
   End
   Begin VB.PictureBox SwatchColor 
      BackColor       =   &H00000000&
      Height          =   555
      Index           =   0
      Left            =   60
      ScaleHeight     =   495
      ScaleWidth      =   525
      TabIndex        =   0
      Top             =   60
      Width           =   585
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the name for the new swatch:"
      Height          =   195
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   60
      Width           =   3075
   End
End
Attribute VB_Name = "frmNewSwatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
    AddAndExit Index
    
End Sub

Private Sub AddAndExit(Index As Integer)
    Dim x As Integer
    Dim Y As Integer
    
    If Index = 0 Then
        If Me.txtSwatch.Text = "" Then
            MsgBox "You have to enter a name for the swatch.", vbExclamation, "Name required"
        Else
            'add
            x = frmMain.Swatch(frmMain.Swatch.UBound).Left + frmMain.Swatch(frmMain.Swatch.UBound).Width - 1
            Y = frmMain.Swatch(frmMain.Swatch.UBound).Top
            
            If x + frmMain.Swatch(frmMain.Swatch.UBound).Width - 1 > frmMain.SwatchScroll.Width Then
                x = 0
                Y = Y + frmMain.Swatch(frmMain.Swatch.UBound).Height - 1
            End If
            
            If frmMain.Swatch.UBound < 0 Then
                frmMain.Swatch(0).BackColor = frmMain.SelColor(0).BackColor
                frmMain.Swatch(0).ToolTipText = Me.txtSwatch.Text
                frmMain.Swatch(0).Visible = True
            Else
                Load frmMain.Swatch(frmMain.Swatch.UBound + 1)
                frmMain.Swatch(frmMain.Swatch.UBound).BackColor = frmMain.SelColor(0).BackColor
                frmMain.Swatch(frmMain.Swatch.UBound).ToolTipText = Me.txtSwatch.Text
                frmMain.Swatch(frmMain.Swatch.UBound).Move x, Y
                frmMain.Swatch(frmMain.Swatch.UBound).ZOrder 0
                frmMain.Swatch(frmMain.Swatch.UBound).Visible = True
                frmMain.SwatchScroll.Height = (Y + frmMain.Swatch(0).Height * 3)
                
                If frmMain.SwatchScroll.Height > frmMain.SwatchesBg.ScaleHeight Then
                    frmMain.ScrollSwatch.Min = 0
                    frmMain.ScrollSwatch.Max = frmMain.SwatchScroll.Height - frmMain.SwatchesBg.ScaleHeight
                    frmMain.ScrollSwatch.SmallChange = frmMain.SwatchScroll.Height / (frmMain.Swatch(0).Height - 1)
                    frmMain.ScrollSwatch.LargeChange = frmMain.SwatchScroll.Height - (((frmMain.SwatchScroll.Height - frmMain.SwatchScroll.ScaleHeight) / frmMain.SwatchScroll.Height) * frmMain.SwatchScroll.Height)
                    frmMain.ScrollSwatch.Enabled = True
                Else
                    frmMain.ScrollSwatch.Enabled = False
                End If
                
                'save the file...
                SaveSwatches "default.swt"
            End If
            
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.SwatchColor(0).BackColor = frmMain.SelColor(0).BackColor
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Enabled = True
    
End Sub



Private Sub txtSwatch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        AddAndExit 0
    End If
End Sub
