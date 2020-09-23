VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "E-Paint"
   ClientHeight    =   9180
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10980
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox ProcessBg 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   10980
      TabIndex        =   56
      Top             =   8865
      Visible         =   0   'False
      Width           =   10980
      Begin VB.PictureBox Process 
         Height          =   255
         Left            =   960
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   657
         TabIndex        =   58
         Top             =   60
         Width           =   9915
         Begin VB.PictureBox ProcessBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   0
            ScaleHeight     =   315
            ScaleWidth      =   1215
            TabIndex        =   59
            Top             =   -60
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Processing:"
         Height          =   195
         Left            =   60
         TabIndex        =   57
         Top             =   60
         Width           =   825
      End
   End
   Begin VB.PictureBox RightBar 
      Align           =   4  'Align Right
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   8370
      Left            =   8775
      ScaleHeight     =   558
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   147
      TabIndex        =   30
      Top             =   495
      Width           =   2205
      Begin VB.FileListBox lstFilters 
         Height          =   870
         Left            =   420
         Pattern         =   "*.exe"
         TabIndex        =   60
         Top             =   4560
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox SplitH 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   5
         Left            =   -420
         ScaleHeight     =   4
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   225
         TabIndex        =   55
         Top             =   3960
         Width           =   3375
      End
      Begin VB.PictureBox SplitH 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   4
         Left            =   -180
         ScaleHeight     =   4
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   225
         TabIndex        =   40
         Top             =   1860
         Width           =   3375
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   1755
         Left            =   60
         ScaleHeight     =   117
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   141
         TabIndex        =   34
         Top             =   60
         Width           =   2115
         Begin VB.PictureBox SwatchBg 
            Height          =   1455
            Left            =   0
            ScaleHeight     =   93
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   121
            TabIndex        =   36
            ToolTipText     =   "Add New Swatch"
            Top             =   240
            Width           =   1875
            Begin VB.PictureBox SwatchScroll 
               BorderStyle     =   0  'None
               Height          =   315
               Left            =   0
               ScaleHeight     =   21
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   121
               TabIndex        =   37
               ToolTipText     =   "Add New Swatch"
               Top             =   0
               Width           =   1815
               Begin VB.PictureBox Swatch 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Index           =   0
                  Left            =   0
                  MouseIcon       =   "frmMain.frx":0000
                  MousePointer    =   99  'Custom
                  ScaleHeight     =   11
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   11
                  TabIndex        =   38
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   195
               End
            End
         End
         Begin VB.VScrollBar ScrollSwatch 
            Enabled         =   0   'False
            Height          =   1455
            Left            =   1860
            TabIndex        =   35
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Swatches:"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   39
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.PictureBox SwatchesBg 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1995
         Left            =   60
         ScaleHeight     =   133
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   145
         TabIndex        =   32
         Top             =   1920
         Width           =   2175
         Begin VB.PictureBox ColorBlend 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   540
            ScaleHeight     =   285
            ScaleWidth      =   585
            TabIndex        =   53
            Top             =   240
            Width           =   615
         End
         Begin VB.PictureBox ColorBlend 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   1020
            ScaleHeight     =   285
            ScaleWidth      =   585
            TabIndex        =   54
            Top             =   420
            Width           =   615
         End
         Begin VB.PictureBox ColorScroll 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   75
            Index           =   2
            Left            =   180
            Picture         =   "frmMain.frx":08CA
            ScaleHeight     =   5
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   9
            TabIndex        =   52
            Top             =   1860
            Width           =   135
         End
         Begin VB.PictureBox ColorScroll 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   75
            Index           =   1
            Left            =   180
            Picture         =   "frmMain.frx":0C10
            ScaleHeight     =   5
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   9
            TabIndex        =   51
            Top             =   1500
            Width           =   135
         End
         Begin VB.PictureBox ColorScroll 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   75
            Index           =   0
            Left            =   180
            Picture         =   "frmMain.frx":0F56
            ScaleHeight     =   5
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   9
            TabIndex        =   50
            Top             =   1140
            Width           =   135
         End
         Begin VB.PictureBox ColorBar 
            AutoRedraw      =   -1  'True
            Height          =   135
            Index           =   2
            Left            =   180
            ScaleHeight     =   5
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   97
            TabIndex        =   47
            Top             =   1680
            Width           =   1515
         End
         Begin VB.PictureBox ColorBar 
            AutoRedraw      =   -1  'True
            Height          =   135
            Index           =   1
            Left            =   180
            ScaleHeight     =   5
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   97
            TabIndex        =   44
            Top             =   1320
            Width           =   1515
         End
         Begin VB.PictureBox ColorBar 
            AutoRedraw      =   -1  'True
            Height          =   135
            Index           =   0
            Left            =   180
            ScaleHeight     =   5
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   97
            TabIndex        =   41
            Top             =   960
            Width           =   1515
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "B"
            Height          =   195
            Index           =   2
            Left            =   0
            TabIndex        =   49
            Top             =   1650
            Width           =   105
         End
         Begin VB.Label lblColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            Height          =   255
            Index           =   2
            Left            =   1800
            TabIndex        =   48
            Top             =   1620
            Width           =   330
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "G"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   46
            Top             =   1290
            Width           =   120
         End
         Begin VB.Label lblColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   45
            Top             =   1260
            Width           =   330
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "R"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   43
            Top             =   930
            Width           =   120
         End
         Begin VB.Label lblColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            Height          =   255
            Index           =   0
            Left            =   1800
            TabIndex        =   42
            Top             =   900
            Width           =   330
         End
         Begin VB.Label Label1 
            Caption         =   "Color:"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.PictureBox SplitH 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   3
         Left            =   -540
         ScaleHeight     =   4
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   225
         TabIndex        =   31
         Top             =   0
         Width           =   3375
      End
   End
   Begin VB.Timer CoordsTimer 
      Interval        =   1000
      Left            =   1080
      Top             =   7980
   End
   Begin VB.PictureBox LeftBar 
      Align           =   3  'Align Left
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   8370
      Left            =   0
      ScaleHeight     =   558
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   1
      Top             =   495
      Width           =   975
      Begin VB.PictureBox Cmd 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   13
         Left            =   0
         Picture         =   "frmMain.frx":129C
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   29
         ToolTipText     =   "Blur Tool (W)"
         Top             =   2520
         Width           =   420
      End
      Begin VB.PictureBox Cmd 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   12
         Left            =   480
         Picture         =   "frmMain.frx":162B
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   28
         ToolTipText     =   "Text Tool (T)"
         Top             =   2520
         Width           =   420
      End
      Begin VB.PictureBox Cmd 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   11
         Left            =   0
         Picture         =   "frmMain.frx":19AD
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   27
         ToolTipText     =   "Brighten Tool (L)"
         Top             =   3060
         Width           =   420
      End
      Begin VB.PictureBox Cmd 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   10
         Left            =   0
         Picture         =   "frmMain.frx":1D44
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   26
         ToolTipText     =   "Circle Tool (E)"
         Top             =   2040
         Width           =   420
      End
      Begin VB.PictureBox Cmd 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   9
         Left            =   480
         Picture         =   "frmMain.frx":20D5
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   25
         ToolTipText     =   "Rectangle Tool (B)"
         Top             =   2040
         Width           =   420
      End
      Begin VB.PictureBox SplitH 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   2
         Left            =   0
         ScaleHeight     =   4
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   61
         TabIndex        =   24
         Top             =   3000
         Width           =   915
      End
      Begin VB.PictureBox SplitH 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   1
         Left            =   0
         ScaleHeight     =   4
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   61
         TabIndex        =   23
         Top             =   1020
         Width           =   915
      End
      Begin VB.PictureBox Cmd 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   8
         Left            =   480
         Picture         =   "frmMain.frx":246C
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   22
         ToolTipText     =   "Eyedropper Tool (Y)"
         Top             =   3060
         Width           =   420
      End
      Begin VB.PictureBox Cmd 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   7
         Left            =   0
         Picture         =   "frmMain.frx":27F4
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   21
         ToolTipText     =   "Crop Tool (R)"
         Top             =   540
         Width           =   420
      End
      Begin VB.PictureBox Cmd 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   6
         Left            =   0
         Picture         =   "frmMain.frx":2B84
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   20
         ToolTipText     =   "Airbrush Tool (A)"
         Top             =   1080
         Width           =   420
      End
      Begin VB.PictureBox Cmd 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   5
         Left            =   480
         Picture         =   "frmMain.frx":2F18
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   19
         ToolTipText     =   "Select Area Tool (S)"
         Top             =   60
         Width           =   420
      End
      Begin VB.PictureBox Cmd 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   4
         Left            =   0
         Picture         =   "frmMain.frx":329B
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   18
         ToolTipText     =   "Gradient Tool (G)"
         Top             =   1560
         Width           =   420
      End
      Begin VB.PictureBox SplitH 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   0
         Left            =   -240
         ScaleHeight     =   4
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   93
         TabIndex        =   16
         Top             =   0
         Width           =   1395
      End
      Begin VB.PictureBox ColorsBg 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   0
         ScaleHeight     =   81
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   61
         TabIndex        =   11
         Top             =   3540
         Width           =   915
         Begin VB.PictureBox Gradient 
            AutoRedraw      =   -1  'True
            Height          =   255
            Left            =   0
            MouseIcon       =   "frmMain.frx":3644
            MousePointer    =   99  'Custom
            ScaleHeight     =   13
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   57
            TabIndex        =   17
            ToolTipText     =   "Gradient Scale"
            Top             =   840
            Width           =   915
         End
         Begin VB.PictureBox RestoreColors 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   150
            Left            =   180
            Picture         =   "frmMain.frx":3F0E
            ScaleHeight     =   150
            ScaleWidth      =   150
            TabIndex        =   15
            ToolTipText     =   "Restore colors to black and white"
            Top             =   540
            Width           =   150
         End
         Begin VB.PictureBox ShiftColors 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   150
            Left            =   600
            Picture         =   "frmMain.frx":4263
            ScaleHeight     =   150
            ScaleWidth      =   150
            TabIndex        =   14
            ToolTipText     =   "Flip foreground and background color"
            Top             =   180
            Width           =   150
         End
         Begin VB.PictureBox SelColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   120
            ScaleHeight     =   345
            ScaleWidth      =   405
            TabIndex        =   12
            ToolTipText     =   "Select foreground color"
            Top             =   120
            Width           =   435
         End
         Begin VB.PictureBox SelColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   360
            ScaleHeight     =   345
            ScaleWidth      =   405
            TabIndex        =   13
            ToolTipText     =   "Select background color"
            Top             =   360
            Width           =   435
         End
      End
      Begin VB.PictureBox Cmd 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   3
         Left            =   480
         Picture         =   "frmMain.frx":45B6
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   5
         ToolTipText     =   "Floodfill Tool (F)"
         Top             =   1560
         Width           =   420
      End
      Begin VB.PictureBox Cmd 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   2
         Left            =   480
         Picture         =   "frmMain.frx":4957
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   4
         ToolTipText     =   "Magnifying Tool (Z)"
         Top             =   540
         Width           =   420
      End
      Begin VB.Timer CmdTimer 
         Interval        =   50
         Left            =   0
         Top             =   7500
      End
      Begin VB.PictureBox Cmd 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   1
         Left            =   480
         Picture         =   "frmMain.frx":4CE5
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   3
         ToolTipText     =   "Pen Tool (P)"
         Top             =   1080
         Width           =   420
      End
      Begin VB.PictureBox Cmd 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   0
         Left            =   0
         Picture         =   "frmMain.frx":506D
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   2
         ToolTipText     =   "Cursor (C)"
         Top             =   60
         Width           =   420
      End
   End
   Begin VB.PictureBox TopBar 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   732
      TabIndex        =   0
      Top             =   0
      Width           =   10980
      Begin VB.PictureBox CoordsInfo 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   8460
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   205
         TabIndex        =   6
         Top             =   30
         Width           =   3075
         Begin VB.PictureBox Picture2 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   1
            Left            =   1740
            Picture         =   "frmMain.frx":53EB
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   8
            Top             =   120
            Width           =   225
         End
         Begin VB.PictureBox Picture2 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   0
            Left            =   180
            Picture         =   "frmMain.frx":575B
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   7
            Top             =   120
            Width           =   225
         End
         Begin VB.Label lblArea 
            Caption         =   "0 x 0"
            Height          =   195
            Left            =   2040
            TabIndex        =   10
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label lblCoords 
            Caption         =   "0, 0"
            Height          =   195
            Left            =   480
            TabIndex        =   9
            Top             =   120
            Width           =   1095
         End
      End
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu NewImage 
         Caption         =   "&New..."
         Shortcut        =   ^N
      End
      Begin VB.Menu Open 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu split006 
         Caption         =   "-"
      End
      Begin VB.Menu Save 
         Caption         =   "&Save..."
         Shortcut        =   ^S
      End
      Begin VB.Menu SaveAs 
         Caption         =   "Save &As..."
         Shortcut        =   ^D
      End
      Begin VB.Menu split005 
         Caption         =   "-"
      End
      Begin VB.Menu Close 
         Caption         =   "&Close"
         Shortcut        =   ^W
      End
      Begin VB.Menu split004 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "&Edit"
      Begin VB.Menu Undo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu split001 
         Caption         =   "-"
      End
      Begin VB.Menu Cut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu Copy 
         Caption         =   "C&opy"
         Shortcut        =   ^C
      End
      Begin VB.Menu Paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu split002 
         Caption         =   "-"
      End
      Begin VB.Menu PasteSpecial 
         Caption         =   "Pa&ste Special"
         Shortcut        =   ^B
      End
      Begin VB.Menu split003 
         Caption         =   "-"
      End
      Begin VB.Menu SelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu ClearSelection 
         Caption         =   "C&lear Selection"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu Image 
      Caption         =   "&Image"
      Begin VB.Menu FlipHorizontal 
         Caption         =   "Flip &Horizontal"
      End
      Begin VB.Menu FlipVertical 
         Caption         =   "Flip &Vertical"
      End
      Begin VB.Menu split009 
         Caption         =   "-"
      End
      Begin VB.Menu InvertImage 
         Caption         =   "&Invert Image"
         Shortcut        =   ^I
      End
      Begin VB.Menu GreyScalemenu 
         Caption         =   "Make &Greyscale"
         Begin VB.Menu Greyscale 
            Caption         =   "&Black && White"
            Index           =   0
         End
         Begin VB.Menu Greyscale 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu Greyscale 
            Caption         =   "&Red && White"
            Index           =   2
         End
         Begin VB.Menu Greyscale 
            Caption         =   "R&ed && Black"
            Index           =   3
         End
         Begin VB.Menu Greyscale 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu Greyscale 
            Caption         =   "&Green && White"
            Index           =   5
         End
         Begin VB.Menu Greyscale 
            Caption         =   "Gr&een && Black"
            Index           =   6
         End
         Begin VB.Menu Greyscale 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu Greyscale 
            Caption         =   "B&lue && White"
            Index           =   8
         End
         Begin VB.Menu Greyscale 
            Caption         =   "Bl&ue && Black"
            Index           =   9
         End
      End
      Begin VB.Menu split010 
         Caption         =   "-"
      End
      Begin VB.Menu ImageSize 
         Caption         =   "Image &Size"
      End
   End
   Begin VB.Menu Filters 
      Caption         =   "&Filters"
      Begin VB.Menu Filter 
         Caption         =   "Default"
         Index           =   0
      End
   End
   Begin VB.Menu Window 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu ArrHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu ArrVertival 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu Cascade 
         Caption         =   "&Cascade"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu AboutEPaint 
         Caption         =   "&About EPaint..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MouseDown As Boolean
Dim Xpos As Integer
Dim SelectedColor As Integer



Private Sub AboutEPaint_Click()
    Me.Enabled = False
    frmAbout.Show
    
End Sub

Private Sub ArrHorizontal_Click()
    frmMain.Arrange 1
    
    'ExecFilter "filter.exe"
    
End Sub

Private Sub ArrVertival_Click()
    frmMain.Arrange 2
End Sub

Private Sub Cascade_Click()
    frmMain.Arrange 0
    ExecFilter "filter.exe"
End Sub

Private Sub ClearSelection_Click()
    'prepare for undo...
    frmMain.ActiveForm.Undo.Width = frmMain.ActiveForm.Buffer.Width
    frmMain.ActiveForm.Undo.Height = frmMain.ActiveForm.Buffer.Height
    BitBlt frmMain.ActiveForm.Undo.hdc, 0, 0, frmMain.ActiveForm.Buffer.ScaleWidth, frmMain.ActiveForm.Buffer.ScaleHeight, frmMain.ActiveForm.Buffer.hdc, 0, 0, vbSrcCopy

    frmMain.ActiveForm.BufferSelected.BackColor = frmMain.SelColor(1).BackColor
    UpdateArea frmMain.ActiveForm.Buffer, Abs(frmMain.ActiveForm.PaintArea.Left) / frmMain.ActiveForm.GetZoomFactor * 100, Abs(frmMain.ActiveForm.PaintArea.Top) / frmMain.ActiveForm.GetZoomFactor * 100, frmMain.ActiveForm.GetZoomFactor
    frmMain.ActiveForm.SelectArea.Visible = False
    
End Sub

Private Sub Close_Click()
    Unload frmMain.ActiveForm
    
End Sub

Private Sub Cmd_Click(Index As Integer)
    SelectTool Index
End Sub

Private Sub CmdTimer_Timer()
    CheckFlatButtons
    
    frmMain.TopBar.Line (0, 0)-(frmMain.TopBar.ScaleWidth, 0), vb3DShadow
    frmMain.TopBar.Line (0, 1)-(frmMain.TopBar.ScaleWidth, 1), vb3DHighlight
    
End Sub




Private Sub ColorBlend_Click(Index As Integer)
    SelectedColor = Index
    SetColorBars
    
End Sub

Private Sub ColorScroll_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Xpos = X
    MouseDown = True
    
End Sub

Private Sub ColorScroll_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim TmpX As Integer
    If MouseDown = True Then
        TmpX = Me.ColorScroll(Index).Left + X - Xpos
        
        If TmpX < Me.ColorBar(Index).Left - Me.ColorScroll(Index).Width / 2 Then TmpX = Me.ColorBar(Index).Left - Me.ColorScroll(Index).Width / 2
        If TmpX > Me.ColorBar(Index).Left + Me.ColorBar(Index).Width - Me.ColorScroll(Index).Width / 2 Then TmpX = Me.ColorBar(Index).Left + Me.ColorBar(Index).Width - Me.ColorScroll(Index).Width / 2
        
        Me.ColorScroll(Index).Left = TmpX
        SetColorBars
        
        On Error Resume Next
        frmMain.ActiveForm.Buffer.ForeColor = frmMain.SelColor(0).BackColor
        frmMain.ActiveForm.TextInput.ForeColor = frmMain.SelColor(0).BackColor
        frmMain.ActiveForm.TextInput.BackColor = frmMain.SelColor(1).BackColor
    End If
End Sub

Private Sub ColorScroll_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseDown = False
    DrawGradient Me.Gradient, Me.SelColor(0).BackColor, Me.SelColor(1).BackColor
    
End Sub

Public Sub SetColorBars()
    Dim i As Integer
    Dim Value As Integer
    Dim r As Long, g As Long, b As Long
    
    If MouseDown = False Then
        b = frmMain.SelColor(SelectedColor).BackColor \ 65536
        g = (frmMain.SelColor(SelectedColor).BackColor - b * 65536) \ 256
        r = frmMain.SelColor(SelectedColor).BackColor - b * 65536 - g * 256
        
        If r > 255 Then r = 255
        If r < 0 Then r = 0
        If g > 255 Then g = 255
        If g < 0 Then g = 0
        If b > 255 Then b = 255
        If b < 0 Then b = 0
        
        Me.lblColor(0).Caption = r
        Me.lblColor(1).Caption = g
        Me.lblColor(2).Caption = b
        
        Me.ColorScroll(0).Left = (r / 255 * Me.ColorBar(0).Width) + Me.ColorBar(0).Left - Me.ColorScroll(0).Width / 2
        Me.ColorScroll(1).Left = (g / 255 * Me.ColorBar(1).Width) + Me.ColorBar(1).Left - Me.ColorScroll(1).Width / 2
        Me.ColorScroll(2).Left = (b / 255 * Me.ColorBar(2).Width) + Me.ColorBar(2).Left - Me.ColorScroll(2).Width / 2
    Else
    
        For i = 0 To 2
            If (Me.ColorScroll(i).Left - (Me.ColorBar(i).Left - CInt(Me.ColorScroll(i).Width / 2))) <= 0 Then
                Value = 0
            Else
                Value = ((Me.ColorScroll(i).Left - (Me.ColorBar(i).Left - (Me.ColorScroll(i).Width / 2))) / Me.ColorBar(i).Width * 256)
            End If
    
            Me.lblColor(i).Caption = Value
        Next i
        
        
    End If
    
    Me.ColorBlend(SelectedColor).BackColor = RGB(Me.lblColor(0).Caption, Me.lblColor(1).Caption, Me.lblColor(2).Caption)
    Me.SelColor(SelectedColor).BackColor = Me.ColorBlend(SelectedColor).BackColor
    
    
    DrawGradient Me.ColorBar(0), RGB(0, Me.lblColor(1).Caption, Me.lblColor(2).Caption), RGB(255, Me.lblColor(1).Caption, Me.lblColor(2).Caption)
    DrawGradient Me.ColorBar(1), RGB(Me.lblColor(0).Caption, 0, Me.lblColor(2).Caption), RGB(Me.lblColor(0).Caption, 255, Me.lblColor(2).Caption)
    DrawGradient Me.ColorBar(2), RGB(Me.lblColor(0).Caption, Me.lblColor(1).Caption, 0), RGB(Me.lblColor(0).Caption, Me.lblColor(1).Caption, 255)
    
End Sub

Public Sub DrawGradients()
    
End Sub

Private Sub Copy_Click()
   
    Clipboard.Clear
    Clipboard.SetData frmMain.ActiveForm.BufferSelected.Image

End Sub

Private Sub Cut_Click()
    Clipboard.Clear
    Clipboard.SetData frmMain.ActiveForm.BufferSelected.Image
    frmMain.ActiveForm.BufferSelected.BackColor = frmMain.SelColor(1).BackColor
    UpdateArea frmMain.ActiveForm.Buffer, Abs(frmMain.ActiveForm.PaintArea.Left) / frmMain.ActiveForm.GetZoomFactor * 100, Abs(frmMain.ActiveForm.PaintArea.Top) / frmMain.ActiveForm.GetZoomFactor * 100, frmMain.ActiveForm.GetZoomFactor
    frmMain.ActiveForm.SelectArea.Visible = False
    
End Sub

Private Sub Filters_Click()
    Dim i As Integer
    Dim X As String
    
    On Error GoTo NoWindow
    X = frmMain.ActiveForm.Caption
    
    For i = 0 To frmMain.Filter.UBound
        frmMain.Filter(i).Enabled = True
    Next i
    
    Exit Sub
NoWindow:
    For i = 0 To frmMain.Filter.UBound
        frmMain.Filter(i).Enabled = False
    Next i
    
End Sub

Private Sub Filter_Click(Index As Integer)
    ExecFilter Filter(Index).Tag
    
End Sub

Private Sub FlipVertical_Click()
    Dim i As Integer
    Dim mDC As Long, mBitmap As Long
    
    mDC = CreateCompatibleDC(GetDC(0))

    If frmMain.ActiveForm.SelectArea.Visible = True Then
        mBitmap = CreateCompatibleBitmap(GetDC(0), frmMain.ActiveForm.BufferSelected.ScaleWidth, frmMain.ActiveForm.BufferSelected.ScaleHeight)
        SelectObject mDC, mBitmap
        
        For i = 0 To frmMain.ActiveForm.BufferSelected.ScaleHeight - 1
            BitBlt mDC, 0, frmMain.ActiveForm.BufferSelected.ScaleHeight - 1 - i, frmMain.ActiveForm.BufferSelected.ScaleWidth, 1, frmMain.ActiveForm.BufferSelected.hdc, 0, i, vbSrcCopy
        Next i
        
        BitBlt frmMain.ActiveForm.BufferSelected.hdc, 0, 0, frmMain.ActiveForm.BufferSelected.ScaleWidth, frmMain.ActiveForm.BufferSelected.ScaleHeight, mDC, 0, 0, vbSrcCopy
        UpdateArea frmMain.ActiveForm.Buffer, 0, 0, frmMain.ActiveForm.GetZoomFactor
        
    Else
        mBitmap = CreateCompatibleBitmap(GetDC(0), frmMain.ActiveForm.Buffer.ScaleWidth, frmMain.ActiveForm.Buffer.ScaleHeight)
        SelectObject mDC, mBitmap
        
        For i = 0 To frmMain.ActiveForm.Buffer.ScaleHeight - 1
            BitBlt mDC, 0, frmMain.ActiveForm.Buffer.ScaleHeight - 1 - i, frmMain.ActiveForm.Buffer.ScaleWidth, 1, frmMain.ActiveForm.Buffer.hdc, 0, i, vbSrcCopy
        Next i
        
        BitBlt frmMain.ActiveForm.Buffer.hdc, 0, 0, frmMain.ActiveForm.Buffer.ScaleWidth, frmMain.ActiveForm.Buffer.ScaleHeight, mDC, 0, 0, vbSrcCopy
        UpdateArea frmMain.ActiveForm.Buffer, 0, 0, frmMain.ActiveForm.GetZoomFactor
        
    End If
    
    DeleteDC mDC
    DeleteObject mBitmap
End Sub

Private Sub GreyScale_Click(Index As Integer)
    Dim X As Integer, Y As Integer, C As Long
    
    Me.ProcessBg.Visible = True
    Me.ProcessBar.Width = 1
    Me.ProcessBar.Visible = True
    Me.Enabled = False
    
    If frmMain.ActiveForm.SelectArea.Visible = True Then
        frmMain.ActiveForm.Undo.Width = frmMain.ActiveForm.BufferSelected.Width
        frmMain.ActiveForm.Undo.Height = frmMain.ActiveForm.BufferSelected.Height
        BitBlt frmMain.ActiveForm.Undo.hdc, 0, 0, frmMain.ActiveForm.BufferSelected.ScaleWidth, frmMain.ActiveForm.BufferSelected.ScaleHeight, frmMain.ActiveForm.BufferSelected.hdc, 0, 0, vbSrcCopy
        
        For X = 0 To frmMain.ActiveForm.BufferSelected.ScaleWidth
            For Y = 0 To frmMain.ActiveForm.BufferSelected.ScaleHeight
                C = CalcGreyScale(GetPixel(frmMain.ActiveForm.BufferSelected.hdc, X, Y))
                Select Case Index
                    Case 0
                        C = RGB(C, C, C)
                    Case 2
                        C = RGB(255, C, C)
                    Case 3
                        C = RGB(C, 0, 0)
                    Case 5
                        C = RGB(C, 255, C)
                    Case 6
                        C = RGB(0, C, 0)
                    Case 8
                        C = RGB(C, C, 255)
                    Case 9
                        C = RGB(0, 0, C)
                End Select
                SetPixelV frmMain.ActiveForm.BufferSelected.hdc, X, Y, C
            Next Y
            
            Me.ProcessBar.Width = (X / frmMain.ActiveForm.BufferSelected.ScaleWidth) * Me.Process.ScaleWidth
            DoEvents
            
        Next X
        frmMain.ActiveForm.Buffer.Refresh
        UpdateArea frmMain.ActiveForm.Buffer, 0, 0, frmMain.ActiveForm.GetZoomFactor
    
    Else
        frmMain.ActiveForm.Undo.Width = frmMain.ActiveForm.Buffer.Width
        frmMain.ActiveForm.Undo.Height = frmMain.ActiveForm.Buffer.Height
        BitBlt frmMain.ActiveForm.Undo.hdc, 0, 0, frmMain.ActiveForm.Buffer.ScaleWidth, frmMain.ActiveForm.Buffer.ScaleHeight, frmMain.ActiveForm.Buffer.hdc, 0, 0, vbSrcCopy
        
        For X = 0 To frmMain.ActiveForm.Buffer.ScaleWidth
            For Y = 0 To frmMain.ActiveForm.Buffer.ScaleHeight
                C = CalcGreyScale(GetPixel(frmMain.ActiveForm.Buffer.hdc, X, Y))
                Select Case Index
                    Case 0
                        C = RGB(C, C, C)
                    Case 2
                        C = RGB(255, C, C)
                    Case 3
                        C = RGB(C, 0, 0)
                    Case 5
                        C = RGB(C, 255, C)
                    Case 6
                        C = RGB(0, C, 0)
                    Case 8
                        C = RGB(C, C, 255)
                    Case 9
                        C = RGB(0, 0, C)
                End Select
                SetPixelV frmMain.ActiveForm.Buffer.hdc, X, Y, C
            Next Y
            
            Me.ProcessBar.Width = (X / frmMain.ActiveForm.Buffer.ScaleWidth) * Me.Process.ScaleWidth
            DoEvents
            
        Next X
        frmMain.ActiveForm.Buffer.Refresh
        UpdateArea frmMain.ActiveForm.Buffer, 0, 0, frmMain.ActiveForm.GetZoomFactor
    End If
    
    
    Me.Enabled = True
    Me.ProcessBg.Visible = False
    
    
End Sub

Private Sub Image_Click()
    Dim X As String
    
    On Error GoTo NoWindow
    X = frmMain.ActiveForm.Caption
    
    frmMain.FlipHorizontal.Enabled = True
    frmMain.FlipVertical.Enabled = True
    frmMain.InvertImage.Enabled = True
    frmMain.ImageSize.Enabled = True
    frmMain.GreyScalemenu.Enabled = True
    Exit Sub
    
NoWindow:
    frmMain.FlipHorizontal.Enabled = False
    frmMain.FlipVertical.Enabled = False
    frmMain.InvertImage.Enabled = False
    frmMain.ImageSize.Enabled = False
    frmMain.GreyScalemenu.Enabled = False
    
    
End Sub

Private Sub FlipHorizontal_Click()
    Dim i As Integer
    Dim mDC As Long, mBitmap As Long
    
    mDC = CreateCompatibleDC(GetDC(0))

    If frmMain.ActiveForm.SelectArea.Visible = True Then
        mBitmap = CreateCompatibleBitmap(GetDC(0), frmMain.ActiveForm.BufferSelected.ScaleWidth, frmMain.ActiveForm.BufferSelected.ScaleHeight)
        SelectObject mDC, mBitmap
        
        For i = 0 To frmMain.ActiveForm.BufferSelected.ScaleWidth - 1
            BitBlt mDC, frmMain.ActiveForm.BufferSelected.ScaleWidth - 1 - i, 0, 1, frmMain.ActiveForm.BufferSelected.ScaleHeight, frmMain.ActiveForm.BufferSelected.hdc, i, 0, vbSrcCopy
        Next i
        
        BitBlt frmMain.ActiveForm.BufferSelected.hdc, 0, 0, frmMain.ActiveForm.BufferSelected.ScaleWidth, frmMain.ActiveForm.BufferSelected.ScaleHeight, mDC, 0, 0, vbSrcCopy
        UpdateArea frmMain.ActiveForm.Buffer, 0, 0, frmMain.ActiveForm.GetZoomFactor
        
    Else
        mBitmap = CreateCompatibleBitmap(GetDC(0), frmMain.ActiveForm.Buffer.ScaleWidth, frmMain.ActiveForm.Buffer.ScaleHeight)
        SelectObject mDC, mBitmap
        
        For i = 0 To frmMain.ActiveForm.Buffer.ScaleWidth - 1
            BitBlt mDC, frmMain.ActiveForm.Buffer.ScaleWidth - 1 - i, 0, 1, frmMain.ActiveForm.Buffer.ScaleHeight, frmMain.ActiveForm.Buffer.hdc, i, 0, vbSrcCopy
        Next i
        
        BitBlt frmMain.ActiveForm.Buffer.hdc, 0, 0, frmMain.ActiveForm.Buffer.ScaleWidth, frmMain.ActiveForm.Buffer.ScaleHeight, mDC, 0, 0, vbSrcCopy
        UpdateArea frmMain.ActiveForm.Buffer, 0, 0, frmMain.ActiveForm.GetZoomFactor
        
    End If
    
    DeleteDC mDC
    DeleteObject mBitmap
    
End Sub

Private Sub Gradient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Me.SelColor(0).BackColor = Me.Gradient.Point(X, Y)
    Else
        Me.SelColor(1).BackColor = Me.Gradient.Point(X, Y)
    End If
    
    DrawPreviewGradient
    
End Sub

Private Sub ImageSize_Click()
    Me.Enabled = False
    frmSize.txtWidth = frmMain.ActiveForm.Buffer.ScaleWidth
    frmSize.txtHeight = frmMain.ActiveForm.Buffer.ScaleHeight
    frmSize.Show
    
End Sub

Private Sub InvertImage_Click()
    Dim r As RECT
    
    If frmMain.ActiveForm.SelectArea.Visible = True Then
        frmMain.ActiveForm.Undo.Width = frmMain.ActiveForm.BufferSelected.Width
        frmMain.ActiveForm.Undo.Height = frmMain.ActiveForm.BufferSelected.Height
        BitBlt frmMain.ActiveForm.Undo.hdc, 0, 0, frmMain.ActiveForm.BufferSelected.ScaleWidth, frmMain.ActiveForm.BufferSelected.ScaleHeight, frmMain.ActiveForm.BufferSelected.hdc, 0, 0, vbSrcCopy
    
        r.Left = 0
        r.Top = 0
        r.Bottom = frmMain.ActiveForm.BufferSelected.ScaleHeight
        r.Right = frmMain.ActiveForm.BufferSelected.ScaleWidth
        InvertRect frmMain.ActiveForm.BufferSelected.hdc, r
        UpdateArea frmMain.ActiveForm.Buffer, 0, 0, frmMain.ActiveForm.GetZoomFactor
    Else
        frmMain.ActiveForm.Undo.Width = frmMain.ActiveForm.Buffer.Width
        frmMain.ActiveForm.Undo.Height = frmMain.ActiveForm.Buffer.Height
        BitBlt frmMain.ActiveForm.Undo.hdc, 0, 0, frmMain.ActiveForm.Buffer.ScaleWidth, frmMain.ActiveForm.Buffer.ScaleHeight, frmMain.ActiveForm.Buffer.hdc, 0, 0, vbSrcCopy
    
        r.Left = 0
        r.Top = 0
        r.Bottom = frmMain.ActiveForm.Buffer.ScaleHeight
        r.Right = frmMain.ActiveForm.Buffer.ScaleWidth
        InvertRect frmMain.ActiveForm.Buffer.hdc, r
        UpdateArea frmMain.ActiveForm.Buffer, 0, 0, frmMain.ActiveForm.GetZoomFactor
        
    End If
    
End Sub

Private Sub MDIForm_Load()
    Me.LeftBar.BackColor = vb3DFace
    Me.RightBar.BackColor = vb3DFace
    
    SetColorBars
    Init
    DoEvents
    frmSPlash.lblStatus.Caption = "Loading Swatches"
    DoEvents
    LoadSwatches "Default.swt"
    
    SelectTool 0
    SelectToolRect 0
    SelectToolBrush 0
    SelectToolLight 0
    
    Unload frmSPlash
    Me.SetFocus
    
End Sub

Private Sub MDIForm_Resize()
    Dim BarX As Long
    
    On Error Resume Next
    
    BarX = Me.TopBar.ScaleWidth - Me.CoordsInfo.Width
    
    If BarX < frmControls.DrawToolbar(CurrentButton).Left + frmControls.DrawToolbar(CurrentButton).Width Then
        BarX = frmControls.DrawToolbar(CurrentButton).Left + frmControls.DrawToolbar(CurrentButton).Width
    End If
    
    Me.CoordsInfo.Left = BarX
    
    Me.Process.Width = Me.ProcessBg.Width - Me.Process.Left
    
End Sub

Private Sub Edit_Click()
    Dim X As String
    
    On Error GoTo NoWindow
    X = frmMain.ActiveForm.Caption
    
    If frmMain.ActiveForm.SelectArea.Visible = True Then
        frmMain.Copy.Enabled = True
        frmMain.Cut.Enabled = True
        frmMain.ClearSelection.Enabled = True
    Else
        frmMain.Copy.Enabled = False
        frmMain.Cut.Enabled = False
        frmMain.ClearSelection.Enabled = False
    End If
    
    frmMain.Undo.Enabled = True
    frmMain.SelectAll.Enabled = True
    
    If Clipboard.GetData(vbCFBitmap) = 0 Then
        frmMain.Paste.Enabled = False
        frmMain.PasteSpecial.Enabled = False
    Else
        frmMain.Paste.Enabled = True
        frmMain.PasteSpecial.Enabled = True
    End If
    
    Exit Sub
    
NoWindow:
    frmMain.Cut.Enabled = False
    frmMain.Copy.Enabled = False
    frmMain.Undo.Enabled = False
    frmMain.Paste.Enabled = False
    frmMain.PasteSpecial.Enabled = False
    frmMain.SelectAll.Enabled = False
    frmMain.ClearSelection.Enabled = False
End Sub

Private Sub File_Click()
    Dim X As String
    
    On Error GoTo NoWindow
    X = frmMain.ActiveForm.Caption
    frmMain.Save.Enabled = True
    frmMain.SaveAs.Enabled = True
    frmMain.Close.Enabled = True
    
    Exit Sub
    
NoWindow:
    frmMain.Save.Enabled = False
    frmMain.SaveAs.Enabled = False
    frmMain.Close.Enabled = False
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If Cancel = 0 Then
        Unload frmFonts
        Unload frmFontStyle
        Unload frmSelector
        Unload frmColor
        Unload frmControls
        Unload frmNew
        Unload frmNewSwatch
    End If
    
    
    End
    
End Sub

Private Sub NewImage_Click()
    Me.Enabled = False
    frmNew.Show
    frmNew.txtName.Text = "Untitled"
End Sub

Private Sub Open_Click()
 Dim FileName As String

 FileName = GetOpenName("Open...")
  
 If Asc(Left(FileName, 1)) <> 32 Then
    Dim f As New frmPaint
    
    f.Tag = FileName
    f.Caption = FileName & " - 100%"
    f.Show
    
    On Error Resume Next
    f.PaintArea.AutoSize = True
    f.PaintArea.Picture = LoadPicture(FileName)
    f.Buffer.Picture = f.PaintArea.Picture
    f.PaintArea.AutoSize = False
    f.Buffer.Width = f.PaintArea.Width
    f.Buffer.Height = f.PaintArea.Height
    f.Buffer.Tag = FileName
 End If
 
End Sub

Private Sub Paste_Click()
    Dim X As Integer, Y As Integer
    Dim X2 As Integer, Y2 As Integer
    
    If frmMain.ActiveForm.PaintArea.Left < 0 Then
        X = Abs(frmMain.ActiveForm.PaintArea.Left) / frmMain.ActiveForm.GetZoomFactor * 100
        X2 = frmMain.ActiveForm.PaintArea.Left
    Else
        X = 0
        X2 = 0
    End If
    
    If frmMain.ActiveForm.PaintArea.Top < 0 Then
        Y = Abs(frmMain.ActiveForm.PaintArea.Top) / frmMain.ActiveForm.GetZoomFactor * 100
        Y2 = frmMain.ActiveForm.PaintArea.Top
    Else
        Y = 0
        Y2 = 0
    End If
    
    UpdateArea frmMain.ActiveForm.Buffer, X, Y, frmMain.ActiveForm.GetZoomFactor

    frmControls.ClipBoardTemp.Picture = Clipboard.GetData(vbCFBitmap)
    
    frmMain.ActiveForm.SelectArea.Move X2, Y2, frmControls.ClipBoardTemp.ScaleWidth * (frmMain.ActiveForm.GetZoomFactor / 100), frmControls.ClipBoardTemp.ScaleHeight * (frmMain.ActiveForm.GetZoomFactor / 100)
    
    frmMain.ActiveForm.BufferSelected.Left = frmMain.ActiveForm.SelectArea.Left * (100 / frmMain.ActiveForm.GetZoomFactor)
    frmMain.ActiveForm.BufferSelected.Top = frmMain.ActiveForm.SelectArea.Top * (100 / frmMain.ActiveForm.GetZoomFactor)
    frmMain.ActiveForm.BufferSelected.Width = frmMain.ActiveForm.SelectArea.Width * (100 / frmMain.ActiveForm.GetZoomFactor)
    frmMain.ActiveForm.BufferSelected.Height = frmMain.ActiveForm.SelectArea.Height * (100 / frmMain.ActiveForm.GetZoomFactor)
    BitBlt frmMain.ActiveForm.BufferSelected.hdc, 0, 0, frmMain.ActiveForm.BufferSelected.Width, frmMain.ActiveForm.BufferSelected.Height, frmControls.ClipBoardTemp.hdc, 0, 0, vbSrcCopy
    frmMain.ActiveForm.BufferSelected.Refresh
    frmMain.ActiveForm.BufferSelected.CurrentX = frmMain.ActiveForm.Buffer.CurrentX - frmMain.ActiveForm.SelectArea.Left * (100 / frmMain.ActiveForm.GetZoomFactor)
    frmMain.ActiveForm.BufferSelected.CurrentY = frmMain.ActiveForm.Buffer.CurrentY - frmMain.ActiveForm.SelectArea.Top * (100 / frmMain.ActiveForm.GetZoomFactor)
    
    
    frmMain.ActiveForm.SelectedBack.Left = frmMain.ActiveForm.SelectArea.Left * (100 / frmMain.ActiveForm.GetZoomFactor)
    frmMain.ActiveForm.SelectedBack.Top = frmMain.ActiveForm.SelectArea.Top * (100 / frmMain.ActiveForm.GetZoomFactor)
    frmMain.ActiveForm.SelectedBack.Width = frmMain.ActiveForm.SelectArea.Width * (100 / frmMain.ActiveForm.GetZoomFactor)
    frmMain.ActiveForm.SelectedBack.Height = frmMain.ActiveForm.SelectArea.Height * (100 / frmMain.ActiveForm.GetZoomFactor)
    'frmMain.ActiveForm.SelectedBack.BackColor = frmMain.SelColor(1).BackColor
    BitBlt frmMain.ActiveForm.SelectedBack.hdc, 0, 0, frmMain.ActiveForm.BufferSelected.Width, frmMain.ActiveForm.BufferSelected.Height, frmMain.ActiveForm.Buffer.hdc, frmMain.ActiveForm.SelectArea.Left * (100 / frmMain.ActiveForm.GetZoomFactor), frmMain.ActiveForm.SelectArea.Top * (100 / frmMain.ActiveForm.GetZoomFactor), vbSrcCopy
    frmMain.ActiveForm.SelectedBack.Refresh
    frmMain.ActiveForm.SelectedBack.CurrentX = frmMain.ActiveForm.Buffer.CurrentX - frmMain.ActiveForm.SelectArea.Left * (100 / frmMain.ActiveForm.GetZoomFactor)
    frmMain.ActiveForm.SelectedBack.CurrentY = frmMain.ActiveForm.Buffer.CurrentY - frmMain.ActiveForm.SelectArea.Top * (100 / frmMain.ActiveForm.GetZoomFactor)
    
    frmMain.ActiveForm.SelectArea.Visible = True
    
    OriginalSelX = X2 ' / frmMain.ActiveForm.GetZoomFactor * 100
    OriginalSelY = Y2 '/ frmMain.ActiveForm.GetZoomFactor * 100
    
    UpdateArea frmMain.ActiveForm.Buffer, X, Y, frmMain.ActiveForm.GetZoomFactor
    'UpdateArea frmMain.ActiveForm.Buffer, 0, 0, frmMain.ActiveForm.GetZoomFactor
    
End Sub

Private Sub PasteSpecial_Click()
    Dim f As New frmPaint
    
    frmControls.ClipBoardTemp.Picture = Clipboard.GetData(vbCFBitmap)
    
    f.PaintArea.Width = frmControls.ClipBoardTemp.ScaleWidth + 2
    f.PaintArea.Height = frmControls.ClipBoardTemp.ScaleHeight + 2
    
    f.Buffer.Width = frmControls.ClipBoardTemp.ScaleWidth + 2
    f.Buffer.Height = frmControls.ClipBoardTemp.ScaleHeight + 2
    
    f.Caption = "Untitled - 100%"
    f.Tag = "Untitled"
    f.Show

    BitBlt frmMain.ActiveForm.Buffer.hdc, 0, 0, frmControls.ClipBoardTemp.ScaleWidth, frmControls.ClipBoardTemp.ScaleHeight, frmControls.ClipBoardTemp.hdc, 0, 0, vbSrcCopy
    BitBlt frmMain.ActiveForm.PaintArea.hdc, 0, 0, frmControls.ClipBoardTemp.ScaleWidth, frmControls.ClipBoardTemp.ScaleHeight, frmControls.ClipBoardTemp.hdc, 0, 0, vbSrcCopy
    frmMain.ActiveForm.PaintArea.Refresh
    frmMain.ActiveForm.Buffer.Refresh
    frmMain.LeftBar.SetFocus
    
End Sub

Private Sub Picture1_Click()
    frmMain.Enabled = False
    frmNewSwatch.Show
    
End Sub

Private Sub RestoreColors_Click()
    Me.SelColor(0).BackColor = 0
    Me.SelColor(1).BackColor = RGB(255, 255, 255)
    
    Me.ColorBlend(0).BackColor = Me.SelColor(0).BackColor
    Me.ColorBlend(1).BackColor = Me.SelColor(1).BackColor
    SelectedColor = 0
    SetColorBars
    
    DrawPreviewGradient
    
    On Error Resume Next
    frmMain.ActiveForm.Buffer.ForeColor = frmMain.SelColor(0).BackColor
    frmMain.ActiveForm.TextInput.ForeColor = frmMain.SelColor(0).BackColor
    frmMain.ActiveForm.TextInput.BackColor = frmMain.SelColor(1).BackColor
    
End Sub

Private Sub Save_Click()
 Dim FileName As String

    If frmMain.ActiveForm.Buffer.Tag <> "" Then
        SavePicture frmMain.ActiveForm.Buffer.Image, frmMain.ActiveForm.Buffer.Tag
        frmMain.ActiveForm.SetDirtyFalse
    Else
        FileName = GetSaveName("Save As...")
         
        If FileName <> "" Then
            SavePicture frmMain.ActiveForm.Buffer.Image, FileName
            frmMain.ActiveForm.Buffer.Tag = FileName
            frmMain.ActiveForm.Caption = FileName & " - " & frmMain.ActiveForm.GetZoomFactor & "%"
            frmMain.ActiveForm.SetDirtyFalse
        End If
    End If
End Sub

Private Sub SaveAs_Click()
 Dim FileName As String

    FileName = GetSaveName("Save As...")
     
    If FileName <> "" Then
        SavePicture frmMain.ActiveForm.Buffer.Image, FileName
        frmMain.ActiveForm.Buffer.Tag = FileName
        frmMain.ActiveForm.Caption = FileName & " - " & frmMain.ActiveForm.GetZoomFactor & "%"
        frmMain.ActiveForm.SetDirtyFalse
    End If

End Sub

Private Sub ScrollSwatch_Change()
    Me.SwatchScroll.SetFocus
    Me.SwatchScroll.Top = -Me.ScrollSwatch.Value
    
End Sub

Private Sub ScrollSwatch_GotFocus()
    Me.SwatchScroll.SetFocus
End Sub

Private Sub ScrollSwatch_Scroll()
    Me.SwatchScroll.SetFocus
    Me.SwatchScroll.Top = -Me.ScrollSwatch.Value
    
End Sub

Private Sub SelColor_Click(Index As Integer)
    SelColorIndex = Index
    Me.Enabled = False
    frmColor.Show
    
    frmColor.OldColor.BackColor = frmMain.SelColor(Index).BackColor
    
    If Index = 0 Then
        frmColor.lblTitle.Caption = "Select foreground color:"
    Else
        frmColor.lblTitle.Caption = "Select background color:"
    End If
    
    
End Sub

Private Sub SelectAll_Click()
    frmMain.ActiveForm.BufferSelected.Cls
    frmMain.ActiveForm.BufferSelected.Picture = Nothing
    
    frmMain.ActiveForm.SelectArea.Left = 0
    frmMain.ActiveForm.SelectArea.Top = 0
    frmMain.ActiveForm.SelectArea.Width = frmMain.ActiveForm.PaintArea.ScaleWidth
    frmMain.ActiveForm.SelectArea.Height = frmMain.ActiveForm.PaintArea.ScaleHeight
    
    frmMain.ActiveForm.BufferSelected.Left = frmMain.ActiveForm.SelectArea.Left * (100 / frmMain.ActiveForm.GetZoomFactor)
    frmMain.ActiveForm.BufferSelected.Top = frmMain.ActiveForm.SelectArea.Top * (100 / frmMain.ActiveForm.GetZoomFactor)
    frmMain.ActiveForm.BufferSelected.Width = frmMain.ActiveForm.SelectArea.Width * (100 / frmMain.ActiveForm.GetZoomFactor)
    frmMain.ActiveForm.BufferSelected.Height = frmMain.ActiveForm.SelectArea.Height * (100 / frmMain.ActiveForm.GetZoomFactor)
    BitBlt frmMain.ActiveForm.BufferSelected.hdc, 0, 0, frmMain.ActiveForm.BufferSelected.Width, frmMain.ActiveForm.BufferSelected.Height, frmMain.ActiveForm.Buffer.hdc, frmMain.ActiveForm.SelectArea.Left * (100 / frmMain.ActiveForm.GetZoomFactor), frmMain.ActiveForm.SelectArea.Top * (100 / frmMain.ActiveForm.GetZoomFactor), vbSrcCopy
    frmMain.ActiveForm.BufferSelected.Refresh
    frmMain.ActiveForm.BufferSelected.CurrentX = frmMain.ActiveForm.Buffer.CurrentX - frmMain.ActiveForm.SelectArea.Left * (100 / frmMain.ActiveForm.GetZoomFactor)
    frmMain.ActiveForm.BufferSelected.CurrentY = frmMain.ActiveForm.Buffer.CurrentY - frmMain.ActiveForm.SelectArea.Top * (100 / frmMain.ActiveForm.GetZoomFactor)
    
    frmMain.ActiveForm.SelectedBack.Left = frmMain.ActiveForm.SelectArea.Left * (100 / frmMain.ActiveForm.GetZoomFactor)
    frmMain.ActiveForm.SelectedBack.Top = frmMain.ActiveForm.SelectArea.Top * (100 / frmMain.ActiveForm.GetZoomFactor)
    frmMain.ActiveForm.SelectedBack.Width = frmMain.ActiveForm.SelectArea.Width * (100 / frmMain.ActiveForm.GetZoomFactor)
    frmMain.ActiveForm.SelectedBack.Height = frmMain.ActiveForm.SelectArea.Height * (100 / frmMain.ActiveForm.GetZoomFactor)
    frmMain.ActiveForm.SelectedBack.BackColor = frmMain.SelColor(1).BackColor
    
    frmMain.ActiveForm.SelectedBack.Refresh
    frmMain.ActiveForm.SelectedBack.CurrentX = frmMain.ActiveForm.Buffer.CurrentX - frmMain.ActiveForm.SelectArea.Left * (100 / frmMain.ActiveForm.GetZoomFactor)
    frmMain.ActiveForm.SelectedBack.CurrentY = frmMain.ActiveForm.Buffer.CurrentY - frmMain.ActiveForm.SelectArea.Top * (100 / frmMain.ActiveForm.GetZoomFactor)
    
    
    OriginalSelX = 0
    OriginalSelY = 0
    
    frmMain.ActiveForm.SelectArea.Visible = True
    
End Sub

Private Sub ShiftColors_Click()
    Dim TmpColor As Long
    
    SelectedColor = 0
    
    TmpColor = Me.SelColor(0).BackColor
    Me.SelColor(0).BackColor = Me.SelColor(1).BackColor
    Me.SelColor(1).BackColor = TmpColor
    
    Me.ColorBlend(0).BackColor = Me.SelColor(0).BackColor
    Me.ColorBlend(1).BackColor = Me.SelColor(1).BackColor
    SetColorBars
    
    DrawPreviewGradient
    
    On Error Resume Next
    frmMain.ActiveForm.Buffer.ForeColor = frmMain.SelColor(0).BackColor
    frmMain.ActiveForm.TextInput.ForeColor = frmMain.SelColor(0).BackColor
    frmMain.ActiveForm.TextInput.BackColor = frmMain.SelColor(1).BackColor

End Sub

Public Function GetSelectedColor()
    GetSelectedColor = SelectedColor
    
End Function

Private Sub Swatch_Click(Index As Integer)
    Me.SelColor(SelectedColor).BackColor = Me.Swatch(Index).BackColor
    Me.ColorBlend(0).BackColor = Me.SelColor(0).BackColor
    Me.ColorBlend(1).BackColor = Me.SelColor(1).BackColor
    SetColorBars
    DrawPreviewGradient
    
    On Error Resume Next
    frmMain.ActiveForm.Buffer.ForeColor = frmMain.SelColor(0).BackColor
    frmMain.ActiveForm.TextInput.ForeColor = frmMain.SelColor(0).BackColor
    frmMain.ActiveForm.TextInput.BackColor = frmMain.SelColor(1).BackColor
    
End Sub

Private Sub SwatchScroll_Click()
    frmMain.Enabled = False
    frmNewSwatch.Show
    
End Sub

Private Sub Undo_Click()
    Dim mDC As Long, mBitmap As Long
    
    mDC = CreateCompatibleDC(GetDC(0))
    If frmMain.ActiveForm.SelectArea.Visible = False Then
        mBitmap = CreateCompatibleBitmap(GetDC(0), frmMain.ActiveForm.Buffer.ScaleWidth, frmMain.ActiveForm.Buffer.ScaleHeight)
        SelectObject mDC, mBitmap
        
        BitBlt mDC, 0, 0, frmMain.ActiveForm.Buffer.ScaleWidth, frmMain.ActiveForm.Buffer.ScaleHeight, frmMain.ActiveForm.Buffer.hdc, 0, 0, vbSrcCopy
        BitBlt frmMain.ActiveForm.Buffer.hdc, 0, 0, frmMain.ActiveForm.Buffer.ScaleWidth, frmMain.ActiveForm.Buffer.ScaleHeight, frmMain.ActiveForm.Undo.hdc, 0, 0, vbSrcCopy
        BitBlt frmMain.ActiveForm.Undo.hdc, 0, 0, frmMain.ActiveForm.Buffer.ScaleWidth, frmMain.ActiveForm.Buffer.ScaleHeight, mDC, 0, 0, vbSrcCopy
        UpdateArea frmMain.ActiveForm.Buffer, 0, 0, frmMain.ActiveForm.GetZoomFactor
    Else
        mBitmap = CreateCompatibleBitmap(GetDC(0), frmMain.ActiveForm.BufferSelected.ScaleWidth, frmMain.ActiveForm.BufferSelected.ScaleHeight)
        SelectObject mDC, mBitmap
        
        BitBlt mDC, 0, 0, frmMain.ActiveForm.BufferSelected.ScaleWidth, frmMain.ActiveForm.BufferSelected.ScaleHeight, frmMain.ActiveForm.BufferSelected.hdc, 0, 0, vbSrcCopy
        BitBlt frmMain.ActiveForm.BufferSelected.hdc, 0, 0, frmMain.ActiveForm.BufferSelected.ScaleWidth, frmMain.ActiveForm.BufferSelected.ScaleHeight, frmMain.ActiveForm.Undo.hdc, 0, 0, vbSrcCopy
        BitBlt frmMain.ActiveForm.Undo.hdc, 0, 0, frmMain.ActiveForm.BufferSelected.ScaleWidth, frmMain.ActiveForm.BufferSelected.ScaleHeight, mDC, 0, 0, vbSrcCopy
        UpdateArea frmMain.ActiveForm.Buffer, 0, 0, frmMain.ActiveForm.GetZoomFactor
    
    End If
    DeleteDC mDC
    DeleteObject mBitmap
    
End Sub
