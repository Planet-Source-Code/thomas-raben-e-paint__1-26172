VERSION 5.00
Begin VB.Form frmControls 
   BackColor       =   &H80000010&
   Caption         =   "Form1"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   482
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   609
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox MyCursor 
      Height          =   375
      Index           =   13
      Left            =   4980
      MouseIcon       =   "frmControls.frx":0000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   99
      Top             =   6300
      Width           =   375
   End
   Begin VB.PictureBox DrawToolbar 
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   13
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   329
      TabIndex        =   97
      Top             =   6240
      Visible         =   0   'False
      Width           =   4935
      Begin VB.PictureBox CmdDropDown 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   8
         Left            =   1740
         Picture         =   "frmControls.frx":08CA
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   101
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox ToolIcon 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   13
         Left            =   0
         Picture         =   "frmControls.frx":0C22
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   98
         Top             =   0
         Width           =   420
      End
      Begin VB.Label lblSelector 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100 pixel "
         Height          =   255
         Index           =   8
         Left            =   900
         TabIndex        =   102
         Tag             =   " pixel "
         Top             =   90
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Size:"
         Height          =   195
         Index           =   7
         Left            =   480
         TabIndex        =   100
         Top             =   120
         Width           =   345
      End
   End
   Begin VB.PictureBox MyCursor 
      Height          =   375
      Index           =   12
      Left            =   8460
      MouseIcon       =   "frmControls.frx":0FB1
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   84
      Top             =   5820
      Width           =   375
   End
   Begin VB.PictureBox DrawToolbar 
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   12
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   545
      TabIndex        =   82
      Top             =   5760
      Visible         =   0   'False
      Width           =   8175
      Begin VB.PictureBox CmdFontStyle 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   6180
         Picture         =   "frmControls.frx":187B
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   94
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox CmdFont 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   4080
         Picture         =   "frmControls.frx":1BD3
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   90
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox CmdButton 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   7680
         Picture         =   "frmControls.frx":1F2B
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   89
         ToolTipText     =   "Cancel Text"
         Top             =   60
         Width           =   420
      End
      Begin VB.PictureBox CmdButton 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   7200
         Picture         =   "frmControls.frx":22A6
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   88
         ToolTipText     =   "Apply Text"
         Top             =   60
         Width           =   420
      End
      Begin VB.PictureBox CmdDropDown 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   7
         Left            =   1740
         Picture         =   "frmControls.frx":261C
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   85
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox ToolIcon 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   12
         Left            =   0
         Picture         =   "frmControls.frx":2974
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   83
         Top             =   0
         Width           =   420
      End
      Begin VB.Label lblFontStyle 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FontStyle"
         Height          =   255
         Left            =   4920
         TabIndex        =   96
         Tag             =   " pixel "
         Top             =   90
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Font:"
         Height          =   195
         Index           =   8
         Left            =   4500
         TabIndex        =   95
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Font:"
         Height          =   195
         Index           =   7
         Left            =   2160
         TabIndex        =   93
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Action:"
         Height          =   195
         Index           =   6
         Left            =   6600
         TabIndex        =   92
         Top             =   120
         Width           =   495
      End
      Begin VB.Label LblFont 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fontname"
         Height          =   255
         Left            =   2580
         TabIndex        =   91
         Tag             =   " pixel "
         Top             =   90
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Size:"
         Height          =   195
         Index           =   6
         Left            =   480
         TabIndex        =   87
         Top             =   120
         Width           =   345
      End
      Begin VB.Label lblSelector 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100 pixel "
         Height          =   255
         Index           =   7
         Left            =   900
         TabIndex        =   86
         Tag             =   " pixel "
         Top             =   90
         Width           =   795
      End
   End
   Begin VB.Timer LightTimer 
      Interval        =   100
      Left            =   6000
      Top             =   5340
   End
   Begin VB.Timer BrushTimer 
      Interval        =   100
      Left            =   5460
      Top             =   480
   End
   Begin VB.PictureBox ClipBoardTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6780
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   62
      Top             =   540
      Width           =   315
   End
   Begin VB.PictureBox DrawToolbar 
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   11
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   361
      TabIndex        =   57
      Top             =   5280
      Visible         =   0   'False
      Width           =   5415
      Begin VB.PictureBox Cmd4 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   4920
         Picture         =   "frmControls.frx":2CF6
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   81
         ToolTipText     =   "Darken"
         Top             =   60
         Width           =   420
      End
      Begin VB.PictureBox Cmd4 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   4440
         Picture         =   "frmControls.frx":3074
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   80
         ToolTipText     =   "Brighten"
         Top             =   60
         Width           =   420
      End
      Begin VB.PictureBox CmdDropDown 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   3
         Left            =   1740
         Picture         =   "frmControls.frx":33F7
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   70
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox CmdDropDown 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   3540
         Picture         =   "frmControls.frx":374F
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   66
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox ToolIcon 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   11
         Left            =   0
         Picture         =   "frmControls.frx":3AA7
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   58
         Top             =   0
         Width           =   420
      End
      Begin VB.Label lblSelector 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100 pixel "
         Height          =   255
         Index           =   3
         Left            =   900
         TabIndex        =   71
         Tag             =   " pixel "
         Top             =   90
         Width           =   795
      End
      Begin VB.Label lblSelector 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0 % "
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   67
         Tag             =   " % "
         Top             =   90
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Style:"
         Height          =   195
         Index           =   5
         Left            =   3960
         TabIndex        =   61
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Pressure:"
         Height          =   195
         Index           =   4
         Left            =   2160
         TabIndex        =   60
         Top             =   120
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Size:"
         Height          =   195
         Index           =   5
         Left            =   480
         TabIndex        =   59
         Top             =   120
         Width           =   345
      End
   End
   Begin VB.PictureBox MyCursor 
      Height          =   375
      Index           =   11
      Left            =   5580
      MouseIcon       =   "frmControls.frx":3E3E
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   56
      Top             =   5340
      Width           =   375
   End
   Begin VB.Timer CircleTimer 
      Interval        =   100
      Left            =   5520
      Top             =   4800
   End
   Begin VB.Timer CmdTimer 
      Interval        =   100
      Left            =   5520
      Top             =   4320
   End
   Begin VB.PictureBox MyCursor 
      Height          =   375
      Index           =   10
      Left            =   4980
      MouseIcon       =   "frmControls.frx":4708
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   39
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox DrawToolbar 
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   10
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   297
      TabIndex        =   37
      Top             =   4800
      Visible         =   0   'False
      Width           =   4455
      Begin VB.PictureBox CmdDropDown 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   6
         Left            =   2160
         Picture         =   "frmControls.frx":4FD2
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   76
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox Cmd2 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   4020
         Picture         =   "frmControls.frx":532A
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   49
         ToolTipText     =   "Fill Only"
         Top             =   60
         Width           =   420
      End
      Begin VB.PictureBox Cmd2 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   3540
         Picture         =   "frmControls.frx":569E
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   48
         ToolTipText     =   "Border and Fill"
         Top             =   60
         Width           =   420
      End
      Begin VB.PictureBox Cmd2 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   3060
         Picture         =   "frmControls.frx":5A21
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   47
         ToolTipText     =   "Border Only"
         Top             =   60
         Width           =   420
      End
      Begin VB.PictureBox ToolIcon 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   10
         Left            =   0
         Picture         =   "frmControls.frx":5D97
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   38
         Top             =   0
         Width           =   420
      End
      Begin VB.Label lblSelector 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100 pixel "
         Height          =   255
         Index           =   6
         Left            =   1320
         TabIndex        =   77
         Tag             =   " pixel "
         Top             =   90
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bordersize:"
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   46
         Top             =   120
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Style:"
         Height          =   195
         Index           =   3
         Left            =   2580
         TabIndex        =   45
         Top             =   120
         Width           =   390
      End
   End
   Begin VB.PictureBox MyCursor 
      Height          =   375
      Index           =   9
      Left            =   4980
      MouseIcon       =   "frmControls.frx":6128
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   36
      Top             =   4320
      Width           =   375
   End
   Begin VB.PictureBox DrawToolbar 
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   9
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   297
      TabIndex        =   34
      Top             =   4320
      Visible         =   0   'False
      Width           =   4455
      Begin VB.PictureBox CmdDropDown 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   5
         Left            =   2160
         Picture         =   "frmControls.frx":69F2
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   74
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox Cmd 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   4020
         Picture         =   "frmControls.frx":6D4A
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   44
         ToolTipText     =   "Fill Only"
         Top             =   60
         Width           =   420
      End
      Begin VB.PictureBox Cmd 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   3540
         Picture         =   "frmControls.frx":70BB
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   43
         ToolTipText     =   "Border and Fill"
         Top             =   60
         Width           =   420
      End
      Begin VB.PictureBox Cmd 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   3060
         Picture         =   "frmControls.frx":743E
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   42
         ToolTipText     =   "Border Only"
         Top             =   60
         Width           =   420
      End
      Begin VB.PictureBox ToolIcon 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   9
         Left            =   0
         Picture         =   "frmControls.frx":77AF
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   35
         Top             =   0
         Width           =   420
      End
      Begin VB.Label lblSelector 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100 pixel "
         Height          =   255
         Index           =   5
         Left            =   1320
         TabIndex        =   75
         Tag             =   " pixel "
         Top             =   90
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Style:"
         Height          =   195
         Index           =   2
         Left            =   2580
         TabIndex        =   41
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bordersize:"
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   40
         Top             =   120
         Width           =   780
      End
   End
   Begin VB.PictureBox DrawToolbar 
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   8
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   225
      TabIndex        =   30
      Top             =   3840
      Visible         =   0   'False
      Width           =   3375
      Begin VB.PictureBox ColorPick 
         BackColor       =   &H00000000&
         Height          =   315
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   50
         Top             =   60
         Width           =   435
      End
      Begin VB.PictureBox ToolIcon 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   8
         Left            =   0
         Picture         =   "frmControls.frx":7B46
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   31
         Top             =   0
         Width           =   420
      End
      Begin VB.Label lblBlue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   2940
         TabIndex        =   55
         Top             =   90
         Width           =   375
      End
      Begin VB.Label lblGreen 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   2520
         TabIndex        =   54
         Top             =   90
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "RGB:"
         Height          =   195
         Index           =   2
         Left            =   1620
         TabIndex        =   53
         Top             =   120
         Width           =   390
      End
      Begin VB.Label lblRed 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   2100
         TabIndex        =   52
         Top             =   90
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Color:"
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   51
         Top             =   120
         Width           =   405
      End
   End
   Begin VB.PictureBox MyCursor 
      Height          =   375
      Index           =   8
      Left            =   4980
      MouseIcon       =   "frmControls.frx":7ECE
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   29
      Top             =   3840
      Width           =   375
   End
   Begin VB.PictureBox MyCursor 
      Height          =   375
      Index           =   7
      Left            =   4980
      MouseIcon       =   "frmControls.frx":8798
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   28
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox MyCursor 
      Height          =   375
      Index           =   6
      Left            =   4980
      MouseIcon       =   "frmControls.frx":9062
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   27
      Top             =   2880
      Width           =   375
   End
   Begin VB.PictureBox MyCursor 
      Height          =   375
      Index           =   5
      Left            =   4980
      MouseIcon       =   "frmControls.frx":992C
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   26
      Top             =   2400
      Width           =   375
   End
   Begin VB.PictureBox MyCursor 
      Height          =   375
      Index           =   4
      Left            =   4980
      MouseIcon       =   "frmControls.frx":A1F6
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   25
      Top             =   1920
      Width           =   375
   End
   Begin VB.PictureBox MyCursor 
      Height          =   375
      Index           =   3
      Left            =   4980
      MouseIcon       =   "frmControls.frx":AAC0
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   24
      Top             =   1440
      Width           =   375
   End
   Begin VB.PictureBox MyCursor 
      Height          =   375
      Index           =   2
      Left            =   4980
      MouseIcon       =   "frmControls.frx":B38A
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   23
      Top             =   960
      Width           =   375
   End
   Begin VB.PictureBox MyCursor 
      Height          =   375
      Index           =   1
      Left            =   4980
      MouseIcon       =   "frmControls.frx":BC54
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   22
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox MyCursor 
      Height          =   375
      Index           =   0
      Left            =   4980
      MouseIcon       =   "frmControls.frx":C51E
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   21
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox DrawToolbar 
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   7
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   329
      TabIndex        =   16
      Top             =   3360
      Visible         =   0   'False
      Width           =   4935
      Begin VB.PictureBox ToolIcon 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   7
         Left            =   0
         Picture         =   "frmControls.frx":CDE8
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   17
         Top             =   0
         Width           =   420
      End
   End
   Begin VB.PictureBox DrawToolbar 
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   6
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   14
      Top             =   2880
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox CmdDropDown 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   1740
         Picture         =   "frmControls.frx":D178
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   68
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox CmdDropDown 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   0
         Left            =   3540
         Picture         =   "frmControls.frx":D4D0
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   65
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox ToolIcon 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   6
         Left            =   0
         Picture         =   "frmControls.frx":D828
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   15
         Top             =   0
         Width           =   420
      End
      Begin VB.Label lblSelector 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100 pixel "
         Height          =   255
         Index           =   2
         Left            =   900
         TabIndex        =   69
         Tag             =   " pixel "
         Top             =   90
         Width           =   795
      End
      Begin VB.Label lblSelector 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0 % "
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   64
         Tag             =   " % "
         Top             =   90
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Size:"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   20
         Top             =   120
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Pressure:"
         Height          =   195
         Index           =   1
         Left            =   2160
         TabIndex        =   19
         Top             =   120
         Width           =   660
      End
   End
   Begin VB.PictureBox DrawToolbar 
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   5
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   329
      TabIndex        =   12
      Top             =   2400
      Visible         =   0   'False
      Width           =   4935
      Begin VB.PictureBox ToolIcon 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   5
         Left            =   0
         Picture         =   "frmControls.frx":DBBC
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   13
         Top             =   0
         Width           =   420
      End
   End
   Begin VB.PictureBox DrawToolbar 
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   4
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   109
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   1635
      Begin VB.PictureBox ToolIcon 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   4
         Left            =   0
         Picture         =   "frmControls.frx":DF3F
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   11
         Top             =   0
         Width           =   420
      End
      Begin VB.Label lblAngle 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N/A"
         Height          =   255
         Left            =   1020
         TabIndex        =   33
         Top             =   90
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Angle:"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   32
         Top             =   120
         Width           =   450
      End
   End
   Begin VB.PictureBox DrawToolbar 
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   3
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   329
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   4935
      Begin VB.PictureBox ToolIcon 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   3
         Left            =   0
         Picture         =   "frmControls.frx":E2E8
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   7
         Top             =   0
         Width           =   420
      End
   End
   Begin VB.PictureBox DrawToolbar 
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   2
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   1815
      Begin VB.PictureBox ToolIcon 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   2
         Left            =   0
         Picture         =   "frmControls.frx":E689
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   6
         Top             =   0
         Width           =   420
      End
      Begin VB.Label lblZoom 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100 %"
         Height          =   255
         Left            =   1020
         TabIndex        =   63
         Top             =   90
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Zoom:"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   18
         Top             =   120
         Width           =   450
      End
   End
   Begin VB.PictureBox DrawToolbar 
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   1
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   249
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   3735
      Begin VB.PictureBox Cmd3 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   2760
         Picture         =   "frmControls.frx":EA17
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   79
         ToolTipText     =   "Freehand"
         Top             =   60
         Width           =   420
      End
      Begin VB.PictureBox Cmd3 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   3240
         Picture         =   "frmControls.frx":ED8D
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   78
         ToolTipText     =   "Straight Line"
         Top             =   60
         Width           =   420
      End
      Begin VB.PictureBox CmdDropDown 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   4
         Left            =   1860
         Picture         =   "frmControls.frx":F0FD
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   72
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox ToolIcon 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   1
         Left            =   0
         Picture         =   "frmControls.frx":F455
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   5
         Top             =   0
         Width           =   420
      End
      Begin VB.Label lblSelector 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100 pixel "
         Height          =   255
         Index           =   4
         Left            =   1020
         TabIndex        =   73
         Tag             =   " pixel "
         Top             =   90
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Style:"
         Height          =   195
         Index           =   0
         Left            =   2280
         TabIndex        =   9
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Brush:"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   8
         Top             =   120
         Width           =   450
      End
   End
   Begin VB.PictureBox DrawToolbar 
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   0
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   329
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4935
      Begin VB.PictureBox ToolIcon 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   0
         Left            =   0
         Picture         =   "frmControls.frx":F7DD
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   4
         Top             =   0
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ButtonDown(10) As Boolean
Dim LastClick As Integer
Dim FontDown As Boolean
Dim FontStyleDown As Boolean

Private Sub BrushTimer_Timer()
    CheckFlatButtonsBrush
End Sub

Private Sub CircleTimer_Timer()
    CheckFlatButtonsCircle
End Sub

Private Sub cmbZoom_Click()
    On Error Resume Next
    'MsgBox frmMain.ActiveForm.Caption
End Sub

Private Sub Cmd_Click(Index As Integer)
    SelectToolRect Index
End Sub

Private Sub Cmd2_Click(Index As Integer)
    SelectToolRect Index
End Sub

Public Sub ClickDropDown(Index As Integer)
    Dim i As Integer
    
    If ButtonDown(Index) = False Then
        LastClick = Index
        For i = 0 To CmdDropDown.UBound
            If i <> Index Then
                ButtonDown(i) = False
                DrawButton i, False
            End If
        Next i
        ButtonDown(Index) = True
        DrawButton Index, True
        ShowScale Index
    Else
        LastClick = -1
        UnClickButton
    End If
End Sub

Public Sub UnClickButton()
    Dim i As Integer
    
    For i = 0 To CmdDropDown.UBound
        ButtonDown(i) = False
        DrawButton i, False
        frmSelector.Hide
    Next i
    
End Sub

Public Sub DrawButton(i As Integer, Down As Boolean)
    If Down = False Then
        frmControls.CmdDropDown(i).BackColor = vb3DFace
        frmControls.CmdDropDown(i).Line (0, 0)-(frmControls.CmdDropDown(i).ScaleWidth, 0), vb3DHighlight
        frmControls.CmdDropDown(i).Line (0, 0)-(0, frmControls.CmdDropDown(i).ScaleHeight), vb3DHighlight
        frmControls.CmdDropDown(i).Line (frmControls.CmdDropDown(i).ScaleWidth - 1, 1)-(frmControls.CmdDropDown(i).ScaleWidth - 1, frmControls.CmdDropDown(i).ScaleHeight), vb3DShadow
        frmControls.CmdDropDown(i).Line (1, frmControls.CmdDropDown(i).ScaleHeight - 1)-(frmControls.CmdDropDown(i).ScaleWidth, frmControls.CmdDropDown(i).ScaleHeight - 1), vb3DShadow
        If FontDown = True Then CmdFont_Click
        If FontStyleDown = True Then CmdFontStyle_Click
    Else
        frmControls.CmdDropDown(i).BackColor = vbScrollBars
        frmControls.CmdDropDown(i).Line (0, 0)-(frmControls.CmdDropDown(i).ScaleWidth, 0), vb3DShadow
        frmControls.CmdDropDown(i).Line (0, 0)-(0, frmControls.CmdDropDown(i).ScaleHeight), vb3DShadow
        frmControls.CmdDropDown(i).Line (frmControls.CmdDropDown(i).ScaleWidth - 1, 1)-(frmControls.CmdDropDown(i).ScaleWidth - 1, frmControls.CmdDropDown(i).ScaleHeight), vb3DHighlight
        frmControls.CmdDropDown(i).Line (1, frmControls.CmdDropDown(i).ScaleHeight - 1)-(frmControls.CmdDropDown(i).ScaleWidth, frmControls.CmdDropDown(i).ScaleHeight - 1), vb3DHighlight
    End If
    
End Sub

Private Sub ShowScale(Index As Integer)
    Dim Xpos As Long
    Xpos = ((Me.CmdDropDown(Index).Left + Me.CmdDropDown(Index).Width + 16) * 15) - frmSelector.Width
    
    If Xpos < frmMain.LeftBar.Width - 60 Then Xpos = frmMain.LeftBar.Width - 60
    
    frmSelector.Move Xpos, 28 * 15
    
    frmSelector.ScrollBlock.Left = frmSelector.ScrollBar.Left + (frmSelector.ScrollBar.Width - frmSelector.ScrollBlock.Width) * (SelMethod(Index).Current / (SelMethod(Index).Max - SelMethod(Index).Min)) - SelMethod(Index).Min
       
    Me.lblSelector(Index).Caption = CInt(SelMethod(Index).Current) & Me.lblSelector(Index).Tag
       
    frmSelector.Show
    frmSelector.ScrollBlock.SetFocus
    frmSelector.Tag = Index
    
End Sub

Private Sub ShowFonts()
    Dim Xpos As Long
    
    Xpos = ((Me.CmdFont.Left + Me.CmdFont.Width + Me.DrawToolbar(CurrentButton).Left + 1) * 15) - frmFonts.Width
    If Xpos < frmMain.LeftBar.Width - 60 Then Xpos = frmMain.LeftBar.Width - 60
    frmFonts.Move Xpos, 26 * 15
    frmFonts.Show
    
End Sub

Private Sub ShowFontStyle()
    Dim Xpos As Long
    
    Xpos = ((Me.CmdFontStyle.Left + Me.CmdFontStyle.Width + Me.DrawToolbar(CurrentButton).Left + 1) * 15) - frmFontStyle.Width
    If Xpos < frmMain.LeftBar.Width - 60 Then Xpos = frmMain.LeftBar.Width - 60
    frmFontStyle.Move Xpos, 26 * 15
    frmFontStyle.Show
    
End Sub

Private Sub Cmd3_Click(Index As Integer)
    SelectToolBrush Index
    
End Sub


Private Sub Cmd4_Click(Index As Integer)
    SelectToolLight Index
End Sub

Private Sub CmdButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmControls.CmdButton(Index).BackColor = vbScrollBars
    frmControls.CmdButton(Index).Line (0, 0)-(frmControls.CmdButton(Index).ScaleWidth, 0), vb3DShadow
    frmControls.CmdButton(Index).Line (0, 0)-(0, frmControls.CmdButton(Index).ScaleHeight), vb3DShadow
    frmControls.CmdButton(Index).Line (frmControls.CmdButton(Index).ScaleWidth - 1, 1)-(frmControls.CmdButton(Index).ScaleWidth - 1, frmControls.CmdButton(Index).ScaleHeight), vb3DHighlight
    frmControls.CmdButton(Index).Line (1, frmControls.CmdButton(Index).ScaleHeight - 1)-(frmControls.CmdButton(Index).ScaleWidth, frmControls.CmdButton(Index).ScaleHeight - 1), vb3DHighlight

End Sub

Private Sub CmdButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmControls.CmdButton(Index).BackColor = vb3DFace
    frmControls.CmdButton(Index).Line (0, 0)-(frmControls.CmdButton(Index).ScaleWidth, 0), vb3DHighlight
    frmControls.CmdButton(Index).Line (0, 0)-(0, frmControls.CmdButton(Index).ScaleHeight), vb3DHighlight
    frmControls.CmdButton(Index).Line (frmControls.CmdButton(Index).ScaleWidth - 1, 1)-(frmControls.CmdButton(Index).ScaleWidth - 1, frmControls.CmdButton(Index).ScaleHeight), vb3DShadow
    frmControls.CmdButton(Index).Line (1, frmControls.CmdButton(Index).ScaleHeight - 1)-(frmControls.CmdButton(Index).ScaleWidth, frmControls.CmdButton(Index).ScaleHeight - 1), vb3DShadow
    
    Select Case Index
        Case 0
            ApplyText
        Case 1
            On Error Resume Next
            frmMain.ActiveForm.TextInput.Text = ""
            frmMain.ActiveForm.lblTextSize.Caption = "M"
            frmMain.ActiveForm.TextInput.Visible = False
    End Select
    
End Sub

Private Sub CmdDropDown_Click(Index As Integer)
    ClickDropDown Index
        
End Sub

Private Sub CmdFont_Click()
    ShowHideFonts
End Sub

Public Sub ShowHideFonts()
    Dim i As Integer

    If FontDown = False Then
        
        frmControls.CmdFont.BackColor = vbScrollBars
        frmControls.CmdFont.Line (0, 0)-(frmControls.CmdFont.ScaleWidth, 0), vb3DShadow
        frmControls.CmdFont.Line (0, 0)-(0, frmControls.CmdFont.ScaleHeight), vb3DShadow
        frmControls.CmdFont.Line (frmControls.CmdFont.ScaleWidth - 1, 1)-(frmControls.CmdFont.ScaleWidth - 1, frmControls.CmdFont.ScaleHeight), vb3DHighlight
        frmControls.CmdFont.Line (1, frmControls.CmdFont.ScaleHeight - 1)-(frmControls.CmdFont.ScaleWidth, frmControls.CmdFont.ScaleHeight - 1), vb3DHighlight
        ShowFonts
        
        LastClick = -1
        For i = 0 To CmdDropDown.UBound
            ButtonDown(i) = False
            DrawButton i, False
        Next i
        
        frmSelector.Hide
        FontStyleDown = True
        ShowHideFontStyle
        frmFontStyle.Hide
        
        FontDown = True
    Else
        FontDown = False
        frmControls.CmdFont.BackColor = vb3DFace
        frmControls.CmdFont.Line (0, 0)-(frmControls.CmdFont.ScaleWidth, 0), vb3DHighlight
        frmControls.CmdFont.Line (0, 0)-(0, frmControls.CmdFont.ScaleHeight), vb3DHighlight
        frmControls.CmdFont.Line (frmControls.CmdFont.ScaleWidth - 1, 1)-(frmControls.CmdFont.ScaleWidth - 1, frmControls.CmdFont.ScaleHeight), vb3DShadow
        frmControls.CmdFont.Line (1, frmControls.CmdFont.ScaleHeight - 1)-(frmControls.CmdFont.ScaleWidth, frmControls.CmdFont.ScaleHeight - 1), vb3DShadow
        frmFonts.Hide
        
    End If
End Sub

Public Sub ShowHideFontStyle()
    Dim i As Integer

    If FontStyleDown = False Then
        
        frmControls.CmdFontStyle.BackColor = vbScrollBars
        frmControls.CmdFontStyle.Line (0, 0)-(frmControls.CmdFontStyle.ScaleWidth, 0), vb3DShadow
        frmControls.CmdFontStyle.Line (0, 0)-(0, frmControls.CmdFontStyle.ScaleHeight), vb3DShadow
        frmControls.CmdFontStyle.Line (frmControls.CmdFontStyle.ScaleWidth - 1, 1)-(frmControls.CmdFontStyle.ScaleWidth - 1, frmControls.CmdFontStyle.ScaleHeight), vb3DHighlight
        frmControls.CmdFontStyle.Line (1, frmControls.CmdFontStyle.ScaleHeight - 1)-(frmControls.CmdFontStyle.ScaleWidth, frmControls.CmdFontStyle.ScaleHeight - 1), vb3DHighlight
        ShowFontStyle
        
        LastClick = -1
        For i = 0 To CmdDropDown.UBound
            ButtonDown(i) = False
            DrawButton i, False
        Next i
        
        frmSelector.Hide
        FontDown = False
        frmFonts.Hide
        
        FontStyleDown = True
    Else
        FontStyleDown = False
        frmControls.CmdFontStyle.BackColor = vb3DFace
        frmControls.CmdFontStyle.Line (0, 0)-(frmControls.CmdFontStyle.ScaleWidth, 0), vb3DHighlight
        frmControls.CmdFontStyle.Line (0, 0)-(0, frmControls.CmdFontStyle.ScaleHeight), vb3DHighlight
        frmControls.CmdFontStyle.Line (frmControls.CmdFontStyle.ScaleWidth - 1, 1)-(frmControls.CmdFontStyle.ScaleWidth - 1, frmControls.CmdFontStyle.ScaleHeight), vb3DShadow
        frmControls.CmdFontStyle.Line (1, frmControls.CmdFontStyle.ScaleHeight - 1)-(frmControls.CmdFontStyle.ScaleWidth, frmControls.CmdFontStyle.ScaleHeight - 1), vb3DShadow
        frmFontStyle.Hide
        
    End If
End Sub

Private Sub CmdFontStyle_Click()
    ShowHideFontStyle
End Sub

Private Sub CmdTimer_Timer()
    CheckFlatButtonsRect
End Sub
    
Private Sub ApplyText()
    Dim tempstr As String
    'paste the goddamn text!
    On Error GoTo NoWindow
    tempstr = frmMain.ActiveForm.Caption
    
    If frmMain.ActiveForm.TextInput.Visible = False Then GoTo NoWindow
    
    frmMain.ActiveForm.Buffer.CurrentX = frmMain.ActiveForm.TextInput.Left * (100 / frmMain.ActiveForm.GetZoomFactor)
    frmMain.ActiveForm.Buffer.CurrentY = frmMain.ActiveForm.TextInput.Top * (100 / frmMain.ActiveForm.GetZoomFactor)
    
    tempstr = frmMain.ActiveForm.TextInput.Text
    
    Do While InStr(1, tempstr, vbCrLf) > 0
        frmMain.ActiveForm.Buffer.Print Mid(tempstr, 1, InStr(1, tempstr, vbCrLf) - 1)
        tempstr = Mid(tempstr, InStr(1, tempstr, vbCrLf) + 2)
        frmMain.ActiveForm.Buffer.CurrentX = frmMain.ActiveForm.TextInput.Left * (100 / frmMain.ActiveForm.GetZoomFactor)
    Loop
    frmMain.ActiveForm.Buffer.Print tempstr
    
    frmMain.ActiveForm.TextInput.Text = ""
    frmMain.ActiveForm.lblTextSize.Caption = "M"
    frmMain.ActiveForm.TextInput.Visible = False
    
    frmMain.ActiveForm.Buffer.Refresh
    UpdateArea frmMain.ActiveForm.Buffer, frmMain.ActiveForm.Buffer.CurrentX * (100 / frmMain.ActiveForm.GetZoomFactor), frmMain.ActiveForm.Buffer.CurrentY * (100 / frmMain.ActiveForm.GetZoomFactor), frmMain.ActiveForm.GetZoomFactor
    
    Exit Sub
    
NoWindow:
    
    Exit Sub
End Sub

Private Sub Form_Load()
    LastClick = -1
End Sub

Private Sub LightTimer_Timer()
    CheckFlatButtonsLight
End Sub
