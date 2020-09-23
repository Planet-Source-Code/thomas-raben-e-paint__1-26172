VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   301
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   401
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   4515
      Left            =   0
      Picture         =   "frmSplash.frx":0000
      ScaleHeight     =   4485
      ScaleWidth      =   5985
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Written by: Thomas Raben (tr@vupti.com)"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   4140
         Width           =   5775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmSplash.frx":57E82
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   3360
         Width           =   5775
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Loading..."
         Height          =   255
         Left            =   60
         TabIndex        =   1
         Top             =   2880
         Width           =   5235
      End
   End
End
Attribute VB_Name = "frmSPlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
