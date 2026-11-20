VERSION 5.00
Begin VB.Form frmIntro
   BackColor       =   &H00F8F6F2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nokia 5110 Bitmap Studio  |  Rezultat pe LCD"
   ClientHeight    =   10410
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10800
   ControlBox      =   -1  'True
   LinkTopic       =   "frmIntro"
   MaxButton       =   0  'False
   MinButton       =   0  'False
   ScaleHeight     =   10410
   ScaleWidth      =   10800
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStart
      Caption         =   "DESCHIDE EDITORUL"
      Default         =   -1  'True
      Height          =   630
      Left            =   3650
      TabIndex        =   0
      Top             =   9440
      Width           =   3500
   End
   Begin VB.Label lblFooter
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Imagine reala afisata pe Nokia 5110 cu FPGA Cyclone IV"
      ForeColor       =   &H006A6865&
      Height          =   320
      Left            =   700
      TabIndex        =   1
      Top             =   8890
      Width           =   9400
   End
   Begin VB.Label lblHint
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "84 x 48 pixeli  |  504 octeti  |  HEX si Verilog"
      ForeColor       =   &H006A6865&
      Height          =   320
      Left            =   700
      TabIndex        =   2
      Top             =   8510
      Width           =   9400
   End
   Begin VB.Image imgDemo
      BorderStyle     =   1  'Fixed Single
      Height          =   6650
      Left            =   1460
      Stretch         =   -1  'True
      Top             =   1430
      Width           =   7880
   End
   Begin VB.Label lblTitle
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NOKIA 5110 BITMAP STUDIO"
      ForeColor       =   &H003C3630&
      Height          =   420
      Left            =   500
      TabIndex        =   3
      Top             =   260
      Width           =   9800
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblSubtitle
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "De la fotografie la bitmap, direct pe FPGA"
      ForeColor       =   &H00777777&
      Height          =   300
      Left            =   500
      TabIndex        =   4
      Top             =   790
      Width           =   9800
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error GoTo photoMissing
    Set imgDemo.Picture = LoadPicture(App.Path & "\LCD5110_demo.bmp")
    Exit Sub
photoMissing:
    lblHint.Caption = "Fotografia demonstrativa nu a fost gasita langa proiect."
    lblFooter.Caption = "Editorul poate fi deschis in continuare."
End Sub

Private Sub cmdStart_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub imgDemo_Click()
    cmdStart_Click
End Sub
