VERSION 5.00
Begin VB.Form frmGlass 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4725
   Icon            =   "frmGlass.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Text            =   "frmGlass.frx":038A
      Top             =   960
      Width           =   4335
   End
   Begin VB.TextBox txtcaption 
      Height          =   285
      Left            =   120
      TabIndex        =   19
      Text            =   "Glass Form"
      Top             =   480
      Width           =   1700
   End
   Begin VB.PictureBox picbotones 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FC00FC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1920
      ScaleHeight     =   255
      ScaleWidth      =   1500
      TabIndex        =   3
      Top             =   0
      Width           =   1500
      Begin VB.PictureBox Picsalf 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FC00FC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         Picture         =   "frmGlass.frx":0390
         ScaleHeight     =   255
         ScaleWidth      =   630
         TabIndex        =   18
         Top             =   1080
         Width           =   630
      End
      Begin VB.PictureBox picsaln 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FC00FC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         Picture         =   "frmGlass.frx":0C52
         ScaleHeight     =   255
         ScaleWidth      =   630
         TabIndex        =   17
         Top             =   1440
         Width           =   630
      End
      Begin VB.PictureBox Picsalfp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FC00FC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         Picture         =   "frmGlass.frx":11F1
         ScaleHeight     =   255
         ScaleWidth      =   630
         TabIndex        =   16
         Top             =   1800
         Width           =   630
      End
      Begin VB.PictureBox Picsal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FC00FC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   870
         Picture         =   "frmGlass.frx":1AB3
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   42
         TabIndex        =   15
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox PICRESMAXP 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FC00FC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         Picture         =   "frmGlass.frx":2052
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   14
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox PICRESMAXF 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FC00FC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         Picture         =   "frmGlass.frx":25AC
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picresmax 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FC00FC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         Picture         =   "frmGlass.frx":2AF9
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   12
         Top             =   840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picresp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FC00FC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         Picture         =   "frmGlass.frx":2FB6
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   11
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox picresf 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FC00FC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         Picture         =   "frmGlass.frx":3513
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   10
         Top             =   1800
         Width           =   495
      End
      Begin VB.PictureBox picresn 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FC00FC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         Picture         =   "frmGlass.frx":3A6C
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   9
         Top             =   2040
         Width           =   495
      End
      Begin VB.PictureBox Picres 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FC00FC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   375
         Picture         =   "frmGlass.frx":3F0A
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   8
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox Picminf 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FC00FC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         Picture         =   "frmGlass.frx":43A8
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   7
         Top             =   960
         Width           =   375
      End
      Begin VB.PictureBox Picminn 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FC00FC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         Picture         =   "frmGlass.frx":48F6
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   6
         Top             =   1320
         Width           =   375
      End
      Begin VB.PictureBox Picminfp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FC00FC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         Picture         =   "frmGlass.frx":4E44
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   5
         Top             =   1680
         Width           =   375
      End
      Begin VB.PictureBox picmin 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FC00FC&
         BorderStyle     =   0  'None
         FillColor       =   &H00FC00FC&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         Picture         =   "frmGlass.frx":5392
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   4
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox Pimage 
      BackColor       =   &H00FC00FC&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   120
      Picture         =   "frmGlass.frx":58E0
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   60
      Width           =   240
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   4695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "etiqueta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   0
      Top             =   50
      Width           =   855
   End
End
Attribute VB_Name = "frmGlass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Activate()
If frmGlass.WindowState <> 1 Then
frmGlass.Caption = ""
frmBlank.Visible = True
End If
frmGlass.SetFocus
End Sub

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()
Label1.Caption = txtcaption.Text
glass
If Me.Tag = "maximizado" Then
picmin.Picture = Picminn.Picture
Picsal.Picture = picsaln.Picture
Picres.Picture = picresmax.Picture
Else
picmin.Picture = Picminn.Picture
Picsal.Picture = picsaln.Picture
Picres.Picture = picresn.Picture
End If
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call frmBlank.pcaption_MouseDown(Button, Shift, x, y)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call frmBlank.pcaption_MouseMove(Button, Shift, x, y)
End Sub

Private Sub picmin_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
picmin.Picture = Picminfp.Picture
End Sub

Private Sub picmin_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Me.Tag = "maximizado" Then
picmin.Picture = Picminf.Picture
Picsal.Picture = picsaln.Picture
Picres.Picture = picresmax.Picture
Else
picmin.Picture = Picminf.Picture
Picsal.Picture = picsaln.Picture
Picres.Picture = picresn.Picture
End If
End Sub
Private Sub picmin_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
picmin.Picture = Picminn.Picture
frmBlank.Visible = False
frmGlass.WindowState = 1
frmGlass.Caption = Me.txtcaption
End Sub

Private Sub Picres_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Picres.Picture = picresp.Picture
End Sub

Private Sub Picres_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Me.Tag = "maximizado" Then
picmin.Picture = Picminn.Picture
Picsal.Picture = picsaln.Picture
Picres.Picture = PICRESMAXF.Picture
Else
picmin.Picture = Picminn.Picture
Picsal.Picture = picsaln.Picture
Picres.Picture = picresf.Picture
End If
End Sub

Private Sub Picres_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Me.Tag <> "maximizado" Then
frmBlank.WindowState = 2
Me.Tag = "maximizado"
Picres.Picture = picresmax.Picture
Else
Me.Tag = "normal"
Picres.Picture = picresn.Picture
frmBlank.WindowState = 0
End If

End Sub


Private Sub Picsal_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Picsal.Picture = Picsalfp.Picture

End Sub

Private Sub Picsal_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Me.Tag = "maximizado" Then
Picsal.Picture = Picsalf.Picture
picmin.Picture = Picminn.Picture
Picres.Picture = picresmax.Picture
Else
Picsal.Picture = Picsalf.Picture
picmin.Picture = Picminn.Picture
Picres.Picture = picresn.Picture
End If

End Sub

Private Sub Picsal_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Picsal.Picture = picsaln.Picture
End
End Sub
Sub glass()
SetWindowLong Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
Me.BackColor = &HFC00FC
Pimage.BackColor = Me.BackColor
SetTrans Me, , Me.BackColor
SetTrans frmBlank, 100
Me.Show
frmBlank.Show
Me.Move frmBlank.Left + 150, frmBlank.Top, frmBlank.Width - 200, frmBlank.Height - 100
Me.SetFocus
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.SetFocus
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.SetFocus
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.SetFocus
End Sub
Private Sub Form_Resize()
If frmGlass.WindowState <> 1 Then
frmBlank.Visible = True
Me.Move frmBlank.Left + 150, frmBlank.Top, frmBlank.Width - 200, frmBlank.Height - 100
picbotones.Left = frmBlank.Width - picbotones.Width - 250
Me.SetFocus
Else
frmBlank.Visible = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmBlank
Set frmBlank = Nothing
End Sub

Private Sub Pimage_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call frmBlank.pcaption_MouseDown(Button, Shift, x, y)
End Sub

Private Sub Pimage_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call frmBlank.pcaption_MouseMove(Button, Shift, x, y)
End Sub

Private Sub txtcaption_Change()
Label1.Caption = txtcaption.Text
End Sub
