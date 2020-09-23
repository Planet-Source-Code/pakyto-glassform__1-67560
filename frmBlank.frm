VERSION 5.00
Begin VB.Form frmBlank 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F8F8F8&
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   6285
   ControlBox      =   0   'False
   DrawMode        =   6  'Mask Pen Not
   DrawWidth       =   2
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmBlank.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   4365
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcaption 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F8F8F8&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   6285
      TabIndex        =   0
      Top             =   0
      Width           =   6285
   End
End
Attribute VB_Name = "frmBlank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CurrX, CurrY As Single
Private Sub Form_Activate()
If frmGlass.WindowState <> 1 Then
frmGlass.Caption = ""
End If
frmGlass.SetFocus
End Sub

Private Sub Form_Load()
RedondearEsquinaForm Me, 15
ProcOld = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 frmGlass.SetFocus
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If frmGlass.Tag = "maximizado" Then
frmGlass.picmin.Picture = frmGlass.Picminn.Picture
frmGlass.Picsal.Picture = frmGlass.picsaln.Picture
frmGlass.Picres.Picture = frmGlass.picresmax.Picture
Else
frmGlass.picmin.Picture = frmGlass.Picminn.Picture
frmGlass.Picsal.Picture = frmGlass.picsaln.Picture
frmGlass.Picres.Picture = frmGlass.picresn.Picture
End If
 frmGlass.SetFocus
End Sub
Private Sub Form_Unload(Cancel As Integer)
Call SetWindowLong(Me.hWnd, GWL_WNDPROC, ProcOld)
End Sub

Public Sub pcaption_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
CurrX = x
CurrY = y
frmGlass.SetFocus
End Sub
Public Sub pcaption_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If frmGlass.Tag = "maximizado" Then
frmGlass.picmin.Picture = frmGlass.Picminn.Picture
frmGlass.Picsal.Picture = frmGlass.picsaln.Picture
frmGlass.Picres.Picture = frmGlass.picresmax.Picture
Else
frmGlass.picmin.Picture = frmGlass.Picminn.Picture
frmGlass.Picsal.Picture = frmGlass.picsaln.Picture
frmGlass.Picres.Picture = frmGlass.picresn.Picture
End If
Call DragWindow(Button, x, y)
frmGlass.SetFocus
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 frmGlass.SetFocus
End Sub
Public Function DragWindow(Button As Integer, x As Single, y As Single)
If Me.WindowState = 2 Then Exit Function
If Button = 1 Then
Me.Left = Me.Left + (x - CurrX)
Me.Top = Me.Top + (y - CurrY)
frmGlass.Move frmBlank.Left + 150, frmBlank.Top, frmBlank.Width - 200, frmBlank.Height - 100
End If
frmGlass.SetFocus
End Function
Private Sub Form_Resize()
If frmBlank.WindowState = 1 Then Exit Sub
RedondearEsquinaForm Me, 15
frmGlass.Move frmBlank.Left + 150, frmBlank.Top, frmBlank.Width - 200, frmBlank.Height - 100
frmGlass.Label2.Width = frmGlass.Width - 100: frmGlass.Label2.Height = frmGlass.Height - 400
frmGlass.Text1.Width = frmGlass.Label2.Width - 250
frmGlass.Text1.Height = frmGlass.Label2.Height - 800
frmGlass.SetFocus
End Sub

