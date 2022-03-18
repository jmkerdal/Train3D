VERSION 5.00
Begin VB.Form Textures 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Textures"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   Icon            =   "Textures.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   256
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   464
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox MiniTextures 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3840
      Left            =   3120
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   1
      Top             =   0
      Width           =   3840
   End
   Begin VB.ListBox Liste_Texture 
      Height          =   3780
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   2520
      Top             =   1800
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "Textures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    Call List_Update
End Sub

'
' ***********************
' Change the texture path
' ***********************
'
Private Sub Liste_Texture_DblClick()
    If Liste_Texture.ListIndex <> 0 Then
        Call Texture_Edit.Show(vbModal)
    End If
End Sub

'
' ****************
' Delete a texture
' ****************
'
Private Sub Liste_Texture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim n%
    n% = Liste_Texture.ListIndex
    If n% > 0 Then
        If Texture(n%).Name$ <> "" Then
            If Button = vbKeyRButton Then
                If MsgBox("Remove this texture", vbQuestion + vbQuestion + vbYesNo) = vbYes Then
                    Texture(n%).Name$ = ""
                    Set dxTexture(n%) = Nothing
                    Call Add_Entry(Textures.Liste_Texture, Texture(n%).Name$, n%)
                    Call Wall3D.Add_Texture(Texture(n%).Name$, n%)
                End If
            End If
        End If
    End If
    If Button = vbKeyLButton Then
        If Texture(n%).Name$ <> "" Then
            Liste_Texture.ToolTipText = Texture(n%).Name$
        Else
            Liste_Texture.ToolTipText = ""
        End If
    End If
End Sub

'
' *******************************
' Select a texture with the mouse
' *******************************
'
Private Sub MiniTextures_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Textures.Liste_Texture.ListIndex = 1 + (X \ 16) + (Y \ 16) * 16
    End If
End Sub

'
' *******************************************
' Texture list update after loading from disk
' from any file holding texture definition
' *******************************************
'
Public Sub List_Update()
    Dim i%
    Textures.Liste_Texture.Enabled = False
    Call Add_Entry(Textures.Liste_Texture, "Textures list", 0)
    Call Wall3D.Add_Texture("<None>", 0)
    For i% = 1 To NbTexture%
        Call Add_Entry(Textures.Liste_Texture, Texture(i%).Name$, i%)
        Call Wall3D.Add_Texture(Texture(i%).Name$, i%)
    Next i%
    Textures.Liste_Texture.Enabled = True
End Sub

