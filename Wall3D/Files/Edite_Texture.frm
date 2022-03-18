VERSION 5.00
Object = "{EC2CC72E-13BA-11D5-BB31-400001686160}#1.0#0"; "SELECTCOLOR.OCX"
Begin VB.Form Texture_Edit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Texture name"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2535
   Icon            =   "Edite_Texture.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   2535
   StartUpPosition =   3  'Windows Default
   Begin SelectColor.UserColor UserColor 
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1720
   End
   Begin VB.CheckBox CTransparent 
      Caption         =   "Transparent"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CCharge_Texture 
      Caption         =   "&Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Texture_Edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Dim No% ' Texture number to modify

'
' *******************************
' Load the texture name from file
' *******************************
'
Private Sub CCharge_Texture_Click()
    Dim FileName$
    FileName$ = Open_Box$("Load textures", "", "Picture|*.bmp", BOX_LOAD, Wall3D.BoiteMur3D)
    If FileName$ = "" Then Exit Sub
    Texture(No%).Name$ = FileName$
    Set dxTexture(No%) = dxEngine.Load_Texture(Texture(No%).Name$)
    Call Add_Entry(Textures.Liste_Texture, Texture(No%).Name$, No%)
    Call Wall3D.Add_Texture(Texture(No%).Name$, No%)
    Texture_Edit.Caption = Texture(No%).Name$
End Sub

'
' **************************
' Switch to transparent mode
' **************************
'
Private Sub CTransparent_Click()
    Texture(No%).Transparency = CTransparent.Value
    If dxTexture(No%) Is Nothing Then Exit Sub
    If CTransparent.Value <> 0 Then
        Call dxTexture(No%).SetDecalTransparency(True)
    Else
        Call dxTexture(No%).SetDecalTransparency(False)
    End If
End Sub

'
' *************************************
' Move the form to his initial position
' *************************************
'
Private Sub Form_Activate()
    Call Me.Move(Textures.Left, Textures.Top)
End Sub

'
' ********************
' Load previous values
' ********************
'
Private Sub Form_Load()
    Dim r%, g%, b% ' Color depth
    UserColor.ShowAlpha = False
    No% = Textures.Liste_Texture.ListIndex
    Me.Caption = Texture(No%).Name$
    CTransparent.Value = Texture(No%).Transparency
    r% = dxEngine.dX7.ColorGetRed(Texture(No%).Color) * 255
    g% = dxEngine.dX7.ColorGetGreen(Texture(No%).Color) * 255
    b% = dxEngine.dX7.ColorGetBlue(Texture(No%).Color) * 255
    UserColor.Red.Value = r%
    UserColor.Green.Value = g%
    UserColor.Blue = b%
End Sub

'
' ****************************
' Change the transparent color
' ****************************
'
Private Sub Form_Unload(Cancel As Integer)
    Texture(No%).Color = dxEngine.dX7.CreateColorRGBA(UserColor.Red.Value / 255 _
    , UserColor.Green.Value / 255, UserColor.Blue.Value / 255, 1)
    If dxTexture(No%) Is Nothing Then Exit Sub
    Call dxTexture(No%).SetDecalTransparentColor(Texture(No%).Color)
End Sub

