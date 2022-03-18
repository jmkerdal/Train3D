VERSION 5.00
Begin VB.Form Axe_Camera 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Camera"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox ShowAxes 
      Caption         =   "Show axes"
      Height          =   195
      Left            =   0
      TabIndex        =   37
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Frame Angles 
      Caption         =   "Look direction"
      Height          =   1335
      Left            =   0
      TabIndex        =   16
      Top             =   1080
      Width           =   2535
      Begin VB.CommandButton PlusVite 
         Caption         =   ">>"
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   35
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton Plus 
         Caption         =   ">"
         Height          =   255
         Index           =   6
         Left            =   1800
         TabIndex        =   34
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton Moins 
         Caption         =   "<"
         Height          =   255
         Index           =   6
         Left            =   960
         TabIndex        =   33
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton MoinsVite 
         Caption         =   "<<"
         Height          =   255
         Index           =   6
         Left            =   600
         TabIndex        =   32
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton PlusVite 
         Caption         =   ">>"
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   28
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton PlusVite 
         Caption         =   ">>"
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   27
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Plus 
         Caption         =   ">"
         Height          =   255
         Index           =   5
         Left            =   1800
         TabIndex        =   26
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton Plus 
         Caption         =   ">"
         Height          =   255
         Index           =   4
         Left            =   1800
         TabIndex        =   25
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton PlusVite 
         Caption         =   ">>"
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   24
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Plus 
         Caption         =   ">"
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   23
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Moins 
         Caption         =   "<"
         Height          =   255
         Index           =   5
         Left            =   960
         TabIndex        =   22
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton Moins 
         Caption         =   "<"
         Height          =   255
         Index           =   4
         Left            =   960
         TabIndex        =   21
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton MoinsVite 
         Caption         =   "<<"
         Height          =   255
         Index           =   5
         Left            =   600
         TabIndex        =   20
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton MoinsVite 
         Caption         =   "<<"
         Height          =   255
         Index           =   4
         Left            =   600
         TabIndex        =   19
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Moins 
         Caption         =   "<"
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   18
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton MoinsVite 
         Caption         =   "<<"
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   17
         Top             =   240
         Width           =   375
      End
      Begin VB.Label AxeCamera 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Index           =   6
         Left            =   1200
         TabIndex        =   44
         Top             =   960
         Width           =   615
      End
      Begin VB.Label AxeCamera 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   43
         Top             =   720
         Width           =   615
      End
      Begin VB.Label AxeCamera 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   42
         Top             =   480
         Width           =   615
      End
      Begin VB.Label AxeCamera 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   41
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Falloff"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   36
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Phi"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Theta"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   30
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Range"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Axes 
      Caption         =   "Hock position"
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.CommandButton PlusVite 
         Caption         =   ">>"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   12
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Plus 
         Caption         =   ">"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   11
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton PlusVite 
         Caption         =   ">>"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   10
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Plus 
         Caption         =   ">"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   9
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton PlusVite 
         Caption         =   ">>"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Plus 
         Caption         =   ">"
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Moins 
         Caption         =   "<"
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   6
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton MoinsVite 
         Caption         =   "<<"
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   5
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Moins 
         Caption         =   "<"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   4
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton MoinsVite 
         Caption         =   "<<"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   3
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Moins 
         Caption         =   "<"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton MoinsVite 
         Caption         =   "<<"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
      Begin VB.Label AxeCamera 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   40
         Top             =   720
         Width           =   615
      End
      Begin VB.Label AxeCamera 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   39
         Top             =   480
         Width           =   615
      End
      Begin VB.Label AxeCamera 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   38
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Z"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Y"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "X"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "Axe_Camera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' ****************************
' Input a new value for camera
' ****************************
'
Private Sub AxeCamera_Click(Index As Integer)
    Dim v%
    If Index = 3 Then
        v% = InputBox("New value", "", TheCamera.range / 10)
        If v% < 0 Then v% = 1
        TheCamera.range = v% * 10
        AxeCamera(Index).Caption = TheCamera.range / 10
        Call Position
    End If
End Sub

'
' ************************
' Substract step to camera
' ************************
'
Private Sub Moins_Click(Index As Integer)
    Call Traveling(Index%, -1)
End Sub

'
' ***************************
' Substract 10 step to camera
' ***************************
'
Private Sub MoinsVite_Click(Index As Integer)
    Call Traveling(Index%, -10)
End Sub

'
' **********************
' Add one step to camera
' **********************
'
Private Sub Plus_Click(Index As Integer)
    Call Traveling(Index%, 1)
End Sub

'
' *********************************
' Calculate new position for camera
' *********************************
'
Public Sub Position()
    DoEvents ' Force update value in label
    Call dxEngine.dxCamera.SetPosition(dxEngine.dxScene, TheCamera.Position.X, TheCamera.Position.Y, TheCamera.Position.z)
    Call dxEngine.dxCamera.SetOrientation(dxEngine.dxScene, 0, 0, 1, 0, 1, 0)
    Call dxEngine.dxCamera.AddRotation(D3DRMCOMBINE_BEFORE, 0, 1, 0, TheCamera.theta / 180 * PI!)
    Call dxEngine.dxCamera.AddRotation(D3DRMCOMBINE_BEFORE, 0, 0, 1, TheCamera.phi / 180 * PI!)
    Call dxEngine.dxCamera.AddRotation(D3DRMCOMBINE_BEFORE, 1, 0, 0, TheCamera.falloff / 180 * PI!)
    Call dxEngine.dxCamera.AddTranslation(D3DRMCOMBINE_BEFORE, 0, 0, -TheCamera.range / 10)
    If Orthographic = True Then
        Call dxEngine.dxViewport.SetField(TheCamera.range / 20)
    End If
    Call Wall3D.ReDraw(0, False)
    Call Wall3D.ReDraw(1, False)
End Sub

'
' *********************
' Add 10 step to camera
' *********************
'
Private Sub PlusVite_Click(Index As Integer)
    Call Traveling(Index%, 10)
End Sub

'
' ****************************
' Default value for the camera
' ****************************
'
Public Sub Init()
    With TheCamera
        .Position.X = 0
        .Position.Y = 4
        .Position.z = 0
        .range = 200
        .theta = 0
        .phi = 0
        .falloff = 0
        AxeCamera(0).Caption = .Position.X / 10
        AxeCamera(1).Caption = .Position.Y / 10
        AxeCamera(2).Caption = .Position.z / 10
        AxeCamera(3).Caption = .range / 10
        AxeCamera(4).Caption = .theta
        AxeCamera(5).Caption = .phi
        AxeCamera(6).Caption = .falloff
    End With
End Sub

'
' ***********************
' Show axes to help user
' to set position of mesh
' ***********************
'
Private Sub ShowAxes_Click()
    Call Wall3D.ReDraw(0, False)
    Call Wall3D.ReDraw(1, False)
End Sub

'
' ***************
' Move the camera
' ***************
'
Public Sub Traveling(Index%, ByVal Delta%)
    Select Case Index%
    Case 0
        TheCamera.Position.X = TheCamera.Position.X + Delta%
        AxeCamera(0).Caption = TheCamera.Position.X / 10
    Case 1
        TheCamera.Position.Y = TheCamera.Position.Y + Delta%
        AxeCamera(1).Caption = TheCamera.Position.Y / 10
    Case 2
        TheCamera.Position.z = TheCamera.Position.z + Delta%
        AxeCamera(2).Caption = TheCamera.Position.z / 10
    Case 3
        TheCamera.range = TheCamera.range + Delta%
        AxeCamera(3).Caption = TheCamera.range / 10
    Case 4
        TheCamera.theta = TheCamera.theta + Delta%
        If TheCamera.theta >= 360 Then
            TheCamera.theta = TheCamera.theta - 360
        End If
        If TheCamera.theta < 0 Then
            TheCamera.theta = TheCamera.theta + 360
        End If
        AxeCamera(4).Caption = TheCamera.theta
    Case 5
        TheCamera.phi = TheCamera.phi + Delta%
        If TheCamera.phi >= 360 Then
            TheCamera.phi = TheCamera.phi - 360
        End If
        If TheCamera.phi < 0 Then
            TheCamera.phi = TheCamera.phi + 360
        End If
        AxeCamera(5).Caption = TheCamera.phi
    Case 6
        TheCamera.falloff = TheCamera.falloff + Delta%
        If TheCamera.falloff >= 360 Then
            TheCamera.falloff = TheCamera.falloff - 360
        End If
        If TheCamera.falloff < 0 Then
            TheCamera.falloff = TheCamera.falloff + 360
        End If
        AxeCamera(6).Caption = TheCamera.falloff
    End Select
    Call Position
End Sub

'
' ***************************
' Force orientation on 3 axes
' ***************************
'
Public Sub Orientation(Dtheta%, Dphi%, Dfalloff%)
    TheCamera.theta = TheCamera.theta + Dtheta%
    If TheCamera.theta >= 360 Then
        TheCamera.theta = TheCamera.theta - 360
    End If
    If TheCamera.theta < 0 Then
        TheCamera.theta = TheCamera.theta + 360
    End If
    AxeCamera(4).Caption = TheCamera.theta
    '
    TheCamera.phi = TheCamera.phi + Dphi%
    If TheCamera.phi >= 360 Then
        TheCamera.phi = TheCamera.phi - 360
    End If
    If TheCamera.phi < 0 Then
        TheCamera.phi = TheCamera.phi + 360
    End If
    AxeCamera(5).Caption = TheCamera.phi
    TheCamera.falloff = TheCamera.falloff + Dfalloff%
    If TheCamera.falloff >= 360 Then
        TheCamera.falloff = TheCamera.falloff - 360
    End If
    If TheCamera.falloff < 0 Then
        TheCamera.falloff = TheCamera.falloff + 360
    End If
    AxeCamera(6).Caption = TheCamera.falloff
    Call Position
End Sub

