VERSION 5.00
Begin VB.Form FrmAnimation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Animation"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2295
   Icon            =   "Animation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   2295
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox AnimationChoice 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Text            =   "<Void>"
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox EndFrame 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Text            =   "0"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox StartFrame 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Text            =   "0"
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Link 
      Caption         =   "Link Frame"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Delete 
      Caption         =   "Delete Frame"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Add 
      Caption         =   "Add Frame"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.HScrollBar Step 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label LLink 
      Caption         =   "End"
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   9
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label LLink 
      Caption         =   "Start"
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   8
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label LStep 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Step"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "FrmAnimation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' **********************
' Add an animation frame
' **********************
'
Private Sub Add_Click()
    Dim i%, j%
    NbAnimation% = NbAnimation% + 1
    ReDim Preserve Animation&(NbWall%, 11, NbAnimation%)
    Call Wall3D.ReDraw(1, True)
    Step.MAX = NbAnimation%
    Step.Value = NbAnimation%
    NoAnimation% = NbAnimation%
    Call Wall3D.ReDraw(1, True)
    For i% = 0 To NbWall%
        For j% = 0 To 11
            Animation&(i%, j%, NbAnimation%) = Animation&(i%, j%, 0)
        Next j%
    Next i%
    Call Wall3D.No_Mur_Change
    Call Wall3D.ReDraw(0, False)
    Call Wall3D.ReDraw(1, False)
End Sub

'
' *************************
' Delete an animation frame
' *************************
'
Private Sub Delete_Click()
    Dim i%, j%, k%
    If NoAnimation% <> NbAnimation% Then
        For i% = 0 To NbWall%
            For j% = 0 To 11
                For k% = NoAnimation% To NbAnimation% - 1
                    Animation&(i%, j%, k%) = Animation&(i%, j%, k% + 1)
                Next k%
            Next j%
        Next i%
    End If
    NbAnimation% = NbAnimation% - 1
    ReDim Preserve Animation&(NbWall%, 11, NbAnimation%)
    Step.MAX = NbAnimation%
    Step.Value = 0
    NoAnimation% = 0
    Call Wall3D.No_Mur_Change
    Call Wall3D.ReDraw(0, False)
    Call Wall3D.ReDraw(1, False)
End Sub

'
' *********************
' Initialize link value
' *********************
'
Private Sub Form_Load()
    With AnimationChoice
        Call .AddItem("Delta X")
        Call .AddItem("Delta Y")
        Call .AddItem("Delta Z")
        Call .AddItem("Size X")
        Call .AddItem("Size Y")
        Call .AddItem("Size Z")
        Call .AddItem("Theta")
        Call .AddItem("Phi")
        Call .AddItem("Roll")
        Call .AddItem("Axe Theta")
        Call .AddItem("Axe Phi")
        Call .AddItem("Axe Roll")
        .ListIndex = 0
    End With
    Step.MIN = 0
    Step.MAX = NbAnimation%
    Step.Value = 0
    Call Step_Change
End Sub

'
' **********************
' Close animation window
' **********************
'
Private Sub Form_Unload(Cancel As Integer)
    NoAnimation% = 0
    Call Wall3D.No_Mur_Change
    Call Wall3D.ReDraw(0, False)
    Call Wall3D.ReDraw(1, False)
End Sub

'
' **********************
' Change animation frame
' **********************
'
Private Sub Step_Change()
    NoAnimation% = Step.Value
    LStep = NoAnimation%
    If Step.Value = 0 Then
        Delete.Enabled = False
        Link.Enabled = False
    Else
        Delete.Enabled = True
        Link.Enabled = True
    End If
    Call Wall3D.No_Mur_Change
    Call Wall3D.ReDraw(0, False)
    Call Wall3D.ReDraw(1, False)
End Sub

