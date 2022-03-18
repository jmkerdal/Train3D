VERSION 5.00
Begin VB.Form Transform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Position"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   1455
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option 
      Height          =   375
      Index           =   3
      Left            =   1080
      MaskColor       =   &H0000FFFF&
      Picture         =   "Transform.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Axe rotate"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.HScrollBar Step 
      Height          =   255
      Left            =   0
      Max             =   2
      Min             =   -5
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.OptionButton Option 
      Height          =   375
      Index           =   2
      Left            =   720
      MaskColor       =   &H0000FFFF&
      Picture         =   "Transform.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Size"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option 
      Height          =   375
      Index           =   1
      Left            =   360
      MaskColor       =   &H0000FFFF&
      Picture         =   "Transform.frx":0684
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Rotate"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option 
      Height          =   375
      Index           =   0
      Left            =   0
      MaskColor       =   &H0000FFFF&
      Picture         =   "Transform.frx":09C6
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Move"
      Top             =   0
      UseMaskColor    =   -1  'True
      Value           =   -1  'True
      Width           =   375
   End
   Begin VB.Label LStep 
      Caption         =   "Step: 1"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "Transform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' **************************************
' Select transform to apply for the mesh
' **************************************
'
Option Explicit
Public StepValue!

'
' ***********************
' Init value for one step
' ***********************
'
Private Sub Form_Load()
    StepValue! = 10 ^ 5
End Sub

'
' *********************
' No focus on this form
' *********************
'
Private Sub LStep_Click()
    Call View.SetFocus
End Sub

'
' *********************
' No focus on this form
' *********************
'
Private Sub Option_Click(Index As Integer)
    Call View.SetFocus
End Sub

'
' *********************
' Change the step value
' *********************
'
Private Sub Step_Change()
    StepValue! = 10 ^ (Step.Value + 5)
    LStep.Caption = "Step: " & Format(StepValue! / 10 ^ 5)
    Call View.SetFocus
End Sub

