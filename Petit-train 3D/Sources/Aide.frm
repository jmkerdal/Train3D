VERSION 5.00
Begin VB.Form Aide 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Help"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6270
   Icon            =   "Aide.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "Aide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' ************************
' Ecrit un nouveau message
' ************************
'
Public Sub Ecrit(Message$)
    Me.Show
    Call List1.AddItem(Message$)
End Sub

'
' **********************************
' Vide la liste lors de la fermeture
' **********************************
'
Private Sub Form_Unload(Cancel As Integer)
    List1.Clear
End Sub

