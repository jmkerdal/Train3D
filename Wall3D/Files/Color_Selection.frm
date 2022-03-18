VERSION 5.00
Object = "{EC2CC72E-13BA-11D5-BB31-400001686160}#1.0#0"; "SELECTCOLOR.OCX"
Begin VB.Form ColorSelection 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Color"
   ClientHeight    =   735
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   2295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SelectColor.UserColor TheColor 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
   End
End
Attribute VB_Name = "ColorSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' *****************************
' Quit the form and save values
' *****************************
'
Private Sub Form_Unload(Cancel As Integer)
    TheRed = TheColor.Red
    TheGreen = TheColor.Green
    TheBlue = TheColor.Blue
    TheAlpha = TheColor.Alpha
End Sub

