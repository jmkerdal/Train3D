Attribute VB_Name = "Tools"
'
' ******************************
' Tool Box
' © 1998-2000 Kerdal Jean-Michel
' ******************************
'
Option Explicit
'
' ***** Box constants
'
Enum EnumBox
    BOX_LOAD
    BOX_SAVE
    BOX_COLOR
End Enum
Enum EnumOpen
    OPEN_NORMAL ' Open the file in text
    OPEN_BINARY ' Open the file in binary
End Enum

'
' *********************
' Add a entry in a list
' *********************
'
Sub Add_Entry(ListName As Control, FileName$, Position%)
    If Position% + 1 > ListName.ListCount Then
        ListName.AddItem "[" + Format$(Position%) + "] " + FileName$
    Else
        ListName.RemoveItem Position%
        ListName.AddItem "[" + Format$(Position%) + "] " + FileName$, Position%
    End If
End Sub

'
' *********************************
' Search if this file already exist
' *********************************
'
Function Exist(File$) As Boolean
    Dim f%
    On Error GoTo Failed
    f% = FreeFile()
    Open File$ For Input As #f%
    Close #f%
    Exist = True
    On Error GoTo 0
    Exit Function
Failed:
    Close #f%
    Exist = False
    On Error GoTo 0
End Function

'
' ******************************
' Open a file with test and mode
' ******************************
'
Sub Open_File(Nom$, File%, Mode As EnumOpen)
    Dim r%
    Do While Exist(Nom$) = False
        If MsgBox("This file " + Nom$ + vbCr + "doesn't exist...", vbQuestion + vbRetryCancel) = vbCancel Then
            End
        End If
    Loop
    If Mode = OPEN_BINARY Then
        Open Nom$ For Binary As #File% Len = 4096
    Else
        Open Nom$ For Input As #File% Len = 512
    End If
End Sub

'
' **************************
' Open the selector file box
' and reduce path to relatif
' **************************
'
Function Open_Box$(Title$, Name$, BoxFilter$, OpenBox As EnumBox, Boite As CommonDialog, Optional Folder$)
    Dim d$ ' Save the CurDir because the selector modify this value
    Dim BoxFile$
    d$ = CurDir
    With Boite
        .DialogTitle = Title$
        .Filter = BoxFilter$
        .FileName = Name$
        Select Case OpenBox
        Case BOX_LOAD
            .ShowOpen
        Case BOX_SAVE
            .ShowSave
        Case BOX_COLOR
            .ShowColor
        End Select
        Folder$ = CurDir$
        ChDir d$
        BoxFile$ = .FileName
    End With
    If Left$(BoxFile$, Len(d$)) = d$ Then
        Open_Box$ = "." + Right$(BoxFile$, Len(BoxFile$) - Len(d$))
    Else
        Open_Box$ = BoxFile$
    End If
End Function

