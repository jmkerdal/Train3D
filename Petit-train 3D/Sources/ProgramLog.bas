Attribute VB_Name = "ProgramLog"
Option Explicit
'
Dim LogPath$
'
Enum EnumLogMode
    LogMode_Append
    LogMode_Overwrite
End Enum

'
' ****************************
' Select Path for the log file
' ****************************
'
Public Sub Select_Path(Name$, Mode As EnumLogMode)
    LogPath$ = Name$
    If Mode = LogMode_Append Then Exit Sub
    Dim f%
    f% = FreeFile()
    Open LogPath$ For Output As #f%
    Close #f%
End Sub

'
' ************************************************
' Write date, number and a comment to the log file
' ************************************************
'
Public Sub Write_File(n%, a$)
    Dim f%
    If LogPath$ = "" Then Exit Sub ' Not enabled
    f% = FreeFile()
    Open LogPath$ For Append As #f%
    Print #f%, Now; n%; a$
    Close #f%
End Sub

