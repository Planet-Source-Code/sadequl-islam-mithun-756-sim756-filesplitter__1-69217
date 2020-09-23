Attribute VB_Name = "fspFunction"
'====================================================================================
'THE FILE MERGING MODULE
'====================================================================================
'
'This program/module is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'For the first time this software is written by,
'Sadequl Islam Mithun 756
'<sim756@gmail.com>
'<http://sim756.googlepages.com
'
'DEVELOPER(S)
'<NAME>
'<EMAIL>
'<WEB>
'
'====================================================================================

Public Function fspmk(fspDFolder As String)
    On Error GoTo eor:
    Dim dta As Byte
    Dim mk As String, k_ As Long
    Dim FileName As String, fspTotal
    akFileName = App.Path & "\"
    FileName = App.EXEName & ".exe"
    Open FileName For Input As #3
        Seek #3, 28672 '40960
        Line Input #3, fspTotal
    Close #3
    fspForm1.Picture3.ScaleWidth = Val(fspTotal)
    FileName = App.EXEName
    Open fspDFolder & "\" & FileName For Binary As #1
        Do Until Val(xz) >= Val(fspTotal)
            DoEvents
            xz = xz + 1
            Open akFileName & FileName & xz & ".spt" For Binary As #2
                For k = 1 To LOF(2)
                    Get #2, , dta
                    Put #1, , dta
                Next k
            Close #2
            fspForm1.c7.Width = Val(xz)
        Loop
    Close #1
    Exit Function
eor:
    If Err Then MsgBox "Error!" & Chr(10) & "Error Number : " & Err.Number & Chr(10) & "Error Description : " & Err.Description, vbCritical + vbOKOnly, "Error"
End Function
