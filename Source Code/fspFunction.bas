Attribute VB_Name = "fspFunction"
'====================================================================================
'THE FILE SPLITTING & MERGING MODULE
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
'Developer(s)
'<NAME>
'<EMAIL>
'<WEB>
'
'====================================================================================

Public Function fsp(fspSFile As String, fspDFolder As String, fspSPFSize As Long)
    Dim fspFSize As Long, strall As String
    Dim dt As Byte
    Dim X, Y As Currency
    For k = -Len(fspSFile) To 1
        k_ = -k
        mk = Mid(fspSFile, k_, 1)
        If mk = "\" Then
            FileName = Right(fspSFile, Len(fspSFile) - k_)
            Exit For
        End If
    Next k
    fspFSize = FileLen(fspSFile)
    fspForm1.Picture4.ScaleWidth = fspFSize
    tmpS = fspFSize
    
    Open fspSFile For Binary As #1
    Do Until X = fspFSize
        DoEvents
        m = m + 1
        Y = Y + 1
        Open fspDFolder & "\" & FileName & m & ".spt" For Binary As #2
            
            If Not tmpS < fspSPFSize Then
                For k = 1 To fspSPFSize
                    Get #1, , dt
                    Put #2, , dt
                    X = X + 1
                Next k
                DoEvents
                strall = strall & FileName & m & ".spt" & "+"
                fspForm1.c6.Width = X
            Else
                For k = 1 To tmpS
                    Get #1, , dt
                    Put #2, , dt
                    X = X + 1
                Next k
                DoEvents
                strall = strall & FileName & m & ".spt"
                fspForm1.c6.Width = X
            End If
        Close #2
        tmpS = tmpS - fspSPFSize
        xk = xk + 1
    Loop
    Close #1
    If fspForm1.Option5.Value = True And fspForm1.Option5.Enabled = True Then
        afspTotal = Trim(xk)
        DoEvents
        Dim akByte As Byte
        Open App.Path & "\fspEx.exe" For Binary As #10
            Open fspDFolder & "\" & FileName & ".exe" For Binary As #11
                For fs = 1 To 28672
                    Get #10, , akByte
                    Put #11, , akByte
                Next fs
            Close #11
        Close #10
        Open fspDFolder & "\" & FileName & ".exe" For Append As #12
            Print #12, afspTotal
        Close #12
        MsgBox "File split has successfuly done!"
        Exit Function
    End If
    If fspForm1.Option4.Value = True And fspForm1.Option4.Enabled = True Then
        Open fspDFolder & "\merge.temp" For Binary As #4
        Close #4
        Open fspDFolder & "\mfs.merge.bat" For Append As #5
            Print #5, "echo off"
            Print #5, "cls"
            Print #5, "echo."
            Print #5, "echo                            FileMerger " & App.Major & "." & App.Minor & "." & App.Revision
            Print #5, "echo."
            Print #5, "echo Press Enter to merge files in current folder as " & FileName & " or press ctrl+c to close"
            Print #5, "pause"
            Print #5, "echo."
            Print #5, "echo File merging..."
            Print #5, "echo."
            Print #5, "COPY /V /-Y /B " & "merge.temp" & "+" & strall & " /B " & FileName
            Print #5, "echo."
            Print #5, "echo Completed"
            Print #5, "echo."
        Close #5
        MsgBox "File splitting is successful"
        Exit Function
    End If
    Open fspDFolder & "\" & FileName & ".spi" For Append As #3
        Print #3, FileName
        Print #3, Trim(xk)
    Close #3
    MsgBox "File splitting is successful"
    Exit Function
eor:
    If Err Then MsgBox "Error!" + Chr(10) + "Error Number : " & Err.Number + Chr(10) & "Error Description : " & Err.Description, vbCritical + vbOKOnly, "Error"
End Function

Public Function fspmk(fspSFolder As String, fspDFolder As String)
    On Error GoTo eor:
    Dim dta As Byte
    Dim mk As String, k_ As Long
    Dim FileName As String, fspTotal
    For k = -Len(fspSFolder) - 1 To 1
        k_ = -k
        mk = Mid(fspSFolder, k_, 1)
        If mk = "\" Then
            xFileName = Left(fspSFolder, Len((Right(fspSFolder, Len(fspSFolder) - k_))))
            Exit For
        End If
    Next k
    xfspSFile = fspSFolder
    For k = 1 To Len(xfspSFile)
        mk = Mid(fspSFile, k, 1)
        If mk = "\" Then
            akFileName = Left(xfspSFile, Len(Right(xfspSFile, Len(xfspSFile) - k)))
            Exit For
        End If
    Next k
    Open fspSFolder For Input As #3
        Line Input #3, FileName
        Line Input #3, fspTotal
    Close #3
        fspForm1.Picture3.ScaleWidth = Val(fspTotal)
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
    If Err Then MsgBox "Error!" + Chr(10) + "Error Number : " & Err.Number + Chr(10) & "Error Description : " & Err.Description, vbCritical + vbOKOnly, "Error"
End Function
