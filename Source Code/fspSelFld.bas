Attribute VB_Name = "fspSelFld"
'====================================================================================
'ADDITIONAL API
'====================================================================================
'
'DEVELOPER(S)
'Sadequl Islam Mithun 756
'<sim756@gmail.com>
'<http://sim756.googlepages.com
'
'<NAME>
'<EMAIL>
'<WEB>
'
'====================================================================================

Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const MAX_PATH = 260

Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Public Type BrowseInfo
    hWndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Public Function SelFolder(ahWnd As Long, ahTitle, mrg As Boolean) As String
    On Error Resume Next
    tmp = fspForm1.Text4.Text
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo

    szTitle = ahTitle
    With tBrowseInfo
        .hWndOwner = ahWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        If mrg = 1 Then
            fspForm1.Text4.Text = sBuffer
        End If
        If mrg = 0 Then
            fspForm1.Text2.Text = sBuffer
        End If
    Else
        If mrg = 1 Then
            fspForm1.Text4.Text = tmp
        End If
        If mrg = 0 Then
            fspForm1.Text2.Text = tmp
        End If
    End If
End Function
