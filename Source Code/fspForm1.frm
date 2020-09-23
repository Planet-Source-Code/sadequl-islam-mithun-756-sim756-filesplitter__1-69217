VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fspForm1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "fsp"
   ClientHeight    =   5925
   ClientLeft      =   2850
   ClientTop       =   2925
   ClientWidth     =   5715
   Icon            =   "fspForm1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   5715
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox hdr1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   60
      ScaleHeight     =   825
      ScaleWidth      =   5580
      TabIndex        =   38
      Top             =   60
      Width           =   5610
      Begin VB.Label hdr2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<hdr>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   39
         Top             =   150
         Width           =   1140
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2640
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox p3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   -2400
      ScaleHeight     =   4455
      ScaleWidth      =   8115
      TabIndex        =   13
      Top             =   1500
      Visible         =   0   'False
      Width           =   8115
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00F4F4F4&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3915
         Left            =   2460
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   35
         Text            =   "fspForm1.frx":030A
         Top             =   480
         Width           =   5595
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "FileSplitter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         TabIndex        =   37
         Top             =   60
         Width           =   1590
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "<Ver>"
         Height          =   195
         Left            =   4200
         TabIndex        =   36
         Top             =   180
         Width           =   420
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   60
      ScaleHeight     =   435
      ScaleWidth      =   5595
      TabIndex        =   29
      Top             =   1020
      Width           =   5595
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3620
         ScaleHeight     =   615
         ScaleWidth      =   105
         TabIndex        =   34
         Top             =   0
         Width           =   100
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1770
         ScaleHeight     =   615
         ScaleWidth      =   90
         TabIndex        =   33
         Top             =   0
         Width           =   90
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F4F4F4&
         Caption         =   "About"
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   2
         Left            =   3660
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   -60
         Width           =   1995
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F4F4F4&
         Caption         =   "Merge File"
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   1
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   -60
         Width           =   1875
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C9C9C9&
         Caption         =   "Split File"
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   0
         Left            =   -60
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   -60
         Value           =   -1  'True
         Width           =   1875
      End
   End
   Begin VB.PictureBox p2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4395
      Left            =   60
      ScaleHeight     =   4395
      ScaleWidth      =   5595
      TabIndex        =   12
      Top             =   1500
      Visible         =   0   'False
      Width           =   5595
      Begin VB.CommandButton Command6 
         Caption         =   "..."
         Height          =   315
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   960
         Width           =   315
      End
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         Height          =   315
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   315
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   180
         TabIndex        =   22
         Top             =   960
         Width           =   4875
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00EBEBEB&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   180
         ScaleHeight     =   285
         ScaleWidth      =   5145
         TabIndex        =   18
         Top             =   2040
         Width           =   5175
         Begin VB.PictureBox c7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000003&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   0
            ScaleHeight     =   315
            ScaleWidth      =   15
            TabIndex        =   28
            Top             =   0
            Width           =   15
         End
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Merge"
         Height          =   375
         Left            =   3900
         TabIndex        =   17
         Top             =   1500
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   180
         TabIndex        =   15
         Top             =   360
         Width           =   4875
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Merge location"
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Merge progress"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select <filename>.spi (merge info)"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   120
         Width           =   2370
      End
   End
   Begin VB.PictureBox P1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4395
      Left            =   60
      ScaleHeight     =   4395
      ScaleWidth      =   5595
      TabIndex        =   0
      Top             =   1500
      Width           =   5595
      Begin VB.OptionButton Option4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Make a dos batch script to merge file (312 Bytes)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   300
         TabIndex        =   26
         Top             =   2700
         Width           =   4200
      End
      Begin VB.OptionButton Option5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Make a PE (exe) module to merge file ( .......Bytes)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   300
         TabIndex        =   25
         Top             =   2400
         Value           =   -1  'True
         Width           =   4700
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00EBEBEB&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   180
         ScaleHeight     =   285
         ScaleWidth      =   5145
         TabIndex        =   20
         Top             =   3840
         Width           =   5175
         Begin VB.PictureBox c6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000003&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   0
            ScaleHeight     =   315
            ScaleWidth      =   15
            TabIndex        =   27
            Top             =   0
            Width           =   15
         End
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "fspForm1.frx":50D7
         Left            =   180
         List            =   "fspForm1.frx":50F6
         TabIndex        =   10
         Text            =   "1.44 MB"
         Top             =   1620
         Width           =   1935
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   180
         TabIndex        =   9
         Top             =   2100
         Width           =   5175
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Make merging module otherwise *.spi will be made"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   120
            TabIndex        =   11
            Top             =   -60
            Width           =   3900
         End
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Split"
         Height          =   375
         Left            =   3900
         TabIndex        =   8
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   315
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   315
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   315
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   960
         Width           =   315
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   960
         Width           =   4875
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   360
         Width           =   4875
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Split progress"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   3600
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Per splited file size"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   1380
         Width           =   1290
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To: (Folder)"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From: (File)"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   120
         Width           =   765
      End
   End
End
Attribute VB_Name = "fspForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================================================
'THE FILESPLITTER
'====================================================================================
'
'This program is free software; you can redistribute it and/or modify
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

Private Sub Check2_Click()
    Option4.Enabled = Check2.Value
    Option5.Enabled = Check2.Value
End Sub

Private Sub HScroll1_Change()
    Text3.Text = HScroll1.Value
End Sub

Private Sub Command1_Click()
    SelFolder Me.hWnd, "Select any Driver or Folder to move the splitted files", 0
End Sub

Private Sub Command2_Click()
    CommonDialog1.FileName = ""
    CommonDialog1.Filter = "All Files(*.*)|*.*"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then Text1.Text = CommonDialog1.FileName
End Sub

Private Sub Command3_Click()
    c6.Width = 0
    Dim pr1 As String, pr2 As String, pr3 As Long
    pr1 = Text1.Text
    pr2 = Text2.Text
    Dim ar() As String
    ar = Split(Combo1.Text, " ")
    If LCase(ar(1)) = "byte" Then pr3 = CLng(ar(0))
    If LCase(ar(1)) = "kb" Then pr3 = Val(ar(0)) * 1024
    If LCase(ar(1)) = "mb" Then pr3 = Val(ar(0)) * 1024 * 1024
    If LCase(ar(1)) = "gb" Then pr3 = Val(ar(0)) * 1024 * 1024 * 1024
    fspFunction.fsp pr1, pr2, pr3
End Sub

Private Sub Command4_Click()
    CommonDialog1.FileName = ""
    CommonDialog1.Filter = "Splitted file information (*.spi)|*.spi"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then Text3.Text = CommonDialog1.FileName
End Sub

Private Sub Command5_Click()
    c7.Width = 0
    Dim pr1 As String, pr2 As String
    pr1 = Text3.Text
    pr2 = Text4.Text
    fspmk pr1, pr2
End Sub

Private Sub Command6_Click()
    SelFolder Me.hWnd, "Select any Driver or Folder to marge", 1
End Sub

Private Sub Form_Load()
    Me.Caption = "FileSplitter " & App.Major & "." & App.Minor & "." & App.Revision
    hdr2.Caption = Me.Caption
    c6.Width = 0
    c7.Width = 0
    lblVersion.Caption = App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Option1_Click(Index As Integer)
    If Option1(0).Value Then
        P1.Visible = True
        Option1(0).BackColor = &HC9C9C9
        Option1(1).BackColor = &HF4F4F4
        Option1(2).BackColor = &HF4F4F4
        p2.Visible = False
        p3.Visible = False
    End If
    If Option1(1).Value Then
        P1.Visible = False
        p2.Visible = True
        Option1(0).BackColor = &HF4F4F4
        Option1(1).BackColor = &HC9C9C9
        Option1(2).BackColor = &HF4F4F4
        p3.Visible = False
    End If
    If Option1(2).Value Then
        P1.Visible = False
        p2.Visible = False
        p3.Visible = True
        Option1(0).BackColor = &HF4F4F4
        Option1(1).BackColor = &HF4F4F4
        Option1(2).BackColor = &HC9C9C9
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

