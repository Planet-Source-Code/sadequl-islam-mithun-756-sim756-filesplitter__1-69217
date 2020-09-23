VERSION 5.00
Begin VB.Form fspForm1 
   Appearance      =   0  'Flat
   BackColor       =   &H00EAEAEA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FileMerger"
   ClientHeight    =   1860
   ClientLeft      =   3540
   ClientTop       =   5595
   ClientWidth     =   4695
   ControlBox      =   0   'False
   Icon            =   "fspForm1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      BackColor       =   &H00EAEAEA&
      Caption         =   "&Merge"
      Default         =   -1  'True
      Height          =   375
      Left            =   3180
      TabIndex        =   4
      Top             =   1380
      Width           =   1395
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00EAEAEA&
      Caption         =   "Ends after merging"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   1440
      Value           =   1  'Checked
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00EAEAEA&
      Caption         =   "&Exit"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1380
      Width           =   1395
   End
   Begin VB.CommandButton Command6 
      Caption         =   "..."
      Height          =   315
      Left            =   4260
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   315
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBEBEB&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      ScaleHeight     =   285
      ScaleWidth      =   4425
      TabIndex        =   2
      Top             =   960
      Width           =   4455
      Begin VB.PictureBox c7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   0
         ScaleHeight     =   315
         ScaleWidth      =   15
         TabIndex        =   3
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4155
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Merge progress"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Merge location"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1050
   End
   Begin VB.Menu mLanguage 
      Caption         =   "Language"
      Visible         =   0   'False
      Begin VB.Menu mBN_BD 
         Caption         =   "Bangla - Bangladesh"
      End
      Begin VB.Menu mEN_US 
         Caption         =   "English - US"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "fspForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================================================
'PART OF THE FILESPLITTER
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

Private Sub Command1_Click()
    End
End Sub

Private Sub Command5_Click()
    c7.Width = 0
    Dim pr1 As String, pr2 As String
    pr2 = Text4.Text
    fspmk pr2
    If Check1.Value = 1 Then End
End Sub

Private Sub Command6_Click()
    SelFolder Me.hWnd
End Sub

Private Sub Form_Load()
    Me.Caption = "FileMerger " & App.Major & "." & App.Minor & "." & App.Revision
    Text4.Text = App.Path
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
