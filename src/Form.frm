VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form 
   Caption         =   "Переоткрыть все чертежи в каталоге"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10455
   OleObjectBlob   =   "Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Written in 2015 by Eduard E. Tikhenko <aquaried@gmail.com>
'
'To the extent possible under law, the author(s) have dedicated all copyright
'and related and neighboring rights to this software to the public domain
'worldwide. This software is distributed without any warranty.
'You should have received a copy of the CC0 Public Domain Dedication along
'with this software.
'If not, see <http://creativecommons.org/publicdomain/zero/1.0/>

Option Explicit

Private Sub cancelBut_Click()
    ExitApp
End Sub

Private Sub draftBox_Change()
    SaveSetting2 draftProp, draftBox.value
End Sub

Private Sub draftChk_Click()
    EnableChk draftChk, draftBox
End Sub

Private Sub okBut_Click()
    Execute
    ExitApp
End Sub

Private Sub stdBox_Change()
    SaveSetting2 standardFile, stdBox.value
End Sub

Private Sub stdBut_Click()
    Dim filename As String
    
    filename = GetStandardFilename
    If filename <> "" Then
        stdBox.value = filename
    End If
End Sub

Private Sub stdChk_Click()
    EnableChk stdChk, stdBox
End Sub
