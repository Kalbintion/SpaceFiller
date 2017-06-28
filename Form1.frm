VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   11055
   ClientTop       =   3870
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   6585
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SHGetDiskFreeSpace Lib "shell32" Alias "SHGetDiskFreeSpaceA" (ByVal pszVolume As String, pqwFreeCaller As Currency, pqwTot As Currency, pqwFree As Currency) As Long

Private Sub Drive1_Change()

End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Dim FreeCaller As Currency, Tot As Currency, Free As Currency
    Dim fPath As String, fNum As Long
    fPath = Left$(App.Path, 2) & "\kdkSpaceFiller.kdk"
    fNum = FreeFile()
    Open fPath For Output As fNum
    
    Do
        SHGetDiskFreeSpace Left$(App.Path, 2) & "\", FreeCaller, Tot, Free
        Free = Free * 10000
        Select Case Free
            Case Is > 100000000
                Print #fNum, String(100000000, Chr(Rnd * 255 + 1))
            Case Is > 10000000
                Print #fNum, String(10000000, Chr(Rnd * 255 + 1))
            Case Is > 1000000
                Print #fNum, String(1000000, Chr(Rnd * 255 + 1))
            Case Is > 100000
                Print #fNum, String(100000, Chr(Rnd * 255 + 1))
            Case Is > 10000
                Print #fNum, String(10000, Chr(Rnd * 255 + 1))
            Case Is > 1000
                Print #fNum, String(1000, Chr(Rnd * 255 + 1))
            Case Is > 100
                Print #fNum, String(Free - 6, Chr(Rnd * 255 + 1))
        End Select
    Loop While Free > 0
    
    Close #fNum
    
    End
End Sub
