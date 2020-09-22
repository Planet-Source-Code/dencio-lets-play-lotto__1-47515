Attribute VB_Name = "Module1"
Option Explicit

Type LottoResult
    Game As String * 11
    GameMonth As String * 2
    GameDay As String * 2
    GameYear As String * 4
    WinningCombination As String * 17
    Moneypot As String * 16
    Winner As String * 30
End Type

Public FNum
Public Lresult As LottoResult

Public Sub OpenLottoFile()


FNum = FreeFile()
Open App.Path & "\LottoResult.lgr" For Random As #FNum Len = Len(Lresult)


End Sub

Public Sub ViewLottoResult()
FNum = FreeFile()
Open App.Path & "\LottoResult.lgr" For Random As #FNum Len = Len(Lresult)


End Sub

Public Sub EditLottoResult()
Dim FNum, TotRecords, i
Dim Lresult As LottoResult
Lresult.Game = frmView.Text1.Text

FNum = FreeFile()
Open App.Path & "\LottoResult.dat" For Random As #FNum Len = Len(Lresult)
Put #FNum, 1, Lresult
Close #FNum

End Sub


