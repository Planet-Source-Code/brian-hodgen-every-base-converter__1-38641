Attribute VB_Name = "Module1"
Public Function BaseConvert(nInput As String, oBase As Integer, nBase As Integer) As String
Dim newNum As Double, lNum As Integer, tempNum As Integer
Dim nMod As Integer, nDiv As Double, nTemp As Double
Dim nTempStr As String
Dim i As Integer
newNum = 0
lNum = Len(nInput)
    For i = 1 To lNum
            If Asc(Mid(nInput, i, 1)) > 47 And Asc(Mid(nInput, i, 1)) < 58 Then
                tempNum = CInt(Mid(nInput, i, 1))
            Else
                tempNum = (Asc(UCase$(Mid(nInput, i, 1)))) - 55
            End If
        newNum = newNum + (tempNum * ((oBase ^ (lNum - i))))
    Next i
 If newNum = 0 Then nTempStr = newNum
    Do While newNum <> 0
        nMod = newNum Mod nBase
        nDiv = newNum \ nBase
        newNum = nDiv
        If nMod < 10 Then nTempStr = nMod & nTempStr
        If nMod > 9 Then nTempStr = Chr(nMod + 55) & nTempStr
    Loop
BaseConvert = nTempStr
End Function
