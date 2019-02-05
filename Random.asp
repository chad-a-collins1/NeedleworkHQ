<%
Function fn_GetRandomAlphaNumeric(intMinLen, intMaxLen)

Dim i 'As Integer
Dim j 'As Integer
Dim strTmp 'As String
Dim strPartPAss 'As String
Dim intNum 'As Integer
Dim intLower 'As String
Dim intUpper 'As String
Dim intRand 'As Integer
Dim intLen 'As Integer
    
    strTmp = ""
    Randomize
    intLen = Int((intMaxLen - intMinLen + 1) * Rnd + intMinLen)
    
    For i = 1 To intLen
      intNum = CInt(9 * Rnd) + 48
      intUpper = CInt(25 * Rnd) + 65
      intLower = CInt(25 * Rnd) + 97
      intRand = CInt(2 * Rnd) + 1
      If i = 1 Then
         strPartPAss = Chr(intLower)
      Else
         Select Case intRand
         Case 1
           strPartPAss = Chr(intNum)
         Case 2
           strPartPAss = Chr(intUpper)
         Case 3
           strPartPAss = Chr(intLower)
         End Select
      End If
      strTmp = strTmp & strPartPAss
    Next
    
    fn_GetRandomAlphaNumeric = strTmp

End Function

%>