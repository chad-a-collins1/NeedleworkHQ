<%

' Function Get Time Zones
'******************************************************************
Function fn_DL_GetTimeZones(arryTmp)

    Dim cmdTmp 'As ADODB.Command
    Dim rsTmp 'As ADODB.Recordset
    Dim strDescrips 'As String

    fn_DL_GetTimeZones = -1
   
    Set cmdTmp = Server.CreateObject("ADODB.Command")
    Set rsTmp = Server.CreateObject("ADODB.Recordset")
   '**************************************************
    With cmdTmp
       .ActiveConnection = fn_GetConnectionString
       .CommandText = "sp_GetTimeZones"
       .CommandType = adCmdStoredProc

       .Parameters.Append .CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, , -1)

       Set rsTmp = .Execute
   End With   'cmdTmp
   
   '****************************************************
   Dim j       ' As Integer

   j = 0
   With rsTmp
      Do While (Not rsTmp Is Nothing)
         If .State = adStateClosed Then Exit Do
        
         If Not .EOF Then
             If j = 0 Then
                ' arrays are structured as the following -> arry(col , row)
                arryTmp = .GetRows()
             End If  
         End If
       
         Set rsTmp = .NextRecordset
         j = j + 1
      Loop
   End With   'rsTmp
       
   fn_DL_GetTimeZones = cmdTmp.Parameters("@RETURN_VALUE")
    
   Set cmdTmp = Nothing
   Set rsTmp = Nothing
    
End Function  'fn_DL_GetTimeZones



' Function Get Account Status
'******************************************************************
Function fn_DL_GetAccountStatuses(arryTmp)

    Dim cmdTmp 'As ADODB.Command
    Dim rsTmp 'As ADODB.Recordset
    Dim strDescrips 'As String

    fn_DL_GetAccountStatuses = -1
   
    Set cmdTmp = Server.CreateObject("ADODB.Command")
    Set rsTmp = Server.CreateObject("ADODB.Recordset")
   '**************************************************
    With cmdTmp
       .ActiveConnection = fn_GetConnectionString
       .CommandText = "sp_GetAccountStatuses"
       .CommandType = adCmdStoredProc

       .Parameters.Append .CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, , -1)

       Set rsTmp = .Execute
   End With   'cmdTmp
   
   '****************************************************
   Dim j       ' As Integer

   j = 0
   With rsTmp
      Do While (Not rsTmp Is Nothing)
         If .State = adStateClosed Then Exit Do
        
         If Not .EOF Then
             If j = 0 Then
                ' arrays are structured as the following -> arry(col , row)
                arryTmp = .GetRows()
             End If  
         End If
       
         Set rsTmp = .NextRecordset
         j = j + 1
      Loop
   End With   'rsTmp
       
   fn_DL_GetAccountStatuses = cmdTmp.Parameters("@RETURN_VALUE")
    
   Set cmdTmp = Nothing
   Set rsTmp = Nothing
    
End Function  'fn_DL_GetAccountStatuses


' Function Get Account Types
'******************************************************************
Function fn_DL_GetAccountTypes(arryTmp)

    Dim cmdTmp 'As ADODB.Command
    Dim rsTmp 'As ADODB.Recordset
    Dim strDescrips 'As String

    fn_DL_GetAccountTypes = -1
   
    Set cmdTmp = Server.CreateObject("ADODB.Command")
    Set rsTmp = Server.CreateObject("ADODB.Recordset")
   '**************************************************
    With cmdTmp
       .ActiveConnection = fn_GetConnectionString
       .CommandText = "sp_GetAccountTypes"
       .CommandType = adCmdStoredProc

       .Parameters.Append .CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, , -1)

       Set rsTmp = .Execute
   End With   'cmdTmp
   
   '****************************************************
   Dim j       ' As Integer

   j = 0
   With rsTmp
      Do While (Not rsTmp Is Nothing)
         If .State = adStateClosed Then Exit Do
        
         If Not .EOF Then
             If j = 0 Then
                ' arrays are structured as the following -> arry(col , row)
                arryTmp = .GetRows()
             End If  
         End If
       
         Set rsTmp = .NextRecordset
         j = j + 1
      Loop
   End With   'rsTmp
       
   fn_DL_GetAccountTypes = cmdTmp.Parameters("@RETURN_VALUE")
    
   Set cmdTmp = Nothing
   Set rsTmp = Nothing
    
End Function  'fn_DL_GetAccountTypes



'
' Get States
'******************************************************************************
Function fn_DL_GetStates(arryTmp)

    Dim cmdTmp 'As ADODB.Command
    Dim rsTmp 'As ADODB.Recordset
    Dim strDescrips 'As String

    fn_DL_GetStates = -1
   
    Set cmdTmp = Server.CreateObject("ADODB.Command")
    Set rsTmp = Server.CreateObject("ADODB.Recordset")
   '**************************************************
    With cmdTmp
       .ActiveConnection = fn_GetConnectionString
       .CommandText = "sp_GetStates"
       .CommandType = adCmdStoredProc

       .Parameters.Append .CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, , -1)

       Set rsTmp = .Execute
   End With   'cmdTmp
   
   '****************************************************
   Dim j       ' As Integer

   j = 0
   With rsTmp
      Do While (Not rsTmp Is Nothing)
         If .State = adStateClosed Then Exit Do
        
         If Not .EOF Then
             If j = 0 Then
                ' arrays are structured as the following -> arry(col , row)
                arryTmp = .GetRows()
             End If  
         End If
       
         Set rsTmp = .NextRecordset
         j = j + 1
      Loop
   End With   'rsTmp
       
   fn_DL_GetStates = cmdTmp.Parameters("@RETURN_VALUE")
    
   Set cmdTmp = Nothing
   Set rsTmp = Nothing
    
End Function  'fn_DL_GetStates


' Function Get Credit Card Types
'******************************************************************
Function fn_DL_GetCreditCardTypes(arryTmp)

    Dim cmdTmp 'As ADODB.Command
    Dim rsTmp 'As ADODB.Recordset

    fn_DL_GetCreditCardTypes = -1
   
    Set cmdTmp = Server.CreateObject("ADODB.Command")
    Set rsTmp = Server.CreateObject("ADODB.Recordset")
   '**************************************************
    With cmdTmp
       .ActiveConnection = fn_GetConnectionString
       .CommandText = "sp_GetCreditCardTypes"
       .CommandType = adCmdStoredProc

       .Parameters.Append .CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, , -1)

       Set rsTmp = .Execute
   End With   'cmdTmp
   
   '****************************************************
   Dim j       ' As Integer

   j = 0
   With rsTmp
      Do While (Not rsTmp Is Nothing)
         If .State = adStateClosed Then Exit Do
        
         If Not .EOF Then
             If j = 0 Then
                ' arrays are structured as the following -> arry(col , row)
                arryTmp = .GetRows()
             End If  
         End If
       
         Set rsTmp = .NextRecordset
         j = j + 1
      Loop
   End With   'rsTmp
       
   fn_DL_GetCreditCardTypes = cmdTmp.Parameters("@RETURN_VALUE")
    
   Set cmdTmp = Nothing
   Set rsTmp = Nothing
    
End Function  'fn_DL_GetCreditCardTypes


' Function Get DataCalendarColors
'******************************************************************
Function fn_DL_GetDataCalendarColors(arryTmp)

    Dim cmdTmp 'As ADODB.Command
    Dim rsTmp 'As ADODB.Recordset
    Dim strDescrips 'As String

    fn_DL_GetDataCalendarColors = -1
   
    Set cmdTmp = Server.CreateObject("ADODB.Command")
    Set rsTmp = Server.CreateObject("ADODB.Recordset")
   '**************************************************
    With cmdTmp
       .ActiveConnection = fn_GetConnectionString
       .CommandText = "sp_GetDataCalendarColors"
       .CommandType = adCmdStoredProc

       .Parameters.Append .CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, , -1)

       Set rsTmp = .Execute
   End With   'cmdTmp
   
   '****************************************************
   Dim j       ' As Integer

   j = 0
   With rsTmp
      Do While (Not rsTmp Is Nothing)
         If .State = adStateClosed Then Exit Do
        
         If Not .EOF Then
             If j = 0 Then
                ' arrays are structured as the following -> arry(col , row)
                arryTmp = .GetRows()
             End If  
         End If
       
         Set rsTmp = .NextRecordset
         j = j + 1
      Loop
   End With   'rsTmp
       
   fn_DL_GetDataCalendarColors = cmdTmp.Parameters("@RETURN_VALUE")
    
   Set cmdTmp = Nothing
   Set rsTmp = Nothing
    
End Function  'fn_DL_GetDataCalendarColors


' Get Countries
'******************************************************************************
Function fn_DL_GetCountries(arryTmp)

    Dim cmdTmp 'As ADODB.Command
    Dim rsTmp 'As ADODB.Recordset

    fn_DL_GetCountries = -1
   
    Set cmdTmp = Server.CreateObject("ADODB.Command")
    Set rsTmp = Server.CreateObject("ADODB.Recordset")
   '**************************************************
    With cmdTmp
       .ActiveConnection = fn_GetConnectionString
       .CommandText = "sp_GetCountries"
       .CommandType = adCmdStoredProc

       .Parameters.Append .CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue, , -1)

       Set rsTmp = .Execute
   End With   'cmdTmp
   
   '****************************************************
   Dim j       ' As Integer

   j = 0
   With rsTmp
      Do While (Not rsTmp Is Nothing)
         If .State = adStateClosed Then Exit Do
        
         If Not .EOF Then
             If j = 0 Then
                ' arrays are structured as the following -> arry(col , row)
                arryTmp = .GetRows()
             End If  
         End If
       
         Set rsTmp = .NextRecordset
         j = j + 1
      Loop
   End With   'rsTmp
       
   fn_DL_GetCountries = cmdTmp.Parameters("@RETURN_VALUE")
    
   Set cmdTmp = Nothing
   Set rsTmp = Nothing
    
End Function  'fn_DL_GetCountries



%>





















