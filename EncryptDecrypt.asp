<%

   Dim sbox(255)
   Dim key(255)
   
   ' This function Encrypts a String
   '*********************************************************
   Function fn_EncryptString(ByVal strText)
       
       strText = fn_EnDecrypt(strText)
       strText = Replace(strText,Chr(0),"NC999")
   
       fn_EncryptString = strText 
   
   End Function
   
   ' This function Decrypts an Encrypted String
   '*********************************************************
   Function fn_DecryptString(ByVal strText)
       
       strText = Replace(strText,"NC999",Chr(0))
       strText = fn_EnDecrypt(strText)
       
       fn_DecryptString = strText
       
   End Function


':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::  This routine called by EnDeCrypt function. Initializes the :::
':::  sbox and the key array)                                    :::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'*******************************************************************
   Sub RC4Initialize(strPwd)


      dim tempSwap
      dim a
      dim b
      dim intLength
      
      intLength = len(strPwd)
      For a = 0 To 255
         key(a) = asc(mid(strpwd, (a mod intLength)+1, 1))
         sbox(a) = a
      next

      b = 0
      For a = 0 To 255
         b = (b + sbox(a) + key(a)) Mod 256
         tempSwap = sbox(a)
         sbox(a) = sbox(b)
         sbox(b) = tempSwap
      Next
   
   End Sub
 
 
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::  This routine does all the work. Call it both to ENcrypt    :::
':::  and to DEcrypt your data.                                  :::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'***************************************************************      
Function fn_EnDeCrypt(plaintxt)
 
      dim temp
      dim a
      dim i
      dim j
      dim k
      dim cipherby
      dim cipher

      i = 0
      j = 0

      RC4Initialize "wPB5RjV4u416J2"
      'RC4Initialize "1234"

      For a = 1 To Len(plaintxt)
         i = (i + 1) Mod 256
         j = (j + sbox(i)) Mod 256
         temp = sbox(i)
         sbox(i) = sbox(j)
         sbox(j) = temp
   
         k = sbox((sbox(i) + sbox(j)) Mod 256)

         cipherby = Asc(Mid(plaintxt, a, 1)) Xor k
         cipher = cipher & Chr(cipherby)
      Next

      fn_EnDeCrypt = cipher

   End Function


%>

