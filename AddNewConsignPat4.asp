<!--#INCLUDE FILE="clsUpload.asp"-->
<%
Dim oUpload
Dim oFile
Dim sFileName
Dim oFSO
Dim sPath
Dim sNewData
Dim nLength
Dim oFile2
Dim bytBinaryData

Const nForReading = 1
Const nForWriting = 2
Const nForAppending = 8

Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("./") & "\db\NWHQ.mdb"
 
Set oRs2 = Server.CreateObject("ADODB.Recordset")
sSQL2 = "SELECT Top 1 pID, InitCost From tblConsignPatterns Order By pID Desc"
oRs2.Open sSQL2, oConn, 3, 3

pID = oRs2.Fields("pID")

' grab the uploaded file data
Set oUpload = New clsUpload
Set oFile = oUpload("File1")

' parse the file name
sFileName = pID & Right((oFile.FileName), 4)
sFileName_Original = oFile.FileName
If Not InStr(sFileName, "\") = 0 Then
	sFileName = Mid(sFileName, InStrRev(sFileName, "\") + 1)
End If

If Not InStr(sFileName_Original, "\") = 0 Then
	sFileName_Original = Mid(sFileName_Original, InStrRev(sFileName_Original, "\") + 1)
End If


' Convert the binary data to Ascii
bytBinaryData = oFile.BinaryData
nLength = LenB(bytBinaryData)
For nIndex = 1 To nLength
	sNewData = sNewData & Chr(AscB(MidB(bytBinaryData, nIndex, 1)))
Next

' Save the file to the file system
sPath = Server.MapPath(".\dev\ConsignmentShop") & "\"
Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
oFSO.OpenTextFile(sPath & sFileName, nForWriting, True).Write sNewData
Set oFSO = Nothing


Set oFile1 = oFile


Set oFile = Nothing
Set oUpload = Nothing

'********************************************************************************


Set oRs = Server.CreateObject("ADODB.Recordset")

sSQL = "SELECT FileID, FileName, Name, FileSize, pID, ContentType, BinaryData FROM tblConsignmentFiles WHERE 1=2"

oRs.Open sSQL, oConn, 3, 3



		oRs.AddNew
		oRs.Fields("FileName") = sFileName_Original
		oRs.Fields("Name") = sFileName
		oRs.Fields("pID") = pID
		oRs.Fields("FileSize") = oFile1.Length
		oRs.Fields("ContentType") = oFile1.ContentType
		oRs.Fields("BinaryData").AppendChunk = oFile1.BinaryData & ChrB(0)
		oRs.Update


'*************************************************************************************************

%>

<html>
<head></head>
<BODY background="paper_old.gif">
<STYLE type=text/css>
	p {font-size: 8pt;font-family: "Verdana";color: "darkblue"; }
	td {font-size: 11pt;font-family: "Verdana";color: "darkblue"; }
	th {font-size: 12pt;font-family: "Verdana"; color: "black";}
	a {font-size: 8pt;font-family: "Verdana"; color: "blue";}
</STYLE>
<BR><center>
<%

Set oRs4 = Server.CreateObject("ADODB.Recordset")
sSQL4 = "SELECT FileID, FileName, Name, FileSize, pID FROM tblConsignmentFiles WHERE pID = " & pID
oRs4.Open sSQL4, oConn, 3, 3


Do While Not oRs4.EOF
	If InStr(oRs4.Fields("Name"),".jpg") > 0 Then
		FileSize_JPG = CLng(oRs4.Fields("FileSize"))
		fJPEG = oRs4.Fields("FileName")
	ElseIf InStr(oRs4.Fields("Name"),".pat") > 0 Then	
		FileSize_PAT = CLng(oRs4.Fields("FileSize"))
		fPAT = oRs4.Fields("FileName")
	Else
		Response.Write "Error in Calculation."
	End If		

oRs4.MoveNext 
Loop


costJPG = Round(CCur((CDbl(FileSize_JPG) * 0.00007)),2)
costPAT = Round(CCur((CDbl(FileSize_PAT) * 0.00007)),2)
costTOTAL = Round((costJPG + costPAT),2)

Response.Write "<center><h3>Cost Summary</h3></center>"
Response.Write "<center><table bgcolor=darkblue cellpadding=1 cellspacing=1 border=1 bordercolor=black>"
Response.Write "<tr bgcolor=silver><th><center>File Type</center></th><th><center>File Size (bytes)</center></th><th><center>Cost (File Size * 0.0001)</center></th></tr>"
Response.Write "<tr bgcolor=lightblue><td><center><b>" & fJPEG & "</b></center></td><td><center>" & FileSize_JPG & "</center></td><td><center>" & "<b>$" & costJPG & "</b></center></td></tr>"
Response.Write "<tr bgcolor=lightblue><td><center><b>" & fPAT & "</b></center></td><td><center>" & FileSize_PAT & "</center></td><td><center>" & "<b>$" & costPAT & "</b></center></td></tr>"
Response.Write "<tr bgcolor=pink><td>Total:</td><td><center>" & (FileSize_JPG + FileSize_PAT) & "</center></td><td bgcolor=pink><center><b><font color=darkred>$ " & costTOTAL & "</font></b></center></td></tr>"
Response.Write "</table>"

oRs2.Fields("InitCost") = CCur(costTOTAL)
oRs2.Update 

oRs2.Close
Set oRs2 = Nothing


%>


<% 

uid = Request.QueryString("u").Item(1)
pass = Request.QueryString("u").Item(2) 

%>


<center><h3><b>Your total cost is <font color=red>$<%= costTotal %></font></b></h3></center>
<table background="yellow.jpg"><tr><td>
<H5><Font color="red"></Font>To activate your new consignment pattern, please pay the amount of <font color="red"><%= "$" & costTotal %>. </font>If paying by credit card, please click the <I>PayPal</I> payment button below and follow the payment instructions. Your pattern will be automatically activated upon the completion of the transaction. 
<BR>
<BR><Center> 
<form action="https://www.paypal.com/cgi-bin/webscr" method="post" id=form1 name=form1>
<input type="hidden" name="cmd" value="_xclick">
<input type="hidden" name="business" value="sales@crossstitchconnection.com">
<input type="hidden" name="item_name" value="Cross Stitch Connection Consignment Shop Pattern">
<input type="hidden" name="item_number" value="1">
<!--
<input type="hidden" name="amount" value="<%= costTotal %>">
-->
<input type="hidden" name="amount" value="0.00">
<input type="hidden" name="return" value="http://www.crossstitchconnection.com/CSA.asp?u=<%= "1&u=" & uid & "&u=" & pass & "&u=" & pID %>">
<input type="hidden" name="cancel_return" value="http://www.crossstitchconnection.com/CSA.asp?u=<%= "0&u=" & uid & "&u=" & pass %>">
<input type="image" src="x-click-butcc.gif" border="0" name="submit">
</form>
<BR>
</Center>
<!--
<center>*For other payment arrangements (<font color="darkgreen">check</font> or <font color="darkgreen">money order</font>) <a href="OtherPaymentArrangements2.asp">click here</a></center>
-->
</td></tr></table>


</BODY>
<%

oRs.Close

oConn.Close

Set oConn = Nothing


%>
</html>

