Function getmethat(URL1, strDest1)
Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")
objXMLHTTP.open "GET", URL1, false
objXMLHTTP.send()
If objXMLHTTP.Status = 200 Then
	Set objADOStream = CreateObject("ADODB.Stream")
	objADOStream.Open
	objADOStream.Type = 1 'adTypeBinary
	objADOStream.Write objXMLHTTP.ResponseBody
	objADOStream.Position = 0    'Set the stream position to the start
	Set objFSO = Createobject("Scripting.FileSystemObject")
	If objFSO.Fileexists(strDest1) Then objFSO.DeleteFile strDest1
	Set objFSO = Nothing
	objADOStream.SaveToFile strDest1
	objADOStream.Close
	Set objADOStream = Nothing
End if
Set objXMLHTTP = Nothing
getmethat=0
End Function


htmlName="rezo1.html"

Set objFSO=CreateObject("Scripting.FileSystemObject")
Set resFile = objFSO.CreateTextFile(htmlName,True)  

resFile.write ("<html><head><script src=sorttable.js></script></head><body><table  border=1 class=sortable>" & vbCrLf)
resFile.write ("<tr> <th>i</th><th>Date</th> <th>Fname</th> <th>URL</th></tr> " & vbCrLf)


pagina1="https://photokeyonline.com/PhotoShare.aspx?pid="

istart =    15092310
iend = istart+10000


for i=istart to iend step 100
pagina=pagina1+CStr(i)

set xmlhttp = createobject ("msxml2.xmlhttp.3.0")
xmlhttp.open "get", pagina, false
xmlhttp.send
MyText= xmlhttp.responseText
set xmlhttp  = Nothing

PosGuest=InStr(1, MyText,"guestImage")
PosProdSto=InStr(1, MyText,"products_storage")

'MsgBox PosGuest-PosProdSto
'yes - 228
'no -  113

if PosGuest-PosProdSto>114 then
   imgURL1= Mid(MyText, PosProdSto+114, PosGuest-PosProdSto-114-6)    
   imgURL=imgURL1 
'   MsgBox imgURL
   imgURLrev=strReverse(imgURL)
   palka1=InStr(1, imgURLrev,"/")
   imgData=strReverse(Mid(imgURLrev,palka1+1,10))
'   MsgBox imgData
   fname=strReverse(Mid(imgURLrev,1,palka1-1))
'    MsgBox fname
    resFile.write ("<tr> <td>"&Cstr(i)&"</td> <td>"&imgData&"</td> <td>"&fname&"</td> <td><a href="&imgURL&">"&imgURL&"</a></td> </tr>" & vbCrLf )
	
	getmethat imgURL, "i_"&CStr(i)&"_"&fname 
else
   resFile.write ("<tr> <td>"&Cstr(i)&"</td><td>000</td> <td>000</td> <td>000</td> </tr>"& vbCrLf )
end if


WScript.Sleep 3000   ' 7sec


next

resFile.write ("</table></body></html>")
resFile.Close


set shell = WScript.CreateObject("WScript.Shell")
shell.Run "cmd /c  start " + htmlName


'jpg 	https://photokeyonline.com/thumbs/photokey/seaworld orlando/dolphin photo/2016-11-24/8500315_112416_000325179lgthumb.jpg

'  https://photokeyonline.com/Products.aspx

'8540302_112316_059466smthumb.jpg
'8540302_112316_059466lgthumb.jpg


'REAL BIG PHOTOS
'
'   photokeyonline.com/DigitalDownload.aspx?pid=15012326
'https://photokeyonline.com/PhotoShare.aspx?pid=15012326
'  http://photokeyonline.com/p/0_15092328.jpg   --- big foto