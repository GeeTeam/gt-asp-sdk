<%
Class Geetestlib

Public privateKey

Private Sub Class_Initialize()
	privateKey = "55e100f0bd6cc6279d47b80b3ccedbf7"
End Sub

Private Sub Class_Terminate()
	
End Sub

Public Function CheckValidate(challenge,validate,seccode)
	Dim host,path,port,query,responseText
	host = "http://api.geetest.com"
	path = "/validate.php"
	port = 80
	If len(validate)>0 And checkResultByPrivate(challenge,validate)=True Then
		query="seccode="&seccode
		responseText = postValidate(host, path, query, port)
		if responseText = MD5(seccode,32) Then
			CheckValidate = True
			Exit Function
		End If
	End If
	CheckValidate = False
End Function

Private Function checkResultByPrivate(origin,validate)
	Dim encodeStr
	encodeStr = MD5(privateKey&"geetest"&origin,32)
	If validate = encodeStr Then
		checkResultByPrivate = True
	Else
		checkResultByPrivate = False
	End If
End Function

Private Function postValidate(host, path, data, port)
	Dim url,sMyXmlHTTP
	url = host & path
	
	Set sMyXmlHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
	sMyXmlHTTP.Open "POST", url, False
	sMyXmlHTTP.setRequestHeader "Content-Length", Len(data)
	sMyXmlHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;"
	sMyXmlHTTP.send data
	
	If sMyXmlHTTP.readyState = 4 And sMyXmlHTTP.Status = 200 Then
	    postValidate = sMyXmlHTTP.responseText
	Else
		postValidate = "error"
	End If
End Function

End Class
%>