<!--#include file="./Geetestlib.asp"-->
<%
Dim result,geetest_challenge,geetest_validate,geetest_seccode
Dim geetest

Set geetest = (new Geetestlib)("55e100f0bd6cc6279d47b80b3ccedbf7")
geetest_challenge = Request("geetest_challenge")
geetest_validate = Request("geetest_validate")
geetest_seccode = Request("geetest_seccode")
result = geetest.CheckValidate(geetest_challenge,geetest_validate,geetest_seccode)



If result Then
	Response.Write "OK"
Else
    	Response.Write "Error"
End If
Response.End()
%>