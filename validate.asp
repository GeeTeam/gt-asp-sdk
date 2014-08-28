<!--#include file="./geetest_Md5.asp"-->
<!--#include file="./Geetestlib.asp"-->
<%
Dim result,geetest_challenge,geetest_validate,geetest_seccode
Dim geetest

Set geetest = new Geetestlib
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