<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%><%
	session.Abandon()
	response.Redirect("login.asp")
%>