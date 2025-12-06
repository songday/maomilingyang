<%
sub checkuser(u)
	if session("role") = "" or instr(u,session("role"))<1 then
		Dim url
		url = request.ServerVariables("PATH_INFO")&"?"&request.ServerVariables("QUERY_STRING")
		response.Redirect("login.asp?url=" & url)
		response.End()
	end if
end sub
%>