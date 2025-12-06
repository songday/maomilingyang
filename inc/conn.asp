<%
Dim conn,connStr
'On Error Resume Next
Set conn = Server.CreateObject("ADODB.Connection")
connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("/data/lingyang.mdb")
conn.Open connStr

sub closeRs()
	On Error Resume Next
	rs.close
	set rs = nothing
end sub

sub closeConn()
	On Error Resume Next
	conn.close
	set conn = nothing
end sub
%>