<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="inc/utf8.asp"-->
<!--#include file="inc/conn.asp"-->
<%
Dim title,show,table,content,riqi,id,cmd
id = 0
show = request.QueryString("show")
if show = "changjianwenda" then
	table = "changjianwenda"
else
	table = "affiche"
end if
if request.QueryString("id") <> "" then
	id = cint(request.QueryString("id"))
end if
if id > 0 then
	cmd = "Select title,content,riqi From "&table&" Where id="&id
	Set rs = server.createobject("adodb.recordset")
	rs.open cmd,conn,0,1
	if rs.bof and rs.eof then
		Response.Write("参数传递错误。")
		Response.End()
	else
		title = rs("title")
		content = rs("content")
		riqi = rs("riqi")
	end if
else
	Response.Write("参数传递错误。")
	Response.End()
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=title%></title>
<link href="/css/standard.css" rel="stylesheet" type="text/css" />
</head>

<body>
<table id="Table_01" width="760" border="0" cellpadding="0" cellspacing="0" align="center">
	<tr>
		<td>
			<img src="images/lingyang4_01.gif" width="760" height="37" alt=""></td>
	</tr>
	<tr>
		<td>
			<img src="images/lingyang4_02.gif" width="760" height="47" alt=""></td>
	</tr>
	<tr>
		<td>
			<img src="images/lingyang4_03.gif" width="760" height="59" alt=""></td>
	</tr>
	<tr>
		<td>
			<img src="images/lingyang4_04.gif" width="760" height="68" alt=""></td>
	</tr>
</table>
<table width="760" border="0" align="center">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td align="center"><strong><%=title%></strong> 日期：<%=riqi%></td>
  </tr>
  <tr>
    <td><%=content%></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<!--#include file="bottom.asp"-->
</body>
</html>