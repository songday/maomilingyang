<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/md5.asp"-->
<!--#include file="../inc/utf8.asp"-->
<%
Dim action,username,password,msg
action = request.QueryString("action")
if action = "login" then
	username = Replace(request.Form("username"),"'","")
	password = request.Form("password")
	Dim rs,cmd
	cmd = "SELECT username,role FROM admin WHERE username='"&username&"' AND password='"&md5(password)&"'"
	Set rs = server.createobject("adodb.recordset")
	rs.open cmd,conn,0,1
	if rs.eof and rs.bof then
		msg = "用户名或密码错误。"
	else
		session("username") = rs("username")
		session("role") = rs("role")
		if request.QueryString("url") <> "" then
			response.Redirect(request.QueryString("url"))
		else
			response.Redirect("main.asp")
		end if
	end if
elseif request.QueryString("url") <> "" then
	msg = "登录超时或权限不够，请重新登录。"
end if
'response.Write(md5("123"))
'response.Write(request.ServerVariables("PATH_INFO"))
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>管理员登录</title>
<link href="/css/standard.css" rel="stylesheet" type="text/css" />
<script language="javascript">
function goback(){
	top.location = "../index.asp";
}
</script>
</head>

<body>
<form id="form1" name="form1" method="post" action="?action=login&url=<%=Server.URLEncode(request.QueryString("url"))%>">
  <table width="50%" border="0" align="center">
    <tr>
      <td width="20%" align="right">用户名：</td>
      <td><input name="username" type="text" id="username" style="width:200px;" /></td>
    </tr>
    <tr>
      <td align="right">密码：</td>
      <td><input name="password" type="password" id="password" style="width:200px;" /></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><input type="submit" name="Submit" value="登录" />
      <input type="button" name="Submit2" value="返回" onclick="goback();" /></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><font color="#FF0000"><%=msg%></font></td>
    </tr>
  </table>
</form>

</body>
</html>