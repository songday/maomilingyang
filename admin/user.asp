<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="checkuser.asp"-->
<%checkuser("root")%>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/md5.asp"-->
<!--#include file="../inc/utf8.asp"-->
<%
Dim action,username,pass1,pass2,role,msg,users
Dim rs,cmd
Set rs = server.createobject("adodb.recordset")
action = request.QueryString("action")
if action = "add" then
	username = Replace(trim(request.Form("username")),"'","")
	pass1 = trim(request.Form("pass1"))
	pass2 = trim(request.Form("pass2"))
	role = trim(request.Form("role"))
	if username = "" then
		msg = "请输入用户名"
	elseif pass1 = "" then
		msg = "请输入密码"
	elseif pass2 = "" then
		msg = "请输入确认密码"
	elseif pass1 <> pass2 then
		msg = "密码和确认密码不一致"
	elseif role = "" then
		msg = "请选择权限"
	else
		cmd = "SELECT username FROM admin WHERE username='"&username&"'"
		rs.open cmd,conn,0,1
		if rs.eof and rs.bof then
			pass1 = md5(pass1)
			cmd = "Insert Into admin (username,[password],role) values ('"&username&"','"&pass1&"','"&role&"')"
			conn.execute cmd
			msg = "添加成功"
		else
			msg = "添加的管理员名称重复"
		end if
	end if
elseif action = "update" then
	username = Replace(trim(request.QueryString("username")),"'","")
	pass1 = trim(request.Form("pass1"))
	pass2 = trim(request.Form("pass2"))
	role = trim(request.Form("role"))
	if username = "" then
		msg = "请输入用户名"
	elseif pass1 <> pass2 then
		msg = "密码和确认密码不一致"
	elseif role = "" then
		msg = "请选择权限"
	else
		cmd = "Update admin Set "
		if pass1 <> "" then
			pass1 = md5(pass1)
			cmd = cmd&"[password] = '"&pass1&"',"
		end if
		cmd = cmd&"role='"&role&"' where username='"&username&"'"
		conn.execute cmd
		msg = "修改成功"
	end if
elseif action = "showadmin" then
	Dim operateText,parameter
	if request.QueryString("f") = "update" then
		operateText = "编辑"
		parameter = "f"
	elseif request.QueryString("f") = "del" then
		operateText = "删除"
		parameter = "action"
	end if
	cmd = "Select username,role From admin Order By Id Desc"
	rs.open cmd,conn,0,1
	do while NOT rs.EOF
		users = users&"<tr><td>"&rs("username")&"</td><td>"&rs("role")&"</td><td><a href=javascript:goto('"&parameter&"','"&request.QueryString("f")&"','"&rs("username")&"','"&rs("role")&"')>"&operateText&"</a></td></tr> "
		rs.movenext
	Loop
	if users <> "" then
		users = "<table><tr><td>名称</td><td>权限</td><td>操作</td></tr>"&users
		users = users&"</table>"
	end if
elseif action = "del" then
	username = Replace(trim(request.QueryString("username")),"'","")
	cmd = "Delete from admin Where username = '"&username&"'"
	conn.execute cmd
	response.Write("<script>alert('删除成功！');location.href='user.asp?action=showadmin&f=del';</script>")
	response.End()
end if
closeRs()
closeConn()
Dim actionUrl,bottonText
if request.QueryString("f") = "update" then
	actionUrl = "update"
	bottonText = "更新"
else
	actionUrl = "add"
	bottonText = "添加"
end if
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>添加管理员</title>
<link href="/css/standard.css" rel="stylesheet" type="text/css" />
<script language="javascript">
function goto(p,f,u,r){
	if (f == "del"){
		if (confirm("确认删除？"))
			location.href = "user.asp?"+p+"="+f+"&username="+u+"&role="+r;
	}
	else
		location.href = "user.asp?"+p+"="+f+"&username="+u+"&role="+r;
}
</script>
</head>

<body>
<form id="form1" name="form1" method="post" action="?action=<%=actionUrl%>&username=<%=request.QueryString("username")%>">
<%if action <> "showadmin" then%>
  <table width="50%" border="0">
    <%if request.QueryString("f") <> "update" and action <> "update" then%>
    <tr>
      <td align="right">用户名：</td>
      <td><input name="username" type="text" id="username" /></td>
    </tr>
	<%else%>
    <tr>
      <td align="right"></td>
      <td>编辑：<%=request.QueryString("username")%><br />不修改密码则留空</td>
    </tr>
	<%end if%>
    <tr>
      <td width="20%" align="right">密码：</td>
      <td><input name="pass1" type="password" id="pass1" /></td>
    </tr>
    <tr>
      <td align="right">确认密码：</td>
      <td><input name="pass2" type="password" id="pass2" /></td>
    </tr>
    <tr>
      <td align="right">权限：</td>
      <td><select name="role" id="role">
        <option value="root"<%if request.QueryString("role") = "" or request.QueryString("role") = "root" then%> selected="selected"<%end if%>>root</option>
        <option value="system"<%if request.QueryString("role") = "system" then%> selected="selected"<%end if%>>system</option>
        <option value="user"<%if request.QueryString("role") = "user" then%> selected="selected"<%end if%>>user</option>
      </select>      </td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><input type="submit" name="Submit" value="<%=bottonText%>" />
      <input type="button" name="Submit2" value="返回" onclick="top.location='main.asp';" /></td>
    </tr>
    <tr>
      <td></td>
      <td><font color="#FF0000"><%=msg%></font></td>
    </tr>
  </table>
<%else response.Write(users)
end if%>
</form>
</body>
</html>
