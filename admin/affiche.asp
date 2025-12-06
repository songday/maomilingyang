<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="checkuser.asp"-->
<%checkuser("root,system")%>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/utf8.asp"-->
<%
Dim action,title,content,msg,cmd,showlist,target,table
target = request.QueryString("target")
if target = "changjianwenda" then
	table = "changjianwenda"
else
	table = "affiche"
end if
Set rs = server.createobject("adodb.recordset")
action = request.QueryString("action")
if action = "add" or action = "update" then
	title = Replace(trim(request.Form("title")),"'","")
	content = Replace(trim(request.Form("content")),"'","")
	if len(title)<1 or len(title)>50 then
		msg = "标题的长度应保持在1-50个字之间。"
	elseif len(content)<1 or len(content)>250 then
		msg = "内容的长度应保持在1-250个字之间。"
	else
		if action = "add" then
			cmd = "Insert Into "&table&" (title,content) values ('"&title&"','"&content&"')"
			msg = "添加成功"
		else
			cmd = "update "&table&" set title='"&title&"',content='"&content&"' where id="&cint(request.QueryString("id"))
			msg = "更新成功"
		end if
		conn.execute cmd
	end if
elseif action = "showlist" then
	cmd = "Select id,title,riqi From "&table&" Order By Id Desc"
	rs.open cmd,conn,0,1
	do while NOT rs.EOF
		showlist = showlist&"<tr><td><a href=javascript:goto("&rs("id")&",'"&request.QueryString("f")&"')>"&rs("title")&"</a></td><td>"&rs("riqi")&"</td></tr> "
		rs.movenext
	Loop
	if showlist <> "" then
		showlist = "<table><tr><td>标题</td><td>日期</td></tr>"&showlist
		showlist = showlist&"</table>"
	end if
elseif action = "showupdate" then
	if request.QueryString("id") <> "" then
		cmd = "Select title,content From "&table&" Where Id="&cint(request.QueryString("id"))
		rs.open cmd,conn,0,1
		title = rs("title")
		content = rs("content")
	end if
elseif action = "del" then
	if request.QueryString("id") <> "" then
		cmd = "Delete From "&table&" Where Id="&cint(request.QueryString("id"))
		conn.execute cmd
		response.Write("<script>alert('删除成功！');location.href='affiche.asp?action=showlist&f=del&target="&target&"';</script>")
		response.End()
	end if
end if
closeRs()
closeConn()
Dim actionUrl,bottonText
if action = "showupdate" or action = "update" then
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
<title>公告</title>
<link href="/css/standard.css" rel="stylesheet" type="text/css" />
<script language="javascript">
function goto(id,f){
	if (f=="del"){
		if (confirm("确认删除？"))
			location.href = "affiche.asp?id="+id+"&action="+f+"&target=<%=target%>";
	}
	else
		location.href = "affiche.asp?id="+id+"&action="+f+"&target=<%=target%>";
}
</script>
</head>

<body>
<form id="form1" name="form1" method="post" action="?action=<%=actionUrl%>&id=<%=request.QueryString("id")%>&target=<%=target%>">
<%if action <> "showlist" then%>
  <table width="55%" border="0">
    <tr>
      <td width="20%" align="right">标题：</td>
      <td><input name="title" type="text" id="title" value="<%=title%>" size="50" /></td>
    </tr>
    <tr>
      <td align="right" valign="top">内容：</td>
      <td><textarea name="content" cols="50" rows="5" wrap="virtual" id="content"><%=content%></textarea></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><input type="submit" name="Submit" value="<%=bottonText%>" />
	  <input type="button" name="Submit22222211" value="返回" onclick="top.location='main.asp';" />
	  </td>
    </tr>
    <tr>
      <td></td>
      <td><font color="#FF0000"><%=msg%></font></td>
    </tr>
  </table>
<%else
	response.Write(showlist)
end if%>
</form>
</body>
</html>