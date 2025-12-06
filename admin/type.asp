<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="checkuser.asp"-->
<%checkuser("root,system")%>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/utf8.asp"-->
<%
sub showClass(isSubClasses,className)
	Dim cmd,classes
	if isSubClasses then
		cmd = "Select className,class From class Where class<>'' Order By Id Desc"
	else
		cmd = "Select className From class WHere class is null or class='' Order By Id Desc"
	end if
	Set rs = server.createobject("adodb.recordset")
	rs.open cmd,conn,0,1
	do while NOT rs.EOF
		if isSubClasses then
			classes = classes&"<option value="""&rs("class")&"/"&rs("className")&""">"&rs("class")&"/"&rs("className")&"</option>"
		else
			classes = classes&"<option value="""&rs("className")&""">"&rs("className")&"</option>"
		end if
		rs.movenext
	Loop
	rs.close
	Set rs = nothing
	Response.Write(classes)
end sub
Dim cmd,msg,msg1,msg2,action,classname,oldclassname,subclassname,operate,chooseclass
Dim choosesubclass,slashPos,subclasslist,oldsubclassname,classlist
Set rs = server.createobject("adodb.recordset")
action = request.QueryString("action")
if action = "newclass" or action = "modifyclass" then
	classname = Replace(Replace(Trim(request.Form("classname")),"'",""),"/","")
	classlist = Replace(Replace(Trim(request.Form("classlist")),"'",""),"/","")
	if action = "modifyclass" then
		if request.Form("classcaozuo") = "del" then
			cmd = "Delete From class Where class='"&classlist&"'"
			'Response.Write(cmd&"<br>")
			conn.execute cmd
			cmd = "Delete From class Where classname='"&classlist&"'"
			'Response.Write(cmd&"<br>")
			conn.execute cmd
			Response.Write("<script>alert('删除成功');location.href='type.asp';</script>")
			Response.End()
		end if
		classname = Replace(Replace(Trim(request.Form("newclassname")),"'",""),"/","")
	end if
	if len(classname)<1 or len(classname)>50 then
		msg1 = "分类名称的长度应保持在1-50字之间"
	else
		cmd = "Select Id From class Where classname='"&classname&"' And class=''"
		rs.open cmd,conn,0,1
		if rs.eof and rs.bof then
			if action = "newclass" then
				cmd = "Insert into class(classname) Values ('"&classname&"')"
				conn.execute cmd
				msg1 = "添加成功"
			else
				oldclassname = classlist
				cmd = "Update class Set classname='"&classname&"' Where classname='"&oldclassname&"'"
				conn.execute cmd
				cmd = "Update class Set class='"&classname&"' Where class='"&oldclassname&"'"
				conn.execute cmd
				cmd = "Update album Set class='"&classname&"' Where class='"&oldclassname&"'"
				conn.execute cmd
				msg1 = "修改成功"
			end if
		else
			msg1 = "添加的分类名称重复"
		end if
	end if
elseif action = "newsubclass" or action = "modifysubclass" then
	classname = Replace(Replace(Trim(request.Form("chooseclasslist")),"'",""),"/","")
	subclassname = Replace(Replace(Trim(request.Form("subclassname")),"'",""),"/","")
	subclasslist = Replace(Trim(request.Form("subclasslist")),"'","")
	slashPos = instr(subclasslist,"/")
	oldclassname = Left(subclasslist,slashPos - 1)
	oldsubclassname = Right(subclasslist,len(subclasslist) - slashPos)
	if action = "modifysubclass" then
		if request.Form("subclasscaozuo") = "del" then
			cmd = "Delete From class Where classname='"&oldsubclassname&"' and class='"&oldclassname&"'"
			conn.execute cmd
			Response.Write("<script>alert('删除成功');location.href='type.asp';</script>")
			Response.End()
		end if
		classname = Replace(Replace(Trim(request.Form("changeclassname")),"'",""),"/","")
		subclassname = Replace(Replace(Trim(request.Form("newsubclassname")),"'",""),"/","")
	end if
	if len(subclassname)<1 or len(subclassname)>50 then
		msg2 = "分类名称的长度应保持在1-50字之间"
	else
		cmd = "Select Id From class Where classname='"&subclassname&"' And class='"&classname&"'"
		rs.open cmd,conn,0,1
		if rs.eof and rs.bof then
			if action = "newsubclass" then
				cmd = "Insert into class(classname,class) Values ('"&subclassname&"','"&classname&"')"
				conn.execute cmd
				msg2 = "添加成功"
			else
				if slashPos > 0 then
					cmd = "Update class Set classname='"&subclassname&"',class='"&classname&"' Where classname='"&oldsubclassname&"' and class='"&oldclassname&"'"
					conn.execute cmd
					cmd = "Update album Set subclass='"&subclassname&"',class='"&classname&"' Where subclass='"&oldsubclassname&"' and class='"&oldclassname&"'"
					conn.execute cmd
					msg2 = "更新成功"
				else
					msg2 = "参数传递错误"
				end if
			end if
		else
			msg2 = "添加的分类名称重复"
		end if
	end if
end if
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>分类管理</title>
<link href="/css/standard.css" rel="stylesheet" type="text/css" />
<script language="javascript">
function toSubmit(url){
	document.form1.action = url;
	document.form1.submit();
}
</script>
</head>

<body>
<form id="form1" name="form1" method="post" action="">
  <table width="70%" border="0">
    <tr>
      <td>添加一级分类： 
      <input name="classname" type="text" id="classname" />
      <input type="button" name="Button" value="添加" onclick="toSubmit('?action=newclass');" />
	  <input type="button" name="Submit22222" value="返回" onclick="top.location='main.asp';" />
	  </td>
    </tr>
    <tr>
      <td>管理一级分类：
        <select name="classlist" id="classlist">
		<%call showClass(false,"")%>
        </select>
        <select name="classcaozuo" id="classcaozuo">
          <option value="modify" selected="selected">修改</option>
          <option value="del">删除</option>
        </select>
      <input type="button" name="Submit2" value="确定" onclick="toSubmit('?action=modifyclass');" />
      新名称：
      <input name="newclassname" type="text" id="newclassname" /></td>
    </tr>
    <tr>
      <td><font color="#FF0000"><%=msg1%></font></td>
    </tr>
  </table>
  <hr />
  <table width="100%" border="0">
    <tr>
      <td>添加二级分类：
      <input name="subclassname" type="text" id="subclassname" />
      所属：
      <select name="chooseclasslist" id="chooseclasslist">
        <%call showClass(false,"")%>
      </select>
      <input type="button" name="Submit3" value="添加" onclick="toSubmit('?action=newsubclass');" />
	  <input type="button" name="Submit212312" value="返回" onclick="top.location='main.asp';" />
	  </td>
    </tr>
    <tr>
      <td>管理二级分类：
        <select name="subclasslist" id="subclasslist">
          <%call showClass(true,"")%>
        </select>
        <select name="subclasscaozuo" id="subclasscaozuo">
          <option value="modify" selected="selected">修改</option>
          <option value="del">删除</option>
        </select>
      <input type="button" name="Submit4" value="确定" onclick="toSubmit('?action=modifysubclass');" />
      新名称：
      <input name="newsubclassname" type="text" id="newsubclassname" />
      所属：
      <select name="changeclassname" id="changeclassname">
        <%call showClass(false,"")%>
      </select></td>
    </tr>
    <tr>
      <td><font color="#FF0000"><%=msg2%></font></td>
    </tr>
  </table>
</form>
</body>
</html><%
closeRs()
closeConn()
%>