<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="checkuser.asp"-->
<%'checkuser("root,system")%>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/upload.asp"-->
<!--#include file="../inc/utf8.asp"-->
<%
Function createFolder()
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	Dim foldername
	if fso.FolderExists(Server.MapPath("../album_images/")) = false then
		fso.CreateFolder Server.MapPath("../album_images/")
	end if
	Do
		Randomize
		foldername = Year(Now())&Month(Now())&Day(Now())&Hour(Now())&Minute(Now())&Second(Now())&Replace(Mid((Rnd * 1000000),1,4),".","")
	Loop Until fso.FileExists(Server.MapPath("../album_images/"&foldername)) = false
	fso.CreateFolder Server.MapPath("../album_images/"&foldername)
	Set fso = nothing
	createFolder = foldername
	'Response.Write(folderPath)
end Function

Function delFolder(folderpath)
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	if fso.FileExists(Server.MapPath("../album_images/"&folderpath)) = true then
		fso.DeleteFolder(Server.MapPath("/album_images/"&folderpath))
	end if
	Set fso = nothing
end Function

Function getFilename(folderPath,ext)
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	Dim filename
	Do
		Randomize
		filename = Year(Now())&Month(Now())&Day(Now())&Hour(Now())&Minute(Now())&Second(Now())&Replace(Mid((Rnd * 1000000),1,4),".","")&"."&ext
	Loop Until fso.FileExists(Server.MapPath(folderPath&"/"&filename)) = false
	Set fso = nothing
	getFilename = filename
	'Response.Write(Server.MapPath(folderPath&"/"&filename))
end Function

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
			classes = classes&"<option value="&rs("class")&"/"&rs("className")&">"&rs("class")&"/"&rs("className")&"</option>"
		else
			classes = classes&"<option value="&rs("className")&">"&rs("className")&"</option>"
		end if
		rs.movenext
	Loop
	rs.close
	Set rs = nothing
	Response.Write(classes)
end sub

Function subclassScript()
	Set rs = server.createobject("adodb.recordset")
	Dim cmd,cnt
	cnt = 0
	cmd = "Select classname,class From class Where class is not null and class <> ''"
	rs.open cmd,conn,0,1
	do while NOT rs.EOF
		subclassScript = subclassScript&"asubclass["&cnt&"] = new Array("""&rs("class")&""","""&rs("classname")&""");"
		rs.movenext
		cnt = cnt + 1
	Loop
	if subclassScript <> ""  then
		subclassScript = "var asubclass = new Array();"&subclassScript&"var subclassLen = asubclass.length;"
	end if
	rs.close
	Set rs = nothing
end function

Function showclasses()
	Dim cmd,classes
	cmd = "Select classname From class Where class is null or class=''"
	Set rs = server.createobject("adodb.recordset")
	rs.open cmd,conn,0,1
	do while NOT rs.EOF
		classes = classes&"<option value="&rs("className")&">"&rs("className")&"</option>"
		rs.movenext
	Loop
	showclasses = classes
	rs.close
	Set rs = nothing
end function

Dim action,msg,cmd,classes,subclasses,title,description,gongmu,tclass,subclass
Dim allowExt
allowExt = ".jpg|.jpeg|.gif|.png"
Set rs = server.createobject("adodb.recordset")
action = request.QueryString("action")
if action = "add" then
	checkuser("root,system,user")
	Set upload = new upload_5xsoft
	title = upload.form("title")
	description = upload.form("description")
	gongmu = upload.form("gongmu")
	tclass = upload.form("class")
	subclass = upload.form("subclass")
	if len(title)<1 or len(title)>50 then
		msg = "标题的长度应保持在1-50字之间"
	elseif len(description)<1 or len(description)>200 then
		msg = "描述的长度要保持在1-200字之间"
	elseif gongmu = "" then
		msg = "请选择性别"
	elseif tclass = "" then
		msg = "请选择一级分类"
	elseif subclass = "" then
		msg = "请选择二级分类"
	else
		Set preview = upload.file("preview")
		Dim picSize
		picSize = preview.FileSize
		if picSize < 1 then
			msg = "请上传缩略图片"
		elseif picSize > 51200 then
			msg = "缩略图片大小超过50Kb了"
		else
			Dim filename,ext,dotPos
			filename = preview.FileName
			dotPos = InstrRev(filename,".")
			if dotPos < 1 then
				msg = "上传的文件名不规范"
			else
				ext = LCase(Right(filename,(len(filename) - dotPos)))
				if Instr(allowExt,ext) < 1 then
					msg = "上传的文件的后缀名不符合要求"
				else
					'Save File
					Set pic = upload.file("pic")
					filename = pic.FileName
					picSize = pic.FileSize
					if picSize < 1 then
						msg = "请上传原始图片"
					elseif picSize > 204800 then
						msg = "原始图片大小超过200Kb了"
					else
						dotPos = InstrRev(filename,".")
						if dotPos < 1 then
							msg = "上传的文件名不规范"
						else
							ext = LCase(Right(filename,(len(filename) - dotPos)))
							if Instr(allowExt,ext) < 1 then
								msg = "上传的文件的后缀名不符合要求"
							else
								Dim previewSavefile,picSavefile,folderpath
								folderpath = createFolder()
								previewSavefile = getFilename(folderpath,ext)
								picSavefile = getFilename(folderpath,ext)
								cmd = "Insert Into album(title,username,folder,cover,photo,description,photoscount,commentscount,gongmu,class,subclass)Values("
								cmd = cmd&"'"&title&"','"&session("username")&"','"&folderpath&"','"&folderPath&"/"&previewSavefile&"','"&folderPath&"/"&picSavefile&"','"&description&"',1,0,'"&gongmu&"','"&tclass&"','"&subclass&"')"
								conn.execute cmd
								'Response.Write(folderPath&" -> "&previewSavefile&" -> "&picSavefile)
								'Response.End()
								preview.saveAs Server.MapPath("/album_images/"&folderPath&"/"&previewSavefile)
								pic.saveAs Server.MapPath("/album_images/"&folderPath&"/"&picSavefile)
								msg = "添加成功"
							end if
						end if
					end if
					Set pic = nothing
				end if
			end if
		end if
		Set preview = nothing
	end if
	Set upload = nothing
	Response.Write(msg)
	Response.End()
elseif action = "list" then
	'For verify foster
	if session("role") = "user" then
		cmd = "Select id,title,username,ispass From album Where username='"&session("role")&"' Order by Id desc"
	else
		cmd = "Select id,title,username,ispass From album Order by Id desc"
	end if
	'pagination ?
	Dim list,id,f
	f = request.QueryString("f")
	rs.open cmd,conn,0,1
	do while NOT rs.EOF
		id = rs("id")
		list = list&"<tr><td>"&rs("title")&"</td><td>"&rs("username")&"</td><td>"
		if f = "del" then
			list = list&"<a href=javascript:if(confirm('删除？'))location.href='?action=del&id="&rs("id")&"';>删除</a>"
		elseif f = "uploadpic" then
			list = list&"<a href=?action=showuploadpic&id="&rs("id")&">上传图片</a>"
		elseif f = "verify" then
			if rs("ispass") = true then
				list = list&"<a href=?action=verify&verify=false&id="&rs("id")&">不通过</a>"
			else
				list = list&"<a href=?action=verify&verify=true&id="&rs("id")&">通过</a>"
			end if
		elseif f = "modify" then
			list = list&"<a href=?action=showmodify&id="&rs("id")&">编辑</a>"
		end if
		list = list&"</td>"
		rs.movenext
	Loop
	if list <> "" then
		list = "<table border=0><tr><td>名称</td><td>创建人</td><td>操作</td></tr>"&list&"</table>"
	end if
	Response.Write(list)
	Response.End()
elseif action = "del" then
	checkuser("root,system")
	if request.QueryString("id") <> "" and cint(request.QueryString("id")) > 0 then
		cmd = "Select folder From album Where id="&cint(request.QueryString("id"))
		rs.open cmd,conn,0,1
		Dim folder
		folder = rs("folder")
		delFolder(folder)
		cmd = "Delete From photos Where albumId="&cint(request.QueryString("id"))
		conn.execute cmd
		cmd = "Delete From album Where id="&cint(request.QueryString("id"))
		conn.execute cmd
		Response.Write("删除成功")
	else
		Response.Write("参数传递错误")
	end if
	Response.End()
elseif action = "verify" then
	checkuser("root,system")
	if request.QueryString("id") <> "" and cint(request.QueryString("id")) > 0 then
		Dim verify
		verify = request.QueryString("verify")
		if verify = "true" or verify = "false" then
			cmd = "Update album Set ispass="&verify&" Where id="&cint(request.QueryString("id"))
			conn.execute cmd
			Response.Write("修改成功")
		end if
	else
		Response.Write("参数传递错误")
	end if
	Response.End()
elseif action = "showmodify" then
	checkuser("root,system,user")
	if session("role") = "user" then
		cmd = "Select title,description,gongmu From album Where username='"&session("username")&"' And Id="&cint(request.QueryString("Id"))
	else
		cmd = "Select title,description,gongmu From album Where Id="&cint(request.QueryString("Id"))
	end if
	rs.open cmd,conn,0,1
	if rs.bof and rs.eof and session("role") = "user" then
		Response.Write("普通权限的用户不能修改别人创建的领养信息")
		Response.End()
	else
		title = rs("title")
		description = rs("description")
		gongmu = rs("gongmu")
	End if
elseif action = "modify" then
	title = request.Form("title")
	description = request.Form("description")
	gongmu = request.Form("gongmu")
	tclass = request.Form("class")
	subclass = request.Form("subclass")
	if len(title)<1 or len(title)>50 then
		msg = "标题的长度应保持在1-50字之间"
	elseif len(description)<1 or len(description)>200 then
		msg = "描述的长度要保持在1-200字之间"
	elseif gongmu = "" then
		msg = "请选择性别"
	else
		cmd = "update album set title='"&title&"',description='"&description&"',gongmu='"&gongmu&"'"
		if tclass <> "" and subclass <> "" then
			cmd = cmd&",class='"&tclass&"',subclass='"&subclass&"'"
		elseif tclass <> "" and subclass = "" then
			msg = "请选择二级分类"
		end if
		cmd = cmd&" Where Id="&cint(request.QueryString("Id"))
		conn.execute cmd
		msg = "修改成功"
	end if
elseif action = "uploadpic" then
	checkuser("root,system,user")
	Dim upload,uploadpic,albumId
	if session("role") = "user" then
		cmd = "Select Id,folder From album Where username='"&session("username")&"' And Id="&cint(request.QueryString("Id"))
	else
		cmd = "Select Id,folder From album Where Id="&cint(request.QueryString("Id"))
	end if
	rs.open cmd,conn,0,1
	if rs.bof and rs.eof and session("role") = "user" then
		Response.Write("普通权限的用户不能在别人创建的领养信息中上传图片")
		Response.End()
	else
		folderPath = rs("folder")
		albumId = rs("Id")
	End if
	Set upload = new upload_5xsoft
	uploadpic = upload.form("uploadpic")
	description = upload.form("description")
	if len(description)<1 or len(description)>200 then
		msg = "描述的长度要保持在1-200字之间"
	else
		Set uploadpic = upload.file("uploadpic")
		picSize = uploadpic.FileSize
		if picSize < 1 then
			msg = "请上传图片"
		elseif picSize > 204800 then
			msg = "图片大小超过200Kb了"
		else
			filename = uploadpic.FileName
			dotPos = InstrRev(filename,".")
			if dotPos < 1 then
				msg = "上传的文件名不规范"
			else
				ext = LCase(Right(filename,(len(filename) - dotPos)))
				if Instr(allowExt,ext) < 1 then
					msg = "上传的文件的后缀名不符合要求"
				else
					Dim uploadpicSavefile
					uploadpicSavefile = getFilename(folderPath,ext)
					cmd = "Insert Into photos(username,albumId,filename,description)Values("
					cmd = cmd&"'"&session("username")&"',"&albumId&",'"&folderPath&"/"&uploadpicSavefile&"','"&description&"')"
					conn.execute cmd
					cmd = "Update album Set photoscount=photoscount+1 Where Id="&albumId
					conn.execute cmd
					uploadpic.saveAs Server.MapPath("/album_images/"&folderPath&"/"&uploadpicSavefile)
					msg = "添加成功"
				end if
			end if
		end if
		Set uploadpic = nothing
	end if
	Set upload = nothing
end if
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>领养信息</title>
<link href="/css/standard.css" rel="stylesheet" type="text/css" />
<script language="javascript">
<%=subclassScript()%>
function changeSubclass(tar){
	var tclass = tar.value;
	var subclass = document.getElementById("subclass");
	if (tclass == ""){
		subclass.options[0] = new Option("请选择分类","请选择分类");
	}
	else{
		subclass.length = 0;
		for (i=0;i<subclassLen;i++){
			if (asubclass[i][0] == tclass){
				subclass.options[subclass.length] = new Option(asubclass[i][1],asubclass[i][1]);
			}
		}
	}
}
</script>
<style type="text/css">
<!--
.style1 {color: #FF0000}
-->
</style>
</head>

<body>
<%if request.QueryString("action") = "showuploadpic" or request.QueryString("action") = "uploadpic" then%>
<form action="?action=uploadpic&id=<%=request.QueryString("id")%>" method="post" enctype="multipart/form-data" name="form1" id="form1">
  <table width="80%" border="0">
    <tr>
      <td align="right">图片：</td>
      <td><input name="uploadpic" type="file" id="uploadpic" /></td>
    </tr>
    <tr>
      <td width="20%" align="right" valign="top">描述：</td>
      <td><textarea name="description" cols="50" rows="5" wrap="virtual" id="description"></textarea></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><input type="submit" name="Submit2" value="添加" />
	  <input type="button" name="Submit23333" value="返回" onclick="top.location='main.asp';" />
	  </td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><font color="#FF0000"><%=msg%></font></td>
    </tr>
  </table>
</form>
<%else%>
<%if action = "showmodify" then%>
	<form action="?action=modify&id=<%=request.QueryString("id")%>" method="post" name="form1" id="form1">
<%else%>
	<form action="?action=add" method="post" enctype="multipart/form-data" name="form1" id="form1">
<%end if%>
  <table width="80%" border="0">
    <tr>
      <td colspan="2">注：添加的领养信息需要管理员审核后才能通过。</td>
    </tr>
    <tr>
      <td width="20%" align="right">标题：</td>
      <td><input name="title" type="text" id="title" value="<%=title%>" /></td>
    </tr>
	<%if action <> "showmodify" and action <> "modify" then%>
    <tr>
      <td align="right">缩略图：</td>
      <td><input name="preview" type="file" id="preview" />
(&lt;50Kb，最长边请小于130像素，<span class="style1">推荐90*90大小</span>。)</td>
    </tr>
    <tr>
      <td align="right">原始图片：</td>
      <td><input name="pic" type="file" id="pic" />
        (&lt;200Kb)</td>
    </tr>
	<%end if%>
    <tr>
      <td align="right" valign="top">图片描述：</td>
      <td><textarea name="description" cols="50" rows="5" wrap="virtual" id="description"><%=description%></textarea></td>
    </tr>
    <tr>
      <td align="right">分类：</td>
      <td><select name="gongmu" id="gongmu">
        <option value="公"<%if gongmu = "" or gongmu = "公" then%> selected="selected"<%end if%>>公</option>
        <option value="母"<%if gongmu = "母" then%> selected="selected"<%end if%>>母</option>
      </select>
	  		<%if action = "showmodify" or action = "modify" then%>
			<font color="#FF0000" class="style1">（不选择则不改变分类）</font>
			<%end if%>
        <select name="class" id="class" onChange="changeSubclass(this);">
	  	<option value="" selected="selected">请选择分类</option>
		<%=showclasses()%>
      </select>
        <select name="subclass" id="subclass">
		<option value="">请选择分类</option>
        </select>
	  </td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><input type="submit" name="Submit" value="修改" />
	  <input type="button" name="Submit33332" value="返回" onclick="top.location='main.asp';" />
	  </td>
    </tr>
    <tr>
      <td></td>
      <td><font color="#FF0000"><%=msg%></font></td>
    </tr>
  </table>
</form>
<%end if%>
</body>
</html>
<%
closeRs()
closeConn()
%>