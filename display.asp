<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="inc/utf8.asp"-->
<!--#include file="inc/conn.asp"-->
<%
Dim cmd,id,title,username,photo,description,photoscount,commentscount,gongmu,tclass,subclass,riqi,photos,getLastest,comments,changjianwenda
id = cint(request.QueryString("id"))
if id > 0 then
	if request.QueryString("action") = "addcomment" then
		Dim writer,grade,content,msg
		writer = Replace(trim(request.Form("writer")),"'","")
		grade = cint(request.Form("grade"))
		content = Replace(trim(request.Form("content")),"'","")
		if len(writer)<1 or len(writer)>20 then
			msg = "评论人的长度要保持在1-20字之间。"
		elseif grade < 1 or grade > 5 then
			msg = "评分要保持在1-5分。"
		elseif len(content)<1 or len(content)>200 then
			msg = "评论的长度要保持在1-200字之间。"
		else
			cmd = "Insert Into comments(albumid,writer,grade,content)Values("&id&",'"&writer&"',"&grade&",'"&content&"')"
			conn.execute cmd
			cmd = "Update album Set commentscount=commentscount+1 where id="&id
			conn.execute cmd
		end if
		if msg = "" then
			response.Write("<script>alert('评论成功！');location.href='display.asp?id="&id&"';</script>")
		else
			response.Write("<script>alert('"&msg&"');history.back();</script>")
		end if
		response.End()
	end if
	cmd = "Select title,username,photo,description,photoscount,commentscount,gongmu,class,subclass,riqi,ispass From album Where id="&id
	Set rs = server.createobject("adodb.recordset")
	rs.open cmd,conn,0,1
	if rs.bof and rs.eof then
		Response.Write("参数传递错误。")
		Response.End()
	elseif rs("ispass") = false then
		Response.Write("此领养信息还未通过审核。")
		Response.End()
	else
		title = rs("title")
		username = rs("username")
		photo = rs("photo")
		description = rs("description")
		photoscount = rs("photoscount")
		commentscount = rs("commentscount")
		gongmu = rs("gongmu")
		tclass = rs("class")
		subclass = rs("subclass")
		riqi = rs("riqi")
		rs.close
		cmd = "Select filename,description From photos Where albumId="&id
		rs.open cmd,conn,0,1
		do while NOT rs.EOF
			photos = photos&"<br /><img src=/album_images/"&rs("filename")&" />"
			rs.movenext
		Loop
		rs.close
		cmd = "Select top 5 id,title,photoscount From album Where ispass=true order by id desc"
		rs.open cmd,conn,0,1
		do while NOT rs.EOF
			title = rs("title")
			if len(title) > 9 then
				title = left(title,9)&"..."
			end if
			getLastest = getLastest&"<div><a href=display.asp?id="&rs("id")&">"&title&"("&rs("photoscount")&"张)</a><div>"
			rs.movenext
		Loop
		rs.close
		cmd = "Select top 5 id,title From changjianwenda order by id desc"
		rs.open cmd,conn,0,1
		do while NOT rs.EOF
			title = rs("title")
			if len(title) > 9 then
				title = left(title,9)&"..."
			end if
			changjianwenda = changjianwenda&"<div><a target=_blank href=showmsg.asp?show=changjianwenda&id="&rs("id")&">"&title&"</a><div>"
			rs.movenext
		Loop
		rs.close
		Dim gradesum,gradecount
		cmd = "Select writer,grade,content From comments Where albumId="&id&" Order by id desc"
		rs.open cmd,conn,0,1
		do while NOT rs.EOF
			comments = comments&"<div>"&rs("writer")&"说到：<br />"&rs("content")&"</div><br />"
			gradecount = gradecount + 1
			gradesum = gradesum + cint(rs("grade"))
			rs.movenext
		Loop
		rs.close		
	end if
	closeRs()
	closeConn()
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
<style type="text/css">
<!--
body {
	margin-top: 0px;
}
-->
</style></head>

<body>
<table id="Table_01" width="760" border="0" cellpadding="0" cellspacing="0" align="center">
  <tr>
    <td><img src="images/lingyang4_01.gif" width="760" height="37" alt="" /></td>
  </tr>
  <tr>
    <td><img src="images/lingyang4_02.gif" width="760" height="47" alt="" /></td>
  </tr>
  <tr>
    <td><img src="images/lingyang4_03.gif" width="760" height="59" alt="" /></td>
  </tr>
  <tr>
    <td><img src="images/lingyang4_04.gif" width="760" height="68" alt="" /></td>
  </tr>
</table>
<table width="760" border="0" align="center">
  <tr>
    <td width="150" valign="top">
	<div style="font-size:16px;border:#CCCCCC 1px solid;padding:5px;background-color:#f4f4f4;">最新加入</div>
	<div class="showList"><%=getLastest%></div>
	<br />
	<div style="font-size:16px;border:#CCCCCC 1px solid;padding:5px;background-color:#f4f4f4;">常见问题解答</div>
	<div class="showList"><%=changjianwenda%></div>
	</td>
    <td valign="top" style="padding-left:20px;border-left:#CCCCCC 1px solid;">
	<div><a href="index.asp">首页</a> -&gt; <%=tclass%> -> <%=subclass%> -> <%=title%>（<%=gongmu%>）</div>
	<div>上传者：<%=username%>，日期：<%=riqi%>，照片数：<%=photoscount%>，评论数：<%=commentscount%>，平均得分：<%if gradecount>0 then response.Write(gradesum/gradecount) else response.Write("0")%></div>
	<div style="overflow:hidden;width:590px;"><img src="/album_images/<%=photo%>" /><%=photos%></div>
	<hr />
	<div style="width:590px;overflow:hidden;"><%=comments%></div>
	<table width="80%" border="0">
	<form id="form1" name="form1" method="post" action="?action=addcomment&id=<%=request.QueryString("id")%>">
      <tr>
        <td width="20%" align="right">评论人：</td>
        <td>
          <input name="writer" type="text" id="writer" />        </td>
      </tr>
      <tr>
        <td align="right">评分：</td>
        <td><select name="grade" id="grade">
          <option value="1">1</option>
          <option value="2">2</option>
          <option value="3" selected="selected">3</option>
          <option value="4">4</option>
          <option value="5">5</option>
        </select>        </td>
      </tr>
      <tr>
        <td align="right" valign="top">评论内容：</td>
        <td><textarea name="content" cols="50" rows="5" wrap="virtual" id="content"></textarea></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><input type="submit" name="Submit" value="发表评论" /></td>
      </tr>
</form>
    </table></td>
  </tr>
</table>
<!--#include file="bottom.asp"-->
</body>
</html>