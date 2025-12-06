<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="inc/utf8.asp"-->
<!--#include file="inc/conn.asp"-->
<%
Function showclasses()
	Dim cmd,classes
	cmd = "Select classname From class Where class is null or class=''"
	Set rs = server.createobject("adodb.recordset")
	rs.open cmd,conn,0,1
	do while NOT rs.EOF
		classes = classes&"<input name=""class"" type=""checkbox"" id=""class"" value="""&rs("classname")&""" />"&rs("classname")
		rs.movenext
	Loop
	showclasses = classes
	rs.close
	Set rs = nothing
end function

Function gotopage(pagenum)
	Dim url,urla,length
	url = request.ServerVariables("QUERY_STRING")
	if instr(url,"page=") > 0 then
		urla = split(url,"&")
		url = ""
		length = ubound(urla)
		for i=0 to length
			if instr(urla(i),"page=") > 0 then
				url = url & "&page=" & pagenum
			else
				url = url & "&" & urla(i)
			end if
		next
	else
		url = "&" & url & "&page=" & pagenum
	end if
	do while left(url,1) = "&"
		url = right(url,len(url) - 1)
	loop
	gotopage = "?" & url
end Function

Dim affiche,rs,cmd,changjianwenda,currentpage,pagecount
cmd = "SELECT top 5 id,title FROM affiche ORDER BY ID DESC"
Set rs = server.createobject("adodb.recordset")
rs.open cmd,conn,0,1
if rs.eof and rs.bof then
	affiche = ""
else
	affiche = "<MARQUEE onmouseover=this.stop() onmouseout=this.start() scrollAmount=2 scrollDelay=5>"
	do while not rs.eof
	affiche = affiche &"<a target=_blank href=showmsg.asp?id="&rs("id")&">"& rs("title") & "</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
	rs.movenext
	loop
	affiche = affiche & "</MARQUEE>"
end if
'closeRs()
Dim lingyang,showcount,chooseclasses,where
chooseclasses = trim(request.QueryString("class"))
if chooseclasses <> "" then
	classA = split(chooseclasses,",")
	Dim length
	length = Ubound(classA)
	for i=0 to length
		where = where&" or class='"&trim(classA(i))&"'"
	next
	where = " and ("&right(where,len(where)-4)&")"
end if
if request.QueryString("gongmu") <> "" and request.QueryString("gongmu") <> "全部" then
	where = where & " and gongmu='"&request.QueryString("gongmu")&"'"
end if
if request.QueryString("title") <> "" then
	where = where & " and title like '%"&Replace(request.QueryString("title"),"'","")&"%'"
end if
cmd = "SELECT id,title,cover,photoscount FROM album WHERE ispass=true"&where&" ORDER BY ID DESC"
Set rs = server.createobject("adodb.recordset")
rs.open cmd,conn,1,1
rs.pagesize = 9
if rs.bof and rs.eof then
	lingyang = "暂无"
else
	pagecount = rs.pagecount
	Dim cnt
	if request.QueryString("page") = "" then
		currentpage = 0
	else
		currentpage = cint(request.QueryString("page"))
	end if
	if currentpage < 1 then
		currentpage = 1
	elseif currentpage > pagecount then
		currentpage = pagecount
	end if
	rs.absolutepage = currentpage
	cnt = 0
	do while cnt < 9 and NOT rs.EOF
		lingyang = lingyang & "<td align=center style=""width:33%;padding:5px;border:#CCCCCC 1px solid;""><a target=_blank href=display.asp?id="&rs("id")&"><img src=""/album_images/"&rs("cover")&""" border=0 /></a><br />"&rs("title")&" | "&rs("photoscount")&"张</td>"
		cnt = cnt + 1
		if cnt mod 3 = 0 and cnt < 9 then lingyang = lingyang & "</tr><tr>"
		rs.movenext
	loop
	if lingyang <> "" then
		lingyang = "<table width=100% border=0 cellpadding=0 cellspacing=5><tr>" & lingyang & "</tr></table>"
	end if
	'Response.Write(lingyang)
	'Response.End()
end if
rs.close
Dim title
cmd = "Select top 7 id,title From changjianwenda order by id desc"
rs.open cmd,conn,0,1
do while NOT rs.EOF
	title = rs("title")
	if len(title) > 12 then
		title = left(title,12)&"..."
	end if
	changjianwenda = changjianwenda&"<div><a target=_blank href=showmsg.asp?show=changjianwenda&id="&rs("id")&">"&title&"</a><div>"
	rs.movenext
Loop
rs.close
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Songday.com 领养中心</title>
<link href="/css/standard.css" rel="stylesheet" type="text/css" />
<SCRIPT language=JavaScript type=text/JavaScript>
<!--//
function showDate(){
	var yy=new Date();
	var gyy=yy.getYear();
	var ddName=new Array("星期日","星期一","星期二","星期三","星期四","星期五","星期六");
	var mmName=new Array("1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月");
	document.write(yy.getFullYear() +"年"+ mmName[yy.getMonth()] + yy.getDate() + "日&nbsp" + ddName[yy.getDay()]);
}
function gongmu(gm){
	var tar = document.getElementById("gongmu");
	if (gm == "gong")
		tar.options[1].selected = true;
	else if (gm == "mu")
		tar.options[2].selected = true;
	//alert(tar.value);
	document.form1.submit();
}
function gotopage(){
	var tar = document.getElementById("gotopage");
	var url = window.location.search==null?"":window.location.search.replace("?","");
	if (url.indexOf("page=")>-1){
		var searcha = url.split("&");
		var len = searcha.length;
		url = "";
		for (i=0;i<len;i++){
			//alert(searcha[i]);
			if (searcha[i].indexOf("page=") > -1)
				url += "&page=" + tar.value;
			else
				url += "&" + searcha[i];
		}
	}
	else
		url += "&page=" + tar.value;
	location.href = "index.asp?" + url.substring(1);
}
//-->
</SCRIPT>
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
    <td width="30%" align="center"><script>showDate();</script></td>
    <td align="center"><%=affiche%></td>
  </tr>
</table>
<table width="760" border="0" align="center">
  <tr>
    <td width="30%" valign="top"><table width="85%" border="0" align="center" bgcolor="#F4F4F4" style="border:#CCCCCC 1px solid;">
	<form id="form1" name="form1" method="GET" action="index.asp?page=<%=request.QueryString("page")%>">
      <tr>
        <td><input name="title" type="text" id="title" style="height:18px;width:70px;border:#CCCCCC 1px solid;background-color:#ffffff;" />
        &nbsp;&nbsp;
<select name="gongmu" id="gongmu" style="height:18px;border:#FFFFFF 1px solid;">
              <option selected="selected" value="">全部</option>
              <option value="公">公</option>
              <option value="母">母</option>
            </select></td>
        <td><input type="image" name="imageField" src="images/fangdajing.gif" /></td>
      </tr>
      <tr>
        <td colspan="2"><%=showclasses%></td>
        </tr></form>
    </table>
		<table width="85%" border="0" align="center">
          <tr>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td style="font-size:16px;border:#CCCCCC 1px solid;padding:5px;background-color:#f4f4f4;">常见问题解答</td>
          </tr>
          <tr>
            <td class="showList"><%=changjianwenda%></td>
          </tr>
      </table>
      <table width="95%" border="0" align="center">
        <tr></tr>
      </table>
      <table width="85%" border="0" align="center">
        <tr>
          <td><img src="images/hua.gif" width="152" height="131" /></td>
        </tr>
      </table></td>
    <td valign="top" style="border-left:#CCCCCC 1px solid;"><table width="95%" border="0" align="center" style="border-top:#CCCCCC 1px solid;">
        <tr>
          <td colspan="2"><img src="images/wangzigongzhu.gif" width="109" height="28" border="0" usemap="#Map" />
		  <%if where <> "" then%><a href="index.asp">【查看全部】</a><%end if%>
		  </td>
        </tr>
        <tr>
          <td colspan="2">
		  <%=lingyang%>
		  </td>
        </tr>
        <tr>
          <td width="50%">
				<%
				if currentpage<=1 then
				Response.Write"第一页 | 上一页 | "
				else
				Response.Write"<a href='index.asp"&gotopage(1)&"'>第一页</a> | <a href='index.asp"&gotopage(currentpage-1)&"'>上一页</a> | "
				end if
				if currentpage<pagecount then
				Response.Write"<a href='index.asp"&gotopage(currentpage+1)&"'>下一页</a> | <a href='index.asp"&gotopage(pagecount)&"'>最后一页</a>"
				else
				Response.Write"下一页 | 最后一页"
				end if
				%>
		  </td>
          <td align="right">页数：
            <input name="gotopage" type="text" id="gotopage" size="5" />
            <input name="gotopageButton" type="button" id="gotopageButton" value="转到" onclick="gotopage();" />
            ，共<%=pagecount%>页</td>
        </tr>
      </table></td>
  </tr>
</table>
<!--#include file="bottom.asp"-->

<map name="Map" id="Map"><area shape="rect" coords="64,-16,111,32" href="javascript:gongmu('mu');" />
<area shape="rect" coords="-7,-11,45,34" href="javascript:gongmu('gong');" />
</map></body>
</html>
<%
closeRs()
closeConn()
%>