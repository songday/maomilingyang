<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Menu</title>
<link href="/css/standard.css" rel="stylesheet" type="text/css" />
</head>

<body>
<fieldset>
<legend>管理员</legend>
<a href="user.asp" target="mainFrame">添加</a><br />
<a href="user.asp?f=update&action=showadmin" target="mainFrame">编辑</a><br />
<a href="user.asp?f=del&action=showadmin" target="mainFrame">删除</a>
</fieldset>
<fieldset>
<legend>公告</legend>
<a href="affiche.asp" target="mainFrame">添加</a><br />
<a href="affiche.asp?f=showupdate&action=showlist" target="mainFrame">编辑</a><br />
<a href="affiche.asp?f=del&action=showlist" target="mainFrame">删除</a>
</fieldset>
<fieldset>
<legend>常见问答</legend>
<a href="affiche.asp?target=changjianwenda" target="mainFrame">添加</a><br />
<a href="affiche.asp?target=changjianwenda&f=showupdate&action=showlist" target="mainFrame">编辑</a><br />
<a href="affiche.asp?target=changjianwenda&f=del&action=showlist" target="mainFrame">删除</a>
</fieldset>
<fieldset>
<legend>分类</legend>
<a href="type.asp" target="mainFrame">管理</a><br />
</fieldset>
<fieldset>
<legend>领养信息</legend>
<a href="foster.asp" target="mainFrame">发布</a><br />
<a href="foster.asp?f=uploadpic&action=list" target="mainFrame">上传图片</a><br />
<a href="foster.asp?f=modify&action=list" target="mainFrame">编辑</a><br />
<a href="foster.asp?f=verify&action=list" target="mainFrame">审核</a><br />
<a href="foster.asp?f=del&action=list" target="mainFrame">删除</a>
</fieldset>
<fieldset>
<legend>其他</legend>
<a href="welcome.asp" target="mainFrame">信息</a><br />
<a href="logout.asp" target="_top">退出</a>
</fieldset>
</body>
</html>