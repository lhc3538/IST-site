<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="../Include/System_Language.asp" -->
<!--#include file="../Include/Version.Asp" -->
<!--#include file="CheckAdmin.Asp"-->
<%
'*******************************************************
'软件名称：科蚁企业网站内容管理系统（KEYICMS）
'软件开发：成都神笔天成网络科技有限公司
'网   址：http://www.keyicms.com
'本信息不会影响您网站的正常使用，无论免费用户或是收费用户请保留这里的信息.
'神笔(24)天成网络享有软件著作权，如果您擅自删除这里的版权信息,我们将追究法律责任.
'未经公司同意禁止更改以下代码.
'*******************************************************
If InStr(CompanyAdminPurview, "|21,") = 0 Then
	response.Write ("<br /><br /><div align='center'><font style=""color:red; font-size:9pt; "")>您没有管理该模块的权限！</font></div>")
	response.End
End If
%>
<!DOCTYPE html>
<html lang="zh-cn">
<head>
<meta charset="utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="description" content="<%=Keyicms_Lang_Copyright(13)%>">
<meta name="author" content="<%=Keyicms_Lang_Copyright(3)%>">
<title><%=Keyicms_Lang_Copyright(1)%> <%=Str_Soft_Version%></title>
<!--#Include file="Keyicms_css.Asp"-->
</head>

<body>
<!--#Include file="Keyicms_Top.Asp"-->
<div id="wrap">
	<!--#Include file="Keyicms_Meun.Asp"-->
	<%
	If InStr(CompanyAdminPurview, "|14,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style=""color:red; font-size:14px;"")>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If
	%>
    <div class="main">
        <div class="row">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i> <%=Keyicms_Lang_Copyright(14)%></li>
                <li><%=Keyicms_Lang_Copyright(2)%></li>
                <li class="active">企业应用中心</li>
            </ol>	
        </div>	
        <div class="row">
            <div class="col-lg-12 hideInIE8">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-edit"></i><span class="break"></span><strong>应用管理</strong></h2> 
                    </div>
                    <div class="panel-body">
                        <a class="ShortCut1" href="Ky_App.Asp">  
                            <i class="fa fa-suitcase"></i>
                            <div class="title">应用管理</div>
                        </a>	
                    </div>
                </div>
            </div>
        </div>
        <div class="row">
        	<div class="col-lg-12 hideInIE8">
            	<div class="panel panel-default">
                    <div class="panel-heading">
                        <ul class="nav tab-menu nav-tabs" id="myTab">
                            <li class=""><a href="#PC"><%If mStatus Then%>PC<%End If%>应用管理</a></li>
                            <%If mStatus Then%><li class=""><a href="#Mobile">手机应用管理</a></li><%End If%>
                        </ul>
                    </div>
                    <div class="panel-body">
                        <div id="myTabContent" class="tab-content">
                            <div class="tab-pane" id="PC"><%AppList("PC")%></div>
                            <%If mStatus Then%><div class="tab-pane" id="Mobile"><%AppList("Mobile")%></div><%End If%>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <!--#Include file="Keyicms_End.Asp"-->
</div>
<!--#Include file="Keyicms_js.Asp"-->
</body>
</html>
<%
Function AppList(SortName)
datafrom = "keyicms_App"
set rs = server.CreateObject("adodb.recordset")
rs.open "select * from "&datafrom&" where SortName='"&SortName&"' Order by ID desc",conn,1,2
if rs.eof then
	response.Write "<center>暂无应用</center>"
else
	do until rs.eof
		response.Write "<div class='AppDiv'>"
		response.Write "	<div class='Div1'>"
		response.Write "		<img src='"&rs("SmallPic")&"'/>"
		response.Write "		<div class='text'>"&rs("Text")&"</div>"
		response.Write "	</div>"
		response.Write "	<h2 class='text-center'><strong>"&rs("AppName")&"</strong></h2>"
		response.Write "	<div class='row'>"
		if rs("ViewFlag") then
			response.Write "	<div class='col-xs-10'>"
			response.Write "		<a title='管理' href='"&rs("AppPath")&"' class='btn btn-primary btn-block'>管理</a>"
			response.Write "	</div>"
		else
			If rs("OutFlag") Then
				response.Write "    <div class='col-xs-4'>"
				response.Write "        <a title='启用' href='Conversion.Asp?id="&rs("ID")&"&LX="&datafrom&"&Conver=ViewFlag&Operation=up' class='btn btn-success btn-block'>启用</a>"
				response.Write "    </div>"
				response.Write "    <div class='col-xs-4'>"
				response.Write "        <a title='修改' href='Ky_AppAdd.Asp?Result=Modify&id="&rs("ID")&"' class='btn btn-success btn-block'>修改</a>"
				response.Write "    </div>"
				response.Write "    <div class='col-xs-4'>"
				response.Write "        <a title='卸载' href='JavaScript:void(0)' onClick='test_1();' class='btn btn-success btn-block'>卸载</a>"
				response.Write "    </div>"
			Else
				response.Write "    <div class='col-xs-10'>"
				response.Write "        <a title='启用' href='Conversion.Asp?id="&rs("ID")&"&LX="&datafrom&"&Conver=ViewFlag&Operation=up' class='btn btn-success btn-block'>启用</a>"
				response.Write "    </div>"
			End If
		end if
		response.Write "	</div>"
		response.Write "</div>"
		rs.movenext
	loop
end if
rs.close
set rs = nothing
End Function
%>