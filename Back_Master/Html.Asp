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
'神笔(33)天成网络享有软件著作权，如果您擅自删除这里的版权信息,我们将追究法律责任.
'未经公司同意禁止更改以下代码.
'*******************************************************
Call SiteInfo()
%>
<!doctype html>
<html>
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
<div id="wrap">
	<!--#Include file="Keyicms_Meun.Asp"-->
    <%
	If InStr(CompanyAdminPurview, "|20,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style='color:red; font-size:14px;')>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If
	%>
    <div class="main">
        <div class="row">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i> <%=Keyicms_Lang_Copyright(14)%></li>
                <li><%=Keyicms_Lang_Copyright(2)%></li>
                <li class='active'>生成静态</li>
            </ol>	
        </div>	
        <div class="row">
            <div class="col-lg-12 hideInIE8">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-edit"></i><span class="break"></span>生成网站静态页面（PC）</h2> 
                    </div>
                    <div class="panel-body">
                        <%Call Html("PC")%>
                        <div class="btn-group"  style="margin:10px 10px 10px 10px;">             
                            <a target="html_Bar" type="button" class="btn btn-primary" href="Html_Index.Asp?Ke01=PC"><i class="fa fa-home"></i> 生成首页、联系及留言</a>
                        </div>
                        <div class="btn-group"  style="margin:10px 10px 10px 10px;">             
                            <a target="html_Bar" type="button" class="btn btn-danger" href="JavaScript:void(0)" onClick="JavaScript:HtmlSite('PC')"><i class="fa fa-home"></i> 生成PC全站静态</a>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        <%If mStatus Then%>
        <div class="row">
            <div class="col-lg-12 hideInIE8">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-edit"></i><span class="break"></span>生成网站静态页面（手机）</h2> 
                    </div>
                    <div class="panel-body">
						<%Call Html("Mobile")%>
                        <div class="btn-group"  style="margin:10px 10px 10px 10px;">             
                            <a target="html_Bar" type="button" class="btn btn-primary" href="Html_Index.Asp?Ke01=Mobile"><i class="fa fa-home"></i> 生成首页、联系及留言</a>
                        </div>
                        <div class="btn-group"  style="margin:10px 10px 10px 10px;">             
                            <a target="html_Bar" type="button" class="btn btn-danger" href="JavaScript:void(0)" onClick="JavaScript:HtmlSite('Mobile')"><i class="fa fa-home"></i> 生成手机全站静态</a>
                        </div>
                    </div>
                </div>
            </div>
        </div>
		<%End If%>
        <iframe style="border:0; width:100%;" scrolling="no" src="" name="html_Bar" id="html_Bar"></iframe>
    </div>
	<!--#Include file="Keyicms_End.Asp"-->
</div>
<!--#Include file="Keyicms_js.Asp"-->
</body>
</html>
<%
Function Html(Ke01)
set rs = server.CreateObject("adodb.recordset")
rs.open "select * from keyicms_System where ViewFlag Order by ID asc", conn, 1, 2
do until rs.eof
	If rs("DataForm") = 1 Then
		html1 = "Html_AboutSort.Asp?Ke01="&Ke01
		html2 = "Html_About.Asp?Ke01="&Ke01
	End If
	If rs("DataForm") = 2 Then
		html1 = "Html_NewsSort.Asp?Ke01="&Ke01
		html2 = "Html_News.Asp?Ke01="&Ke01
	End If
	If rs("DataForm") = 3 Then
		html1 = "Html_ProductSort.Asp?Ke01="&Ke01
		html2 = "Html_Product.Asp?Ke01="&Ke01
	End If
	If rs("DataForm") = 4 Then
		html1 = "Html_CaseSort.Asp?Ke01="&Ke01
		html2 = "Html_Case.Asp?Ke01="&Ke01
	End If
	If rs("DataForm") = 5 Then
		html1 = "Html_DownSort.Asp?Ke01="&Ke01
		html2 = "Html_Download.Asp?Ke01="&Ke01
	End If
	If rs("DataForm") = 6 Then
		html1 = "Html_JobSort.Asp?Ke01="&Ke01
		html2 = "Html_Job.Asp?Ke01="&Ke01
	End If
	response.Write "<div class='btn-group' style='margin:10px 10px 10px 10px;'>"         
	response.Write "	<button type='button' class='btn btn-success'>"&rs("SystemName")&"静态管理</button>"
	response.Write "	<button type='button' class='btn btn-success dropdown-toggle' data-toggle='dropdown'>"
	response.Write "		<span class='caret'></span>"  
	response.Write "		<span class='sr-only'>"&rs("SystemName")&"静态管理</span>"  
	response.Write "	</button>"
	response.Write "	<ul class='dropdown-menu' role='menu'>"
	If rs("DataForm") = 1 Then
		response.Write "	<li><a target='html_Bar' href='"&html1&"'>生成"&rs("SystemName")&"栏目页</a></li>"
	Else
		response.Write "	<li><a target='html_Bar' href='"&html1&"'>生成"&rs("SystemName")&"栏目及分类页</a></li>"
	End If
	response.Write "	<li><a target='html_Bar' href='"&html2&"'>生成"&rs("SystemName")&"详细页</a></li>"
	response.Write "	</ul>"
	response.Write "</div>"  
	rs.movenext
loop
rs.close : set rs = nothing
End Function
%>