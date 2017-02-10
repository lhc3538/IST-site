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
	If InStr(CompanyAdminPurview, "|7,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style=""color:red; font-size:14px;"")>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If
	%>
    <div class="main">
        <div class="row">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i> <%=Keyicms_Lang_Copyright(14)%></li>
                <li><%=Keyicms_Lang_Copyright(2)%></li>
                <li class="active">功能模块管理</li>
            </ol>	
        </div>	
        <div class="row">
            <div class="col-lg-12 hideInIE8">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-edit"></i><span class="break"></span><strong>功能模块</strong></h2> 
                    </div>
                    <div class="panel-body">
                        <a class="ShortCut1" href="Ky_System.Asp">  
                            <i class="fa fa-list"></i>
                            <div class="title">功能模块列表</div>
                        </a>					
                    </div>
                </div>
            </div>
        </div> 
        <div class="row">
        	<div class="col-lg-12 hideInIE8">
            	<div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-user"></i><span class="break"></span>功能模块管理</h2>
                    </div>
                    <div class="panel-body">
                        <%SystemList()%>
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
Function SystemList()
datafrom = "keyicms_System"
set rs = server.CreateObject("adodb.recordset")
rs.open "select * from "&datafrom&" Order by ID asc",conn,1,2
do until rs.eof
	response.Write "<div class='AppDiv'>"
	response.Write "	<div class='Div1'>"
	response.Write "		<img src='"&rs("Pic")&"'/>"
	response.Write "		<div class='text'></div>"
	response.Write "	</div>"
	response.Write "	<h2 class='text-center'><strong>"&rs("SystemName")&"</strong></h2>"
	response.Write "	<div class='row'>"
	response.Write "		<div class='col-xs-5'>"
	response.Write "			<a title='管理' href='"&rs("SystemPath")&"' class='btn btn-primary btn-block'>管理</a>"
	response.Write "		</div>"
	response.Write "		<div class='col-xs-5'>"
	response.Write "			<a title='修改' href='Ky_SystemEdit.Asp?Result=Modify&ID="&rs("ID")&"' class='btn btn-primary btn-block'>修改</a>"
	response.Write "		</div>"
	response.Write "	</div>"
	response.Write "</div>"
	rs.movenext
loop
rs.close : set rs = nothing
End Function
%>