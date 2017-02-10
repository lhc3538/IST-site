<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="../Include/System_Language.asp" -->
<!--#include file="../Include/Version.Asp" -->
<!--#include file="CheckAdmin.Asp"-->
<%
dim Code, Status1, Status2
Call ContactInfo()
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
	If InStr(CompanyAdminPurview, "|9,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style=""color:red; font-size:14px;"")>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If
	%>
    <div class="main">
        <div class="row ">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i> <%=Keyicms_Lang_Copyright(14)%></li>
                <li>企业应用</li>
                <li class="active">在线客服</li>
            </ol>	
        </div>
        <div class="row">
            <div class="col-lg-12 hideInIE8">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-edit"></i><span class="break"></span><strong>在线客服管理</strong></h2> 
                    </div>
                    <div class="panel-body">
                        <a class="ShortCut1" href="Ky_ContactEdit.Asp?Result=Add">
                            <i class="fa fa-edit"></i>
                            <div class="title">添加内置客服</div>
                        </a>
                        <a class="ShortCut2" href="Ky_Contact.Asp">  
                            <i class="fa fa-film"></i>
                            <div class="title">在线客服管理</div>
                        </a>
                    </div>
                </div>
            </div>
        </div>
        <div class="row">
            <div class="col-lg-12">
                <div class="panel panel-default">
                    <div data-original-title="" class="panel-heading">
                        <h2><i class="fa fa-h-square"></i>在线客服</h2>
                    </div>
                    <div class="panel-body">
                    <form name="editForm" method="post" action="Ky_ContactEdit.Asp?Action=SaveEdit&Result=Modify" class="form-horizontal">
                        <!--<div class="panel-heading">
                            <h2>
                            	<i class="fa fa-columns"></i><span class="break"></span>内置客服插件
                                <label class="switch switch-danger">
                                    <input type="checkbox" class="switch-input" name="Status1" value="-1" id="Status1" <%If Status1 Then response.Write "Checked"%>>
                                    <span class="switch-label" data-on="开启" data-off="关闭"></span>
                                    <span class="switch-handle"></span>
                                </label>
                            </h2>
                        </div>-->
                        <div class="form-group">
							<%Call ContactList()%>
                        </div>
 
                        <!--<div class="panel-heading">
                            <h2>
                            	<i class="fa fa-columns"></i><span class="break"></span>其他客服代码
                                <label class="switch switch-danger">
                                    <input type="checkbox" class="switch-input" name="Status2" value="-1" id="Status2" <%If Status1 Then response.Write "Checked"%>>
                                    <span class="switch-label" data-on="开启" data-off="关闭"></span>
                                    <span class="switch-handle"></span>
                                </label>
                            </h2>
                        </div>-->
                        <div class="form-group">
                            <div class="controls">
                                <textarea name="Code" class="form-control" style="width:100%; overflow:hidden; word-wrap:break-word; resize:horizontal; height:134px;"><%=Code%></textarea>
                            </div>
                        </div>

                        <div class="form-group">
                            <div class="col-md-12">
                                <input name="submitSaveEdit" type="submit" id="submitSaveEdit" class="btn btn-danger btn-lg" value="保存客服信息">
                                <input type="button" class="btn btn-danger btn-lg" value="返回上一页" onClick="history.back(-1)">
                            </div>	
                        </div>
                    </form>
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
Function ContactList()
	datafrom = "keyicms_Contact"
	taxis = "Order by Sequence Asc, ID Desc"
	sql = "select * from ["& datafrom &"] where SortID=1 and ID<>1 "&taxis

	set rs = server.createobject("adodb.recordset")
	rs.open sql, conn, 1, 2
	do until rs.eof
		Response.Write "<div class='col-sm-3'>"
		Response.Write "	<div class='panel panel-default'>"
		Response.Write "		<div class='panel-heading'>"
		Response.Write "        	<h2><i class='fa fa-user'></i>客服信息</h2>"
		Response.Write "            <div class='panel-actions'>"
		Response.Write "				<a title='删除' href='JavaScript:void(0)' class='btn-setting' onclick='test("&rs("ID")&", ""Contact"")'><i class='fa fa-times'></i></a>"
		Response.Write "        	</div>"
		Response.Write "		</div>"
		Response.Write "		<div class='panel-body'>"
		Response.Write "        	<div class='form-group'>"
		Response.Write "                <div class='input-group'>"
		Response.Write "                    <span class='input-group-addon'>联系人</span>"
		Response.Write "                    <input type='text' id='User' name='User_"&rs("ID")&"' class='form-control' value='"&rs("User")&"'>"
		Response.Write "                </div>"
		Response.Write "            </div>"
		Response.Write "			<div class='form-group'>"
		Response.Write "                <div class='input-group'>"
		Response.Write "                	<span class='input-group-addon'>QQ</span>"
		Response.Write "                	<input type='text' id='Qq' name='Qq_"&rs("ID")&"' class='form-control' value='"&rs("Qq")&"'>"
		Response.Write "                </div>"
		Response.Write "			</div>"
		Response.Write "        	<div class='form-group'>"
		Response.Write "            	<div class='input-group'>"
		Response.Write "                	<span class='input-group-addon'>手机</span>"
		Response.Write "                	<input type='text' id='Phone' name='Phone_"&rs("ID")&"' class='form-control' value='"&rs("Phone")&"'>"
		Response.Write "            	</div>"
		Response.Write "        	</div>"
		Response.Write "        	<div class='form-group'>"
		Response.Write "                <div class='input-group'>"
		Response.Write "                	<span class='input-group-addon'>微信</span>"
		Response.Write "                	<input type='text' id='WeChat' name='WeChat_"&rs("ID")&"' class='form-control' value='"&rs("WeChat")&"'>"
		Response.Write "                </div>"
		Response.Write "        	</div>"
		Response.Write "        	<div class='form-group'>"
		Response.Write "                <div class='input-group'>"
		Response.Write "                	<span class='input-group-addon'>排序</span>"
		Response.Write "                	<input type='text' id='Sequence' name='Sequence_"&rs("ID")&"' class='form-control' value='"&rs("Sequence")&"' onKeyDown='OnKeyDown()' onchange=""OnChange(this, "&rs("Sequence")&")"">"
		Response.Write "                </div>"
		Response.Write "        	</div>"
		Response.Write "        	<input name='ID' type='hidden' value='"&rs("ID")&"'>"
		Response.Write "    	</div>"
		Response.Write "	</div>"
		Response.Write "</div>"
		rs.movenext
	loop
	rs.close 
	set rs = nothing
End Function

Function ContactInfo()
	set rs = server.createobject("adodb.recordset")
	sql = "select top 1 * from keyicms_Contact where SortID=0"
	rs.open sql, conn, 1, 2
	'Status1 = rs("Status")
	rs.close 
	set rs = nothing
	
	set rs = server.createobject("adodb.recordset")
	sql = "select top 1 * from keyicms_Contact where SortID=2"
	rs.open sql, conn, 1, 2
	'Code = rs("Code")
	'Status2 = rs("Status")
	rs.close 
	set rs = nothing
End Function
%>