<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="../Include/system_language.asp" -->
<!--#include file="../Include/Version.Asp" -->
<!--#include file="CheckAdmin.Asp"-->
<%
'*******************************************************
'软件名称：科蚁企业网站内容管理系统（KEYICMS）
'软件开发：成都神笔天成网络科技有限公司
'网   址：http://www.keyicms.com
'本信息不会影响您网站的正常使用，无论免费用户或是收费用户请保留这里的信息.
'神笔(31)天成网络享有软件著作权，如果您擅自删除这里的版权信息,我们将追究法律责任.
'未经公司同意禁止更改以下代码.
'*******************************************************
xmlErr = "<h2 class='text-center'>keyicms_template.xml文件中文件有错误</h2>"

Result = request.QueryString("Result")
If Result = "Add" Then Call SaveEdit()
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
<link rel="stylesheet" href="KEditor_keyicms/themes/default/default.css" />
<script charset="utf-8" src="KEditor_keyicms/kindeditor-min.js"></script>
<script charset="utf-8" src="KEditor_keyicms/lang/zh_CN.js"></script>
<script charset="utf-8" src="KEditor_keyicms/plugins/code/prettify.js"></script>
<script>
KindEditor.ready(function(K) {
	var editor = K.editor({
		cssPath : 'KEditor_keyicms/plugins/code/prettify.css',
		uploadJson : 'KEditor_keyicms/asp/upload_json.asp',
		fileManagerJson : 'KEditor_keyicms/asp/file_manager_json.asp',
		allowFileManager : true
	});
	K('#image').click(function() {
		editor.loadPlugin('image', function() {
			editor.plugin.imageDialog({
				imageUrl : K('#SmallPic').val(),
				clickFn : function(url, title, width, height, border, align) {
					K('#SmallPic').val(url);
					editor.hideDialog();
				}
			});
		});
	});
})
</script>
</head>
<body>
<!--#Include file="Keyicms_Top.Asp"-->
<div id="wrap">
	<!--#Include file="Keyicms_Meun.Asp"-->
	<%
	If InStr(CompanyAdminPurview, "|19,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style=""color:red; font-size:14px;"")>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If
	%>
    <div class="main">
        <div class="row">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i> <%=Keyicms_Lang_Copyright(14)%></li>
                <li><%=Keyicms_Lang_Copyright(2)%></li>
                <li class="active"><a href="">模板管理</a></li>
                <li class="active"><%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改<%End If%>模板</li>
            </ol>	
        </div>	
        <div class="row">
            <div class="col-lg-12 hideInIE8">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-edit"></i><span class="break"></span><strong>模板管理</strong></h2> 
                    </div>
                    <div class="panel-body">
                            <a class="ShortCut2" href="Ky_Template.Asp">  
                                <i class="fa fa-desktop"></i>
                                <div class="title">模板管理</div>
                            </a>	
                            <a class="ShortCut4" href="http://www.keyicms.com" target="_blank">  
                                <i class="fa fa-columns"></i>
                                <div class="title">模板市场</div>
                            </a>													
                        </div>
					</div>
				</div>
			</div>
			<div class="row">
            <div class="col-md-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <ul class="nav tab-menu nav-tabs" id="myTab">
                            <li><a href="#PC"><%If mStatus Then%>PC<%End If%>网站模板</a></li>
                            <%If mStatus Then%><li><a href="#Mobile">手机网站模板</a></li><%End If%>
                        </ul>
                    </div>
                    <div class="panel-body">
                        <div id="myTabContent" class="tab-content">
                            <div class="tab-pane" id="PC"><%TemplateList("PC")%></div>
                            <%If mStatus Then%><div class="tab-pane" id="Mobile"><%TemplateList("Mobile")%></div><%End If%>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
	<!--#Include file="Keyicms_End.Asp"-->				
</div>
<!--#Include file="Keyicms_js.Asp"-->
<script>
$("select[name=SortName]").change(function(){
	SortName = $(this).children("option:selected").attr("value")
	$("#TemplatePathText").html("/Template/" + SortName + "/")
})
</script>
</body>
</html>
<%
Function TemplateList(Ke01)
set rs = server.CreateObject("adodb.recordset")
rs.open "select * from keyicms_Template where SortName='"&Ke01&"' Order by UseFlag asc",conn,1,2
if rs.eof then
	response.Write "<center>暂无模板</center>"
else
	do until rs.eof
		response.Write "<div class='TemplateDiv'>"
		response.Write "	<div class='Div1'>"
		response.Write "		<img src='"&rs("SmallPic")&"'/>"
		response.Write "		<div class='text'>"
		response.Write "			<p>作者：<a href='"&rs("WebSite")&"' target='_blank'>"&rs("Author")&"</a></p>"
		response.Write "			<p>版本："&rs("Version")&"</p>"
		response.Write "			<p>目录："&rs("TemplatePath_A")&"</p>"
		response.Write "			<p>简介："&rs("Text")&"</p>"
		response.Write "		</div>"
		response.Write "	</div>"
		response.Write "	<h2 class='text-center'><strong>"&rs("TemplateName")&"</strong></h2>"
		response.Write "	<div class='row'>"
		if not rs("UseFlag") then
			response.Write "    <div class='col-xs-5'>"
			response.Write "        <a href='TemplateUsed.Asp?SortName="&rs("SortName")&"&ID="&rs("ID")&"' class='btn btn-success btn-block'>启用</a>"
			response.Write "    </div>"
			response.Write "    <div class='col-xs-5'>"
			Response.Write "		<a href='JavaScript:void(0)' class='btn btn-success btn-block' onclick='test_1("&rs("ID")&", ""Template"")'>卸载</a>" & vbCrLf
			response.Write "    </div>"
		else
			response.Write "	<div class='col-xs-11'>"
			response.Write "		<a class='btn btn-danger btn-block'>已启用</a>"
			response.Write "	</div>"
		end if
		response.Write "	</div>"
		response.Write "</div>"
		rs.movenext
	loop
end if
rs.close : set rs = nothing

Set FSO = CreateObject("Scripting.FileSystemObject")
Set Folder = FSO.GetFolder(Server.MapPath("/Template/"&Ke01))
For Each Folder_Item in Folder.subfolders
	FolderName = Folder_Item.name
	FolderPath = "/Template/"&Ke01&"/"&FolderName
	
	If Conn.execute("select * from keyicms_template where TemplatePath_A='"&FolderName&"'").eof Then
		response.Write "<div class='TemplateDiv'>"
		If CheckFile(Server.MapPath(FolderPath&"/keyicms_template.xml")) Then
			Set xml = Server.CreateObject("Microsoft.XMLDOM")  
			xml.Async = False
			xml.load(Server.MapPath(FolderPath&"/keyicms_template.xml"))
			set tem = xml.selectSingleNode("template")
			If not tem Is Nothing Then 
				If not tem.selectSingleNode("name") Is Nothing and _
				not tem.selectSingleNode("preview") Is Nothing and _
				not tem.selectSingleNode("name") Is Nothing and _
				not tem.selectSingleNode("author") Is Nothing and _
				not tem.selectSingleNode("website") Is Nothing and _
				not tem.selectSingleNode("version") Is Nothing and _
				not tem.selectSingleNode("text") Is Nothing Then
					xml_preview     = "/Template/"&Ke01&"/"&FolderName&"/"&xml.selectSingleNode("template").selectSingleNode("preview").text
					xml_name        = xml.selectSingleNode("template").selectSingleNode("name").text
					xml_author      = xml.selectSingleNode("template").selectSingleNode("author").text
					xml_website     = xml.selectSingleNode("template").selectSingleNode("website").text
					xml_version     = xml.selectSingleNode("template").selectSingleNode("version").text
					xml_text        = xml.selectSingleNode("template").selectSingleNode("text").text
					
					response.Write "	<div class='Div1'>"
					response.Write "		<img src='"&xml_preview&"'/>"
					response.Write "		<div class='text'>"
					response.Write "			<p>作者：<a href='"&xml_website&"' target='_blank'>"&xml_author&"</a></p>"
					response.Write "			<p>版本："&xml_version&"</p>"
					response.Write "			<p>目录："&FolderName&"</p>"
					response.Write "			<p>简介："&xml_text&"</p>"
					response.Write "		</div>"
					response.Write "	</div>"
					response.Write "	<h2 class='text-center'><strong>"&xml_name&"</strong></h2>"
					response.Write "	<div class='row'>"
					response.Write "		<div class='col-xs-11'>"
					response.Write "			<a href='?Result=Add&Ke01="&Ke01&"&FolderName="&FolderName&"' class='btn btn-info btn-block'>开始安装</a>"
					response.Write "		</div>"
					response.Write "	</div>"
				Else
					response.Write xmlErr
				End If
			Else
				response.Write xmlErr
			End If
			set xml = nothing
		Else
			response.Write "<h2 class='text-center'>"&FolderName&"文件夹下未找到</h2>"
			response.Write "<h2 class='text-center'>keyicms_template.xml文件</h2>"
		End If
		response.Write "</div>"
	End If
Next
Set Folder = Nothing
Set FSO = Nothing
End Function

Function SaveEdit()
	Ke01 = Request.QueryString("Ke01")
	FolderName = Request.QueryString("FolderName")
	FolderPath = "/Template/"&Ke01&"/"&FolderName
	
	Set xml = Server.CreateObject("Microsoft.XMLDOM")  
	xml.Async = False
	xml.load(Server.MapPath(FolderPath&"/keyicms_template.xml"))
	xml_preview     = "/Template/"&Ke01&"/"&FolderName&"/"&xml.selectSingleNode("template").selectSingleNode("preview").text
	xml_name        = xml.selectSingleNode("template").selectSingleNode("name").text
	xml_author      = xml.selectSingleNode("template").selectSingleNode("author").text
	xml_website     = xml.selectSingleNode("template").selectSingleNode("website").text
	xml_version     = xml.selectSingleNode("template").selectSingleNode("version").text
	xml_text        = xml.selectSingleNode("template").selectSingleNode("text").text
	
    set rs = server.createobject("adodb.recordset")
	sql = "select * from keyicms_Template"
	
	rs.open sql,conn,1,3
	rs.addnew
	rs("TemplateName") = trim(xml_name)
	rs("SortName") = trim(Ke01)
	rs("TemplatePath") =  trim(FolderPath)
	rs("TemplatePath_A") = trim(FolderName)
	rs("SmallPic") = trim(xml_preview)
	rs("Author") = trim(xml_author)
	rs("Text") = trim(xml_text)
	rs("Version") = trim(xml_version)
	rs("Website") = trim(xml_website)
	rs("AddTime") = Now()
	rs.update
	rs.close : set rs = Nothing
	Call SweetAlert("success", "操作成功", "设置成功！", "false", "location.replace('Ky_Template.Asp')")
End Function
%>