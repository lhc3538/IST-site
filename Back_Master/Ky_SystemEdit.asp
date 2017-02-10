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
'神笔(20)天成网络享有软件著作权，如果您擅自删除这里的版权信息,我们将追究法律责任.
'未经公司同意禁止更改以下代码.
'*******************************************************
dim Result
Result = request.QueryString("Result")
dim ID, SystemName, ViewFlag, SystemPath, SystemDir
dim SeoTitle, SeoKeywords, SeoDescription

ID = request.QueryString("ID")
Call MainEdit()
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
				showRemote : true,
				imageUrl : K('#Pic').val(),
				clickFn : function(url, title, width, height, border, align) {
					K('#Pic').val(url);
					editor.hideDialog();
				}
			});
		});
	});
});
</script>
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
            <div class="col-md-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-gears"></i><span class="break"></span><strong>修改功能模块</strong></h2>
                    </div>
                    <div class="panel-body">
                    <form name="editForm" method="post" action="?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>" class="form-horizontal"role="form">
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">功能名称：</label>
                            <div class="col-md-4">
                                <input name="SystemName" type="text" id="SystemName" class="form-control" value="<%=SystemName%>" maxlength="100">
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">是否生效：</label>
                            <div class="col-md-5">
                                <label class="switch switchtop switch-danger">
                                    <input type="checkbox" class="switch-input" name="ViewFlag" value="1" id="ViewFlag" <%if ViewFlag then response.write ("checked")%> />
                                    <span class="switch-label" data-on="开启" data-off="关闭"></span>
                                    <span class="switch-handle"></span>
                                </label>
                            </div>
                        </div> 
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">功能路径：</label>
                            <div class="col-md-5">
                                <p class="form-control-static"><%=SystemPath%></p>
                            </div>
                        </div> 
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">功能文件夹：</label>
                            <div class="col-md-3">
                                <input name="SystemDir" type="text" id="SystemDir" class="form-control" value="<%=SystemDir%>">
                                <span class="help-block">若修改文件夹，请到导航选择重新选择，否则导航链接可能会出错。</span>
                            </div>
                        </div>
                        <!--<div class="form-group">
                            <label class="col-md-2 control-label">是否启用PC栏目页：</label>
                            <div class="col-md-10">
                                <label class="radio-inline" for="PcFlag1">
                                    <input name="PcFlag" id="PcFlag1" type="radio" value="1" <%if PcFlag then response.Write "checked"%>> <%=Keyicms_Lang_System(30)%>
                                </label>
                                <label class="radio-inline" for="PcFlag2">
                                    <input name="PcFlag" id="PcFlag2" type="radio" value="0" <%if not PcFlag then response.Write "checked"%>> <%=Keyicms_Lang_System(31)%> 
                                </label>
                                <span class="help-block">若未启用栏目页，将自动使用列表页生成栏目页（企业信息使用信息排序第一的作为栏目页）</span>
                            </div>
                        </div>                        
                        <div class="form-group">
                            <label class="col-md-2 control-label">是否启用手机栏目页：</label>
                            <div class="col-md-10">
                                <label class="radio-inline" for="McFlag1">
                                    <input name="McFlag" id="McFlag1" type="radio" value="1" <%if McFlag then response.Write "checked"%>> <%=Keyicms_Lang_System(30)%>
                                </label>
                                <label class="radio-inline" for="McFlag2">
                                    <input name="McFlag" id="McFlag2" type="radio" value="0" <%if not McFlag then response.Write "checked"%>> <%=Keyicms_Lang_System(31)%> 
                                </label>
                                <span class="help-block">若未启用栏目页，将自动使用列表页生成栏目页（企业信息使用信息排序第一的作为栏目页）</span>
                            </div>
                        </div>-->                        

                        <div class="form-group">
                            <label class="col-md-2 control-label">SEO标题：</label>
                            <div class="col-md-6">
                                <input name="SeoTitle" type="text" id="SeoTitle" class="form-control" value="<%=SeoTitle%>" maxlength="250">
                                <span class="help-block">SEO标题：必填</span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label">SEO关键词：</label>
                            <div class="col-md-6">
                                <input name="SeoKeywords" type="text" id="SeoKeywords" class="form-control" value="<%=SeoKeywords%>" maxlength="250">
                                <span class="help-block">SEO：多个关键词请用“,”半角逗号分开</span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label">SEO描述：</label>
                            <div class="col-md-6">
                                <input name="SeoDescription" type="text" id="SeoDescription" class="form-control" value="<%=SeoDescription%>" maxlength="250">
                                <span class="help-block">SEO：请填写描述</span>
                            </div>
                        </div>

                        <div class="form-group">
                            <div class="col-md-2 control-label"></div>
                            <div class="col-md-10">
                                <input name="submitSaveEdit" type="submit" id="submitSaveEdit" class="btn btn-danger btn-lg" value="保存功能模块">
                                <input type="button" class="btn btn-danger btn-lg" value="<%=Keyicms_Lang_System(32)%>" onClick="history.back(-1)">
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
Sub MainEdit()
Dim Action, rsCheckAdd, rs, sql
Action = request.QueryString("Action")
If Action = "SaveEdit" Then
	Set rs = server.CreateObject("adodb.recordset")
	If Len(Trim(request.Form("SystemDir")))<1 Then
		Call SweetAlert("warning", "友情提示", "请填写功能文件夹名称！", "false", "history.back(-1)")
		response.End
	End If
	If Len(Trim(request.Form("SeoTitle")))<1 Then
		Call SweetAlert("warning", "友情提示", "请填写SEO标题！", "false", "history.back(-1)")
		response.End
	End If

	If Result = "Modify" Then
		sql = "select * from keyicms_System where ID="&ID
		rs.Open sql, conn, 1, 3
		rs("SystemName") = Trim(Request.Form("SystemName"))
		If Request.Form("ViewFlag") = 1 Then
			rs("ViewFlag") = Request.Form("ViewFlag")
		Else
			rs("ViewFlag") = 0
		End If
		rs("SeoTitle") = Trim(Request.Form("SeoTitle"))
		rs("SeoKeywords") = Trim(Request.Form("SeoKeywords"))
		rs("SeoDescription") = Request.Form("SeoDescription")

		If LCase(trim(rs("SystemDir"))) <> LCase(trim(request("SystemDir"))) Then
			Call fldrename("../"&rs("SystemDir"), trim("../"&request("SystemDir")))
			If mStatus Then Call fldrename("../m/"&rs("SystemDir"), trim("../m/"&request("SystemDir")))
			rs("SystemDir") = Trim(Request.Form("SystemDir"))
		End If
		rs.update
		rs.close : set rs = nothing
	End If
	Call SweetAlert("success", "提示", "设置成功！", "false", "location.replace('Ky_System.Asp')")
Else
	If Result = "Modify" Then
		Set rs = server.CreateObject("adodb.recordset")
		sql = "select * from keyicms_System where ID="& ID
		rs.Open sql, conn, 1, 1
		SystemName = rs("SystemName")
		ViewFlag = rs("ViewFlag")
		SystemPath = rs("SystemPath")
		SystemDir = rs("SystemDir")
		SeoTitle = rs("SeoTitle")
		SeoKeywords = rs("SeoKeywords")
		SeoDescription = rs("SeoDescription")
		rs.close : set rs = nothing
	End If
End If
End Sub
%>