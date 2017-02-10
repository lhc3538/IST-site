<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="../Include/system_language.asp" -->
<!--#include file="../Include/Version.Asp" -->
<!--#include file="CheckAdmin.Asp"-->
<!--#include file="../Html/html_Config.Asp"-->
<!--#include file="../Html/html_Function.Asp"-->
<!--#include file="../Html/html_To_Case.Asp"-->
<%

dim Result
Result = request.QueryString("Result")
dim ID, CaseName, ViewFlag, CommendFlag, MobileFlag, ClassSeo, SortName, SortID, SortPath
dim SeoKeywords, SeoDescription, SeoTitle
dim SmallPic, BigPic, Sequence, Content, OtherPic, ClickNumber
ID = request.QueryString("ID")
call CaseEdit()
call SiteInfo
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
<script charset="utf-8" src="KEditor_keyicms/kindeditor.js"></script>
<script charset="utf-8" src="KEditor_keyicms/lang/zh_CN.js"></script>
<script>
KindEditor.ready(function(K) {
	var editor = K.create('textarea[name="Content"]', {
		cssPath : 'KEditor_keyicms/plugins/code/prettify.css',
		uploadJson : 'KEditor_keyicms/asp/upload_json.asp',
		fileManagerJson : 'KEditor_keyicms/asp/file_manager_json.asp',
		allowFileManager : true,
		afterBlur : function() {
		this.sync();
		K.ctrl(document, 13, function() {
		K('form[name=editForm]')[0].submit();
		});
		K.ctrl(this.edit.doc, 13, function() {
		K('form[name=editForm]')[0].submit();
		});
		}
	});
	K('#SmallImage').click(function() {
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
	K('#BigImage').click(function() {
		editor.loadPlugin('image', function() {
			editor.plugin.imageDialog({
				imageUrl : K('#BigPic').val(),
				clickFn : function(url, title, width, height, border, align) {
					K('#BigPic').val(url);
					editor.hideDialog();
				}
			});
		});
	});
	K('#J_selectImage').click(function() {
		editor.loadPlugin('multiimage', function() {
			var aVal=""
			editor.plugin.multiImageDialog({
				clickFn : function(urlList) {
					var div = K('#OtherPic');
					aVal=div.val()
					K.each(urlList, function(i, data) {
						if (div.val() == "")
							aVal =aVal + data.url
						else
							aVal =aVal + "|" +  data.url
					
						div.val(aVal);
					});
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
	If InStr(CompanyAdminPurview, "|4,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style=""color:red; font-size:14px;"")>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If
	%>
    <div class="main">
        <div class="row ">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i> <%=Keyicms_Lang_Copyright(14)%></li>
                <li class="active"><a href="">功能模块</a></li>
                <li class="active">资源管理</li>
            </ol>	
        </div>	
     	<div class="row">
            <div class="col-lg-12 hideInIE8">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-edit"></i><span class="break"></span><strong>资源管理</strong></h2> 
                    </div>
                    <div class="panel-body">
                        <a class="ShortCut1" href="Ky_CaseEdit.Asp?Result=Add">  
                            <i class="fa fa-picture-o"></i>
                            <div class="title">添加资源</div>
                        </a>					
                        <a class="ShortCut2" href="Ky_Case.Asp">  
                            <i class="fa fa-th-large"></i>
                            <div class="title">资源列表</div>
                        </a>	
                        <a class="ShortCut3" href="Ky_CaseSort.Asp">  
                            <i class="fa fa-sitemap"></i>
                            <div class="title">资源分类</div>
                        </a>													
                    </div>
                </div>
            </div>
        </div>
		<div class="row">
			<div class="col-md-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-gears"></i><span class="break"></span><strong><%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改<%End If%>资源</strong></h2>
                    </div>
                    <div class="panel-body">
                    <form name="editForm" method="post" action="?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>" class="form-horizontal">
                        <div class="form-group">
                            <label class="col-md-2 control-label">资源名称：</label>
                            <div class="col-md-4">
                                <input name="CaseName" type="text" id="CaseName"  class="form-control"  value="<%=CaseName%>" >
                            </div>
                        </div> 
                        <%If EnableSeo = True Then%>
                        <div class="form-group">
                            <label class="col-md-2 control-label">资源关键词：</label>
                            <div class="col-md-6">
                            <input name="SeoKeywords" type="text" id="SeoKeywords" class="form-control" value="<%=SeoKeywords%>" maxlength="250">
                                 <span class="help-block">SEO：多个关键词请用“,”半角逗号分开</span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label">资源描述：</label>
                            <div class="col-md-6">
                            <input name="SeoDescription" type="text" id="SeoDescription" class="form-control" value="<%=SeoDescription%>" maxlength="250">
                                 <span class="help-block">SEO：请填写资源描述</span>
                            </div>
                        </div>
                        <%End If%>
                        <div class="form-group">
                            <label class="col-md-2 control-label">资源类别：</label>
                            <div class="col-md-5">
                                <select size="1" class="form-control" name="SortSelect" id="SortSelect" onChange="AddSort()">
                                <option value="" valuepath="">请选择...</option>
                                <%Call CaseSort(0, 0)%>
                                </select>
                                <span class="help-block">
                                    <input name="SortID" type="hidden" id="SortID" value="<%=SortID%>" readonly class="form-control">
                                    <input name="SortPath" type="hidden" id="SortPath" value="<%=SortPath%>" readonly class="form-control">
                                </span>
                            </div>
                        </div> 

                        <div class="form-group">
                            <label class="col-md-2 control-label">资源链接：</label>
                            <div class="col-md-10">
                                <textarea name="Content" style="width:95%;height:400px;visibility:hidden;" class="form-control"><%=Server.HTMLEncode(Content)%></textarea>
                                <span class="help-block"></span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label">排序：</label>
                            <div class="col-md-3">
                                <input name="Sequence" type="text" id="Sequence" class="form-control" value="<%=Sequence%>" onKeyDown="OnKeyDown()" onchange="OnChange(this, <%=Sequence%>)">
                                <span class="help-block">数字越小排名越靠前。</span>
                            </div>
                        </div>
                        <div class="form-group">
                            <div class="col-md-2 control-label"></div>
                            <div class="col-md-10">
                                <input name="submitSaveEdit" type="submit" id="submitSaveEdit" class="btn btn-danger btn-lg" value="保存资源信息">
                                <input type="button" class="btn btn-danger btn-lg" value="<%=Keyicms_Lang_System(32)%>" onClick="history.back(-1)">
                            </div>	
                        </div>
                    </form>
                    </div>
				</div>		
			</div>
        </div>
    </div>
    <!-- end: Content -->

	<!--#Include file="Keyicms_End.Asp"-->				
</div>
<!--#Include file="Keyicms_js.Asp"-->
</body>
</html>

<%
Sub CaseEdit()
dim Action, rs, sql
Action = request.QueryString("Action")
if Action = "SaveEdit" then
	set rs = server.createobject("adodb.recordset")
	if len(trim(request.Form("CaseName")))<1 then
		Call SweetAlert("warning", "友情提示", "请填写资源名称！", "false", "history.back(-1)")
		response.end
	end If
	if Request.Form("SortID")="" and Request.Form("SortPath")="" then
		Call SweetAlert("warning", "友情提示", "请选择所属分类！", "false", "history.back(-1)")
		response.End
	end If
	if ltrim(request.Form("Content")) = "" then
		Call SweetAlert("warning", "友情提示", "请填写资源链接！", "false", "history.back(-1)")
		response.end
	end If

    If Result="Add" Then
		sql="select * from keyicms_Case"
		rs.open sql,conn,1,3
		rs.addnew
		rs("CaseName") = trim(Request.Form("CaseName"))

		rs("ViewFlag") = 1

		if Request.Form("CommendFlag") = 1 then
			rs("CommendFlag") = Request.Form("CommendFlag")
		else
			rs("CommendFlag") = 0
		end if
		if Request.Form("MobileFlag") = 1 then
			rs("MobileFlag") = Request.Form("MobileFlag")
		else
			rs("MobileFlag") = 0
		end if
		If Request.Form("ClassSeo") = "" then
			If Request.Form("oAutopinyin") = "1" then rs("ClassSeo") = TimeToName()
			If Request.Form("oAutopinyin") = "2" then rs("ClassSeo") = Left(Pinyin(trim(request.form("CaseName"))),200)
		Else
			rs("ClassSeo") = trim(Request.form("ClassSeo"))
		End If
		rs("SortID")=Request.Form("SortID")
		rs("SortPath")=Request.Form("SortPath")
		rs("SmallPic")=trim(Request.Form("SmallPic"))
		rs("BigPic")=trim(Request.Form("BigPic"))
		rs("OtherPic")=trim(Request.Form("OtherPic"))
		rs("Content") = rtrim(Request.Form("Content"))
		rs("AddTime") = now()
		rs("UpdateTime") = now()
		rs("SeoTitle") = trim(Request.Form("SeoTitle"))
		rs("SeoKeywords") = trim(Request.Form("SeoKeywords"))
		rs("SeoDescription") = trim(Request.Form("SeoDescription"))

		rs("Sequence") = trim(Request.Form("Sequence"))
		rs.update
		rs.close : set rs = nothing
		
		set rs = server.createobject("adodb.recordset")
		sql="select top 1 ID,ClassSeo from keyicms_Case order by ID desc"
		rs.open sql,conn,1,1
		ID = rs("ID")
		rs.close : set rs = nothing
	  
		'Call Html_To_Case(ID, "PC")
		If mStatus Then Call Html_To_Case(ID, "Mobile")
	End If
	If Result="Modify" Then
		sql="select * from keyicms_Case where ID="&ID
		rs.open sql,conn,1,3
		rs("CaseName") = trim(Request.Form("CaseName"))

		rs("ViewFlag") = 1

		if Request.Form("CommendFlag") = 1 then
			rs("CommendFlag") = Request.Form("CommendFlag")
		else
			rs("CommendFlag") = 0
		end if
		if Request.Form("MobileFlag") = 1 then
			rs("MobileFlag") = Request.Form("MobileFlag")
		else
			rs("MobileFlag") = 0
		end if
		rs("SortID") = Request.Form("SortID")
		rs("SortPath") = Request.Form("SortPath")
		rs("SmallPic") = trim(Request.Form("SmallPic"))
		rs("BigPic") = trim(Request.Form("BigPic"))
		rs("OtherPic") = trim(Request.Form("OtherPic"))
		rs("Content") = rtrim(Request.Form("Content"))
		rs("UpdateTime") = now()
		rs("SeoTitle") = trim(Request.Form("SeoTitle"))
		rs("SeoKeywords") = trim(Request.Form("SeoKeywords"))
		rs("SeoDescription") = trim(Request.Form("SeoDescription"))

		rs("Sequence") = trim(Request.Form("Sequence"))

		ID = rs("ID")
		rs.update
		rs.close : set rs = nothing
	  
		'Call Html_To_Case(ID, "PC")
		If mStatus Then Call Html_To_Case(ID, "Mobile")
	End If
	Call SweetAlert("success", "操作成功", "设置成功，相关页面已更新！", "false", "location.replace('Ky_Case.Asp')")
else
  	if Result = "Modify" then
		set rs = server.createobject("adodb.recordset")
		sql = "select * from keyicms_Case where ID="& ID
		rs.open sql, conn, 1, 1
		if rs.bof and rs.eof then
			response.write ("<center>数据库记录读取错误！</center>")
			response.end
		end if
		CaseName = rs("CaseName")
		ViewFlag = rs("ViewFlag")
		MobileFlag = rs("MobileFlag")
		CommendFlag = rs("CommendFlag")
		Sequence = rs("Sequence")
		ClassSeo = rs("ClassSeo")&Separated&rs("ID")
		SortName=SortText(rs("SortID"))
		SortID = rs("SortID")
		SortPath = rs("SortPath")
		SmallPic = rs("SmallPic")
		BigPic = rs("BigPic")
		OtherPic = rs("OtherPic")
		Content = rs("Content")
		ClickNumber = rs("ClickNumber")
		Sequence = rs("Sequence")
		SeoTitle = rs("SeoTitle")
		SeoKeywords = rs("SeoKeywords")
		SeoDescription = rs("SeoDescription")
		rs.close : set rs = nothing
	else
		ClickNumber = 10
		Sequence = 10
    end if
end if
End Sub

Function SortText(ID)
	Dim rs,sql
	set rs = Server.CreateObject("adodb.recordset")
	sql = "Select * From keyicms_CaseSort where ID="&ID
	rs.open sql, conn, 1, 1
	SortText = rs("SortName")
	rs.close : set rs = nothing
End Function

Function CaseSort(ID, ClassNum)
	Dim rs
	set rs = Server.CreateObject("adodb.recordset")
	rs.open "select * from Keyicms_CaseSort where ParentID="&ID&" Order by ID asc", conn, 1, 2
	Do Until rs.eof
		If SortID = rs("ID") Then
			response.Write "<option value='"&rs("ID")&"' valuepath='"&rs("SortPath")&"' selected>"
		Else
			response.Write "<option value='"&rs("ID")&"' valuepath='"&rs("SortPath")&"'>"
		End If
		For ClassN = 1 to ClassNum
			response.Write "│"
		Next
		response.Write "│├ "&rs("SortName")
		response.Write "</option>"
		Call CaseSort(rs("ID"), ClassNum + 1)
		rs.movenext
	Loop
	rs.close : set rs = nothing
End Function
%>