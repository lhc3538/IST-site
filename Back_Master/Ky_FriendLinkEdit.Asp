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
'神笔(26)天成网络享有软件著作权，如果您擅自删除这里的版权信息,我们将追究法律责任.
'未经公司同意禁止更改以下代码.
'*******************************************************
Dim Result
Result = request.QueryString("Result")
Dim ID, LinkName, ViewFlag, LinkType, LinkFace, LinkUrl, Remark
ID = request.QueryString("ID")
Call FriendLinkEdit()
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
<link rel="stylesheet" href="KEditor_keyicms/themes/default/default.css" />
<script charset="utf-8" src="KEditor_keyicms/kindeditor.js"></script>
<script charset="utf-8" src="KEditor_keyicms/lang/zh_CN.js"></script>
<script>
KindEditor.ready(function(K) {
	var editor = K.editor({
		uploadJson : 'KEditor_keyicms/asp/upload_json.asp',
		fileManagerJson : 'KEditor_keyicms/asp/file_manager_json.asp',
		allowFileManager : true
	});
	K('#ImageLink').click(function() {
		editor.loadPlugin('image', function() {
			editor.plugin.imageDialog({
				showRemote : true,
				imageUrl : K('#LinkFace').val(),
				clickFn : function(url, title, width, height, border, align) {
					K('#LinkFace').val(url);
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
	If InStr(CompanyAdminPurview, "|10,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style=""color:red; font-size:14px;"")>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If
	%>
    <div class="main">
        <div class="row ">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i><%=Keyicms_Lang_Copyright(14)%></li>
             <li class="active"><a href="">企业应用</a></li>
                <li class="active">友情链接管理</li>
            </ol>	
        </div>	
        <div class="row">
            <div class="col-lg-12 hideInIE8">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-edit"></i><span class="break"></span><strong>友情链接管理</strong></h2> 
                    </div>
                    <div class="panel-body">
                        <a class="ShortCut1" href="Ky_FriendLinkEdit.Asp?Result=Add">  
                            <i class="fa fa-sign-in"></i>
                            <div class="title">友情链接添加</div>
                        </a>					
                        <a class="ShortCut2" href="Ky_FriendLink.Asp">  
                            <i class="fa fa-link"></i>
                            <div class="title">友情链接列表</div>
                        </a>										
                    </div>
                </div>
            </div>
        </div>
        <div class="row">
            <div class="col-md-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-gears"></i><span class="break"></span><strong><%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改<%End If%>友情链接</strong></h2>
                    </div>
                    <form name="editForm" method="post" action="?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>" class="form-horizontal">
                    <div class="panel-body">
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">链接名称：</label>
                            <div class="col-md-4">
                                <input name="LinkName" type="text" id="LinkName" class="form-control" value="<%=LinkName%>">
                                <label class="checkbox-inline" for="ViewFlag">
                                    <input type="checkbox" id="ViewFlag" name="ViewFlag" value="1" <%if ViewFlag then response.write ("checked")%>> 是否生效
                                </label>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">显示类型：</label>
                            <div class="col-md-4">
                            	<label class="radio-inline" for="LinkType1">
                                    <input name="LinkType" id="LinkType1" type="radio" value="1" <%if LinkType then response.write ("checked=checked")%>> 图片
                                </label>
                            	<label class="radio-inline" for="LinkType2">
                                    <input name="LinkType" id="LinkType2" type="radio" value="0" <%if not LinkType then response.write ("checked=checked")%>> 文字
                                </label>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">前台显示文字、图片：</label>
                            <div class="col-md-4">
                                <div class="input-group">
                                    <input type="text" id="LinkFace" name="LinkFace" class="form-control" value="<%=LinkFace%>">
                                    <span class="input-group-btn">
                                    <button type="button" id="ImageLink" class="btn btn-default">上传图片</button>
                                    </span>
                                </div>
                                <span class="help-block"></span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">链接网址：</label>
                            <div class="col-md-4">
                                <input name="LinkUrl" type="text" id="LinkUrl" class="form-control" value="<%=LinkUrl%>">
                                 <span class="help-block">请以http://开头</span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">备注：</label>
                            <div class="col-md-4">
                                <textarea name="Remark" id="Remark" class="form-control" rows="8" cols="60"><%=Remark%></textarea>
                                <span class="help-block"></span>
                            </div>
                        </div>
                    </div>
                        <div class="form-group">
                            <div class="col-md-2 control-label"></div>
                            <div class="col-md-10">
                                <input name="submitSaveEdit" type="submit" id="submitSaveEdit" class="btn btn-danger btn-lg" value="保存友情链接">
                                <input type="button" class="btn btn-danger btn-lg" value="<%=Keyicms_Lang_System(32)%>" onClick="history.back(-1)">
                            </div>	
                        </div>
                    </form>
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
Sub FriendLinkEdit()
    Dim Action, rsCheckAdd, rs, sql
    Action = request.QueryString("Action")
    If Action = "SaveEdit" Then
        Set rs = server.CreateObject("adodb.recordset")
        If Len(Trim(request.Form("LinkName")))<2 Then
			Call SweetAlert("warning", "友情提示", "请填写网站名称并保持至少在两个汉字以上！", "false", "history.back(-1)")
            response.End
        End If
        If Trim(request.Form("LinkFace")) = "" Then
			Call SweetAlert("warning", "友情提示", "请填写前台显示文字或上传友情链接LOGO图片！", "false", "history.back(-1)")
            response.End
        End If
        If request.Form("LinkType") = 0 Then
            If Trim(request.Form("LinkFace")) = "" Then
				Call SweetAlert("warning", "友情提示", "请填写前台显示文字或图片地址！", "false", "history.back(-1)")
                response.End
            End If
        End If
        If Trim(request.Form("LinkUrl")) = "" Then
			Call SweetAlert("warning", "友情提示", "请填写友情链接网址！", "false", "history.back(-1)")
            response.End
        End If
        If Result = "Add" Then
            sql = "select * from keyicms_FriendLink"
            rs.Open sql, conn, 1, 3
            rs.addnew
            rs("LinkName") = Trim(Request.Form("LinkName"))
            If Request.Form("ViewFlag") = 1 Then
                rs("ViewFlag") = Request.Form("ViewFlag")
            Else
                rs("ViewFlag") = 0
            End If
            rs("LinkFace") = Trim(Request.Form("LinkFace"))
            rs("LinkUrl") = Trim(Request.Form("LinkUrl"))
            If Request.Form("LinkType") = 1 Then
                rs("LinkType") = Request.Form("LinkType")
            Else
                rs("LinkType") = 0
            End If
            rs("Remark") = Trim(Request.Form("Remark"))
            rs("AddTime") = Now()
        End If
        If Result = "Modify" Then
            sql = "select * from keyicms_FriendLink where ID="&ID
            rs.Open sql, conn, 1, 3
            rs("LinkName") = Trim(Request.Form("LinkName"))
            If Request.Form("ViewFlag") = 1 Then
                rs("ViewFlag") = Request.Form("ViewFlag")
            Else
                rs("ViewFlag") = 0
            End If
            rs("LinkFace") = Trim(Request.Form("LinkFace"))
            rs("LinkUrl") = Trim(Request.Form("LinkUrl"))
            If Request.Form("LinkType") = 1 Then
                rs("LinkType") = Request.Form("LinkType")
            Else
                rs("LinkType") = 0
            End If
            rs("Remark") = Trim(Request.Form("Remark"))
        End If
        rs.update
        rs.close : set rs = nothing
		Call SweetAlert("success", "操作成功", "设置成功！", "false", "location.replace('Ky_FriendLink.Asp')")
    Else
        If Result = "Modify" Then
            Set rs = server.CreateObject("adodb.recordset")
            sql = "select * from keyicms_FriendLink where ID="& ID
            rs.Open sql, conn, 1, 1
            LinkName = rs("LinkName")
            ViewFlag = rs("ViewFlag")
            LinkType = rs("LinkType")
            LinkFace = rs("LinkFace")
            LinkUrl = rs("LinkUrl")
            Remark = rs("Remark")
            rs.close : set rs = nothing
        End If
    End If
End Sub
%>