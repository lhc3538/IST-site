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
'神笔(2)天成网络享有软件著作权，如果您擅自删除这里的版权信息,我们将追究法律责任.
'未经公司同意禁止更改以下代码.
'*******************************************************
Select Case request.QueryString("Action")
	Case "Save"
		SaveSiteInfo
	Case Else
		ViewSiteInfo
End Select
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
				imageUrl : K('#MobileLogo').val(),
				clickFn : function(url, title, width, height, border, align) {
					K('#MobileLogo').val(url);
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
	If InStr(CompanyAdminPurview, "|17,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style=""color:red; font-size:9pt;"")>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If
	%>
    <!-- start: Content -->
    <div class="main">
        <div class="row ">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i> <%=Keyicms_Lang_Copyright(14)%></li>
                <li><%=Keyicms_Lang_Copyright(2)%></li>
                <li class="active"><%=Keyicms_Lang_sys_Mobile(0)%></li>
            </ol>	
        </div>	
        <div class="row">
            <div class="col-md-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-tablet"></i><span class="break"></span><strong><%=Keyicms_Lang_sys_Mobile(0)%></strong></h2>
                    </div>
                    <div class="panel-body">
                    <form action="?Action=Save" method="post" class="form-horizontal">
                        <div class="form-group">
                            <label class="col-md-2 control-label"><%=Keyicms_Lang_sys_Mobile(1)%>：</label>
                            <div class="col-md-5">
                                <label class="switch switchtop switch-danger">
                                  <input type="checkbox" class="switch-input" name="MobileJump" value="-1" id="MobileJump" <%if MobileJump then response.write ("checked")%> />
                                  <span class="switch-label" data-on="开启" data-off="关闭"></span>
                                  <span class="switch-handle"></span>
                                </label>
                                 <span class="help-block"></span>
                            </div>
                        </div> 
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input"><%=Keyicms_Lang_sys_Mobile(2)%>：</label>
                            <div class="col-md-5">
                                <input name="MobileTitle" type="text" id="MobileTitle"  class="form-control" value="<%=MobileTitle%>">
                                <span class="help-block">请填写网站标题，这里的信息显示在浏览器顶部以及搜索引擎搜索出的网站主页标题</span>
                            </div>
                        </div>

                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">手机访问路径：</label>
                            <div class="col-md-4">
                                <p class="form-control-static">/m</p>
                                <span class="help-block">手机网站路径为m，请勿修改文件夹</span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input"><%=Keyicms_Lang_sys_Mobile(4)%>：</label>
                            <div class="col-md-6"> 
                                <div class="input-group">
                                    <input name="MobileLogo" type="text" id="MobileLogo" value="<%=MobileLogo%>"  class="form-control" />
                                    <span class="input-group-btn">
                                    <button type="button" id="image" class="btn btn-default"><%=Keyicms_Lang_System(27)%></button>
                                    </span>
                                </div>
                                <span class="help-block"><font color="red">*//注： PNG格式</font> 请上传手机网站LOGO图标，建议PNG格式，大小根据实际模板样式确定</span>
                            </div>
                        </div> 

                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input"><%=Keyicms_Lang_sys_Mobile(5)%>：</label>
                            <div class="col-md-6">
                                <textarea name="MobileKeywords"  id="MobileKeywords" class="form-control" rows="3" cols="80"><%=MobileKeywords%></textarea>
                                <span class="help-block">填写keywords（网站关键词），多个关键词请用“,”分割。</span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input"><%=Keyicms_Lang_sys_Mobile(6)%>：</label>
                            <div class="col-md-6">
                                <textarea name="MobileDescriptions" id="MobileDescriptions" class="form-control" style="padding:5px;" rows="3" cols="80"><%=MobileDescriptions%></textarea>
                                <span class="help-block">填写Descriptions（网站描述）。</span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input"><%=Keyicms_Lang_sys_Mobile(7)%>：</label>
                            <div class="col-md-4">
                                <input name="MobileTelephone" type="text" id="MobileTelephone"  class="form-control" value="<%=MobileTelephone%>">
                                <span class="help-block"><font color="red">*</font>请填写手机快速拨号电话号码</span>
                            </div>
                        </div>
                        <div class="form-group">
                            <div class="col-md-2 control-label"></div>
                            <div class="col-md-10">
                                <input name="submitSaveEdit" type="submit" id="submitSaveEdit" class="btn btn-danger btn-lg" value="<%=Keyicms_Lang_sys_Mobile(10)%>">
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
Function SaveSiteInfo()
    If Len(Trim(request.Form("MobileTitle")))<4 Then
		Call SweetAlert("warning", "友情提示", "请详细填写您的网站标题并保持至少在两个汉字以上！", "false", "history.back(-1)")
        response.End
    End If
    If Len(Trim(request.Form("MobileKeywords")))>250 Then
		Call SweetAlert("warning", "友情提示", "请详细填写网站关键词并保持在250个字符以内！", "false", "history.back(-1)")
        response.End
    End If
    If Len(Trim(request.Form("MobileDescriptions")))>250 Then
		Call SweetAlert("warning", "友情提示", "请详细填写网站描述并保持在250个字符以内！", "false", "history.back(-1)")
        response.End
    End If

    Call SiteInfo()
    Dim rs, sql
    Set rs = server.CreateObject("adodb.recordset")
    sql = "select top 1 * from keyicms_MobileSite"
    rs.Open sql, conn, 1, 3
    rs("MobileTitle") = Trim(Request.Form("MobileTitle"))
	rs("MobileTelephone") = Trim(Request.Form("MobileTelephone"))
    rs("MobileKeywords") = Trim(Request.Form("MobileKeywords"))
    rs("MobileDescriptions") = Trim(Request.Form("MobileDescriptions"))
    rs("MobileLogo") = Trim(Request.Form("MobileLogo"))
	If Trim(Request.Form("MobileJump")) = "-1" then
		rs("MobileJump") = Trim(Request.Form("MobileJump"))
    Else
		rs("MobileJump") = 0
	End IF
	rs.update
    rs.Close
    Set rs = Nothing
	Call SweetAlert("success", "操作成功", "手机版参数设置成功！", "false", "location.replace('Keyicms_MobileSite.Asp')")
End Function

Function ViewSiteInfo()
	Dim rs, sql
	Set rs = server.CreateObject("adodb.recordset")
	sql = "select top 1 * from keyicms_MobileSite"
	rs.Open sql, conn, 1, 1
	If rs.bof And rs.EOF Then
		response.Write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>"&Keyicms_Lang_System(28)&"</font></div>")
		response.End
	Else
		MobileTitle = rs("MobileTitle")
		MobileJump = rs("MobileJump")
		MobileKeywords = rs("MobileKeywords")
		MobileDescriptions = rs("MobileDescriptions")
		MobileTelephone = rs("MobileTelephone")
		MobileLogo = rs("MobileLogo")
		rs.Close
		Set rs = Nothing
	End If
End Function
%>