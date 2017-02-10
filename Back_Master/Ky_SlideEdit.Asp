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
'神笔(28)天成网络享有软件著作权，如果您擅自删除这里的版权信息,我们将追究法律责任.
'未经公司同意禁止更改以下代码.
'*******************************************************
dim Result
Result = request.QueryString("Result")
dim ID, Url, Pic, State, SlideName
ID = request.QueryString("ID")
Call SlideEdit()
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
    var editor = K.editor({
        uploadJson : 'KEditor_keyicms/asp/upload_json.asp',
        fileManagerJson : 'KEditor_keyicms/asp/file_manager_json.asp',
        allowFileManager : true
    });
    K('#imagePic').click(function() {
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
	If InStr(CompanyAdminPurview, "|12,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style=""color:red; font-size:14px;"")>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If
	%>
    <div class="main">
        <div class="row ">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i> <%=Keyicms_Lang_Copyright(14)%></li>
                <li><%=Keyicms_Lang_Copyright(2)%></li>
                <li class="active">幻灯片管理</li>
            </ol>	
        </div>	
         <div class="row ">
            <div class="col-lg-12 hideInIE8">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-edit"></i><span class="break"></span><strong>幻灯片管理</strong></h2> 
                    </div>
                    <div class="panel-body">
                        <a class="ShortCut1" href="Ky_SlideEdit.Asp?Result=Add">  
                            <i class="fa fa-edit"></i>
                            <div class="title">添加幻灯片</div>
                        </a>					
                        <a class="ShortCut2" href="Ky_Slide.Asp">  
                            <i class="fa fa-film"></i>
                            <div class="title">幻灯片管理</div>
                        </a>														
                    </div>
                </div>
            </div>
        </div>
        <div class="row">
            <div class="col-md-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-gears"></i><span class="break"></span><strong><%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改<%End If%>幻灯片</strong></h2>
                    </div>
                    <div class="panel-body">
                    <form name="editForm" method="post" action="Ky_SlideEdit.Asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>" class="form-horizontal">
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">幻灯片名称：</label>
                            <div class="col-md-4">
                                <input name="SlideName" type="text" id="SlideName" class="form-control" value="<%=SlideName%>">
                                 <span class="help-block"></span>
                            </div>
                        </div> 
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">幻灯片路径：</label>
                            <div class="col-md-4">
                            	<div class="input-group">
                                    <input type="text" id="Pic" name="Pic" class="form-control" value="<%=Pic%>" readonly>
                                    <span class="input-group-btn">
                                    <button type="button" id="imagePic" class="btn btn-default">上传图片</button>
                                    </span>
                                </div>
                                <span class="help-block"></span>
                            </div>
                        </div> 
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">幻灯片链接：</label>
                            <div class="col-md-4">
                            	<input name="Url" type="text" class="form-control" id="Url" value="<%=Url%>">
                                <span class="help-block"></span>
                            </div>
                        </div> 
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">幻灯片排序：</label>
                            <div class="col-md-2">
                                <input name="State" type="text" id="State" class="form-control" value="<%=State%>" onKeyDown="OnKeyDown()" onchange="OnChange(this, <%=State%>)" />
                                 <span class="help-block"></span>
                            </div>
                        </div>
                        <div class="form-group">
                            <div class="col-md-2 control-label"></div>
                            <div class="col-md-10">
                                <input name="submitSaveEdit" type="submit" id="submitSaveEdit" class="btn btn-danger btn-lg" value="保存幻灯片">
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
sub SlideEdit()
dim Action, rs, sql
Action = request.QueryString("Action")
if Action = "SaveEdit" then
    set rs = server.createobject("adodb.recordset")
    if request.Form("Pic") = "" then
		Call SweetAlert("warning", "友情提示", "请填写前台显示的图片地址！", "false", "history.back(-1)")
		response.end
    end if
    if Result = "Add" then
		sql = "select * from keyicms_Slide"
		rs.open sql,conn,1,3
		rs.addnew
		rs("SlideName") = trim(Request.Form("SlideName"))
		rs("Pic") = trim(Request.Form("Pic"))
		if trim(request.Form("Url"))="" then
			rs("Url") = "#"
		else
			rs("Url") = trim(Request.Form("Url"))	
		end if
		rs("State") = trim(Request.Form("State"))
	end if
	if Result = "Modify" then
		sql = "select * from keyicms_Slide where ID="&ID
		rs.open sql,conn,1,3
		rs("SlideName") = trim(Request.Form("SlideName"))
		rs("Pic") = trim(Request.Form("Pic"))
		if trim(request.Form("Url"))="" then
			rs("Url") = "#"
		else
			rs("Url") = trim(Request.Form("Url"))	
		end if
		rs("State") = trim(Request.Form("State"))
	 end if
	rs.update
	rs.close : set rs = nothing
	Call SweetAlert("success", "操作成功", "幻灯片设置成功！", "false", "location.replace('Ky_Slide.Asp')")
else
	if Result="Modify" then
		set rs = server.createobject("adodb.recordset")
		sql="select * from keyicms_Slide where ID="& ID
		rs.open sql,conn,1,1
		Url = rs("Url")
		Pic = rs("Pic")
		State = rs("State")
		SlideName = rs("SlideName")
		rs.close : set rs = nothing
	else
		State = 10
	end if
end if
end sub
%>