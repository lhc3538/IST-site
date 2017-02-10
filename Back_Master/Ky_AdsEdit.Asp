<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="../Include/System_Language.asp" -->
<!--#include file="../Include/Version.Asp" -->
<!--#include file="CheckAdmin.Asp"-->
<%
dim Result
Result = request.QueryString("Result")
dim ID, AdsName, AdsPic, AdsUrl, SortName, SortID, SortPath, Content, Sequence
ID = request.QueryString("ID")
Call AdsEdit()
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
	K('#imagePic').click(function() {
		editor.loadPlugin('image', function() {
			editor.plugin.imageDialog({
				showRemote : true,
				imageUrl : K('#AdsPic').val(),
				clickFn : function(url, title, width, height, border, align) {
					K('#AdsPic').val(url);
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
	If InStr(CompanyAdminPurview, "|11,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style=""color:red; font-size:14px;"")>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If
	%>
    <div class="main">
        <div class="row ">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i> <%=Keyicms_Lang_Copyright(14)%></li>
                <li><%=Keyicms_Lang_Copyright(2)%></li>
                <li class="active">广告管理</li>
            </ol>	
        </div>	
        
        <div class="row ">
            <div class="col-lg-12 hideInIE8">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-edit"></i><span class="break"></span><strong>图文管理</strong></h2> 
                    </div>
                    <div class="panel-body">
                        <a class="ShortCut1" href="Ky_AdsEdit.Asp?Result=Add">  
                            <i class="fa fa-inbox"></i>
                            <div class="title">添加图文信息</div>
                        </a>					
                        <a class="ShortCut2" href="Ky_Ads.Asp">  
                            <i class="fa fa-book"></i>
                            <div class="title">图文管理</div>
                        </a>														
                        <a class="ShortCut3" href="Ky_AdsSort.Asp">  
                            <i class="fa fa-sitemap"></i>
                            <div class="title">图文类别</div>
                        </a>														
                    </div>
                </div>
            </div>
        </div>
        <div class="row">
            <div class="col-md-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-gears"></i><span class="break"></span><strong><%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改<%End If%>图文信息</strong></h2>
                    </div>
                    <div class="panel-body">
                    <form name="editForm" method="post" action="?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>" class="form-horizontal">
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">图文标题：</label>
                            <div class="col-md-4">
                                <input name="AdsName" type="text" id="AdsName" class="form-control" value="<%=AdsName%>">
                                <span class="help-block"></span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">图文类别：</label>
                            <div class="col-md-4">
                                <select size="1" class="form-control" name="SortSelect" id="SortSelect" onChange="AddSort()">
                                <option value="" valuepath="">请选择...</option>
                                <%Call AdsSort(0, 0)%>
                                </select>
                                <span class="help-block">
                                    <input name="SortID" type="hidden" id="SortID" value="<%=SortID%>" readonly class="form-control">
                                    <input name="SortPath" type="hidden" id="SortPath" value="<%=SortPath%>" readonly class="form-control">
                                </span>
                            </div>
                        </div>   
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">相关图片：</label>
                            <div class="col-md-4">
                                <div class="input-group">
                                    <input type="text" id="AdsPic" name="AdsPic" class="form-control" value="<%=AdsPic%>" readonly>
                                    <span class="input-group-btn">
                                    <button type="button" id="imagePic" class="btn btn-default">上传图片</button>
                                    </span>
                                </div>
                                <span class="help-block"></span>
                            </div>
                        </div> 
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">链接网址：</label>
                            <div class="col-md-4">
                                <input name="AdsUrl" type="text" id="AdsUrl" class="form-control" value="<%=AdsUrl%>">
                                 <span class="help-block"></span>
                            </div>
                        </div> 
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">简介：</label>
                            <div class="col-md-4">
                                <textarea name="Content" class="form-control"><%=Content%></textarea>
                                <span class="help-block"></span>
                            </div>
                        </div> 
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">排序：</label>
                            <div class="col-md-3">
                                <input name="Sequence" type="text" id="Sequence" class="form-control" value="<%=Sequence%>" onKeyDown="OnKeyDown()" onchange="OnChange(this, <%=Sequence%>)" />
                                <span class="help-block"></span>
                            </div>
                        </div>
                        <div class="form-group">
                            <div class="col-md-2 control-label"></div>
                            <div class="col-md-10">
                                <input name="submitSaveEdit" type="submit" id="submitSaveEdit" class="btn btn-danger btn-lg" value="保存图文信息">
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
sub AdsEdit()
dim Action, rs, sql
Action = request.QueryString("Action")
set rs = server.createobject("adodb.recordset")
if Action = "SaveEdit" then
	If Len(Trim(request.Form("AdsName"))) < 1 Then
		Call SweetAlert("warning", "友情提示", "请填写图文标题！", "false", "history.back(-1)")
		response.End
	End If
	If Request.Form("SortID") = "" and Request.Form("SortPath")="" Then
		Call SweetAlert("warning", "友情提示", "请选择所属分类！", "false", "history.back(-1)")
		response.end
	End If
	If Trim(request.Form("AdsUrl")) = "" Then
		Call SweetAlert("warning", "友情提示", "请填写链接网址！", "false", "history.back(-1)")
		response.End
	End If
	if Result = "Add" then
		sql = "select * from keyicms_Ads"
		rs.open sql, conn, 1, 3
		rs.addnew
		rs("SortID")= Request.Form("SortID")
		rs("SortPath")= Request.Form("SortPath")
		rs("AdsName")= trim(Request.Form("AdsName"))
		rs("AdsPic") = trim(Request.Form("AdsPic"))
		rs("AdsUrl") = trim(Request.Form("AdsUrl"))
		rs("Content") = trim(Request.Form("Content"))
		rs("Sequence") = trim(Request.Form("Sequence"))
		rs.update
		rs.close
	end if
	if Result = "Modify" then
		sql = "select * from keyicms_Ads where ID="&ID
		rs.open sql, conn, 1, 3
		rs("SortID")= Request.Form("SortID")
		rs("SortPath")= Request.Form("SortPath")
		rs("AdsName")= trim(Request.Form("AdsName"))
		rs("AdsPic") = trim(Request.Form("AdsPic"))
		rs("AdsUrl") = trim(Request.Form("AdsUrl"))
		rs("Content") = trim(Request.Form("Content"))
		rs("Sequence") = trim(Request.Form("Sequence"))
		rs.update
		rs.close
	end if
	set rs = nothing
	Call SweetAlert("success", "操作成功", "设置成功！", "false", "location.replace('Ky_Ads.Asp')")
else
	If Result = "Modify" Then
		sql = "select * from keyicms_Ads where ID="& ID
		rs.open sql, conn, 1, 3
		AdsName = rs("AdsName")
		AdsPic = rs("AdsPic")
		AdsUrl = rs("AdsUrl")
		SortName = SortText(rs("SortID"))
		SortID = rs("SortID")
		SortPath = rs("SortPath")
		Content = rs("Content")
		Sequence = rs("Sequence")
		rs.close : set rs = nothing
	else
		Sequence = 10
	end if
end if
set rs = nothing
end sub

Function SortText(ID)
	Dim rs,sql
	set rs = Server.CreateObject("adodb.recordset")
	sql = "Select * From keyicms_AdsSort where ID="&ID
	rs.open sql,conn,1,1
	SortText = rs("SortName")
	rs.close : set rs = nothing
End Function

Function AdsSort(ID, ClassNum)
	dim rs
	set rs = Server.CreateObject("adodb.recordset")
	rs.open "select * from Keyicms_AdsSort where ParentID="&ID&" Order by ID asc",conn,1,2
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
		ClassNum = ClassNum + 1
		Call AdsSort(rs("ID"), ClassNum)
		ClassNum = ClassNum - 1
		rs.movenext
	Loop
	rs.close : set rs = nothing
End Function
%>
