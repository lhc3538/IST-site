<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="../Include/system_language.asp" -->
<!--#include file="CheckAdmin.Asp"-->
<!--#include file="../Html/html_Config.Asp"-->
<!--#include file="../Html/html_Function.Asp"-->
<!--#include file="../Html/html_To_Download.Asp"-->
<%

Dim Result
Result = request.QueryString("Result")
Dim ID, DownName, ClassSeo, ViewFlag, MobileFlag, SortName, SortID, SortPath, FileUrl, FileSize
Dim Content, SeoKeywords, SeoDescription, ClickNumber
ID = request.QueryString("ID")
Call DownEdit()
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
<script charset="utf-8" src="KEditor_keyicms/kindeditor.js"></script>
<script charset="utf-8" src="KEditor_keyicms/lang/zh_CN.js"></script>
<script charset="utf-8" src="KEditor_keyicms/plugins/code/prettify.js"></script>
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
	K('#InsertFile').click(function() {
		editor.loadPlugin('insertfile', function() {
			editor.plugin.fileDialog({
				fileUrl : K('#FileUrl').val(),
				clickFn : function(url, title) {
					K('#FileUrl').val(url);
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
	If InStr(CompanyAdminPurview, "|5,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style=""color:red; font-size:14px;"")>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If
	%>
    <div class="main">
        <div class="row">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i> <%=Keyicms_Lang_Copyright(14)%></li>
                <li class="active"><a href="">功能模块</a></li>
                <li class="active">下载管理</li>
            </ol>	
        </div>	
        <div class="row">
            <div class="col-lg-12 hideInIE8">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-edit"></i><span class="break"></span><strong>下载管理</strong></h2> 
                    </div>
                    <div class="panel-body">
                        <a class="ShortCut1" href="Ky_DownEdit.Asp?Result=Add">  
                            <i class="fa fa-download"></i>
                            <div class="title">添加下载</div>
                        </a>					
                        <a class="ShortCut2" href="Ky_Download.Asp">  
                            <i class="fa fa-th-list"></i>
                            <div class="title">下载列表</div>
                        </a>	
                        <a class="ShortCut3" href="Ky_DownSort.Asp">  
                            <i class="fa fa-table"></i>
                            <div class="title">下载分类</div>
                        </a>													
                    </div>
                </div>
            </div>
        </div>
        <div class="row">
            <div class="col-md-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-gears"></i><span class="break"></span><strong><%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改<%End If%>下载信息</strong></h2>
                    </div>
                    <div class="panel-body">
                    <form name="editForm" method="post" action="?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>" class="form-horizontal">
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">下载名称：</label>
                            <div class="col-md-4">
                                <input name="DownName" type="text" id="DownName" class="form-control" value="<%=DownName%>" maxlength="100">
                                <label class="checkbox-inline" for="ViewFlag">
                                    <input type="checkbox" id="ViewFlag" name="ViewFlag" value="1" <%if ViewFlag then response.write ("checked")%>> 是否生效
                                </label>                           
                                <%If mStatus Then%>
                                <label class="checkbox-inline" for="MobileFlag">
                                    <input type="checkbox" id="MobileFlag" name="MobileFlag" value="1" <%if MobileFlag then response.write ("checked")%>> 手机显示
                                </label>
                                <%End If%>
                            </div>
                        </div> 
  
                        
                        <%If EnableSeo = True Then%>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">关键词：</label>
                            <div class="col-md-6">
                            	<input name="SeoKeywords" type="text" id="SeoKeywords" class="form-control" value="<%=SeoKeywords%>" maxlength="250">
                                <span class="help-block">SEO：多个关键词请用“,”半角逗号分开</span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">描述：</label>
                            <div class="col-md-6">
                            	<input name="SeoDescription" type="text" id="SeoDescription" class="form-control" value="<%=SeoDescription%>" maxlength="250">
                                <span class="help-block">SEO：请填写描述</span>
                            </div>
                        </div>
                       <%End If%> 
                       <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">下载类别：</label>
                            <div class="col-md-5">
                                <select size="1" class="form-control" name="SortSelect" id="SortSelect" onChange="AddSort()">
                                <option value="" valuepath="">请选择...</option>
                                <%Call DownSort(0, 0)%>
                                </select>
                                <span class="help-block">
                                    <input name="SortID" type="hidden" id="SortID" value="<%=SortID%>" readonly class="form-control">
                                    <input name="SortPath" type="hidden" id="SortPath" value="<%=SortPath%>" readonly class="form-control">
                                </span>
                            </div>
                        </div> 
                        <div class="form-group">
							<label class="col-md-2 control-label" for="text-input">下载地址：</label>
                        <div class="col-md-5"> 
                            <div class="input-group">
                                <input name="FileUrl" type="text" id="FileUrl" class="form-control" value="<%=FileUrl%>" size="45" readonly>
                                <span class="input-group-btn">
                                <button type="button" id="InsertFile" class="btn btn-default">选择文件</button>
                                </span>
                                </div>
                                <span class="help-block">请上传相应文件</span>
                            </div>
                        </div>   
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">文件大小：</label>
                            <div class="col-md-5">
                                <div class="input-group">
                                    <input type="text" id="FileSize" name="FileSize" class="form-control" value="<%=FileSize%>" maxlength="100">
                                    <span class="input-group-addon">MB</span>
                                </div>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">内容：</label>
                            <div class="col-md-10">
                                <textarea name="Content" style="width:95%;height:400px;visibility:hidden;" class="form-control"><%=Server.HTMLEncode(Content)%></textarea>
                                 <span class="help-block"></span>
                            </div>
                        </div>
                        <%If EnableViews = True Then%>  
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">浏览数：</label>
                            <div class="col-md-3">
                                <input type="text" name="ClickNumber" id="ClickNumber" class="form-control" value="<%=ClickNumber%>" onKeyDown="OnKeyDown()" onChange="OnChange(this, <%=ClickNumber%>)" />
                                <span class="help-block">可以手动输入浏览数量。</span>
                            </div>
                        </div> 
                        <%End If%>
                        <div class="form-group">
                            <div class="col-md-2 control-label"></div>
                            <div class="col-md-10">
                                <input name="submitSaveEdit" type="submit" id="submitSaveEdit" class="btn btn-danger btn-lg" value="保存下载信息">
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
Sub DownEdit()
Dim Action, rs, sql
Action = request.QueryString("Action")
If Action = "SaveEdit" Then
	Set rs = server.CreateObject("adodb.recordset")
	If Len(Trim(request.Form("DownName")))<3 Then
		Call SweetAlert("warning", "友情提示", "请填写下载名称！", "false", "history.back(-1)")
		response.End
	End If
	if Request.Form("SortID") = "" And Request.Form("SortPath") = "" then
		Call SweetAlert("warning", "友情提示", "请选择类别！", "false", "history.back(-1)")
		response.end
	end if
	If Trim(request.Form("FileUrl")) = "" Then
		Call SweetAlert("warning", "友情提示", "请上传下载文件！", "false", "history.back(-1)")
		response.End
	End If
	If Trim(request.Form("FileSize")) = "" Then
		Call SweetAlert("warning", "友情提示", "请填写文件大小！", "false", "history.back(-1)")
		response.End
	End If
	If Trim(request.Form("Content")) = "" Then
		Call SweetAlert("warning", "友情提示", "请填写内容！", "false", "history.back(-1)")
		response.End
	End If

	If Result = "Add" Then
		sql = "select * from keyicms_Download"
		rs.Open sql, conn, 1, 3
		rs.addnew
		rs("DownName") = Trim(Request.Form("DownName"))
		If Request.Form("ViewFlag") = 1 Then
			rs("ViewFlag") = Request.Form("ViewFlag")
		Else
			rs("ViewFlag") = 0
		End If
		if Request.Form("MobileFlag") = 1 then
			rs("MobileFlag") = Request.Form("MobileFlag")
		else
			rs("MobileFlag") = 0
		end if
		If Request.Form("ClassSeo") = "" then
			If Request.Form("oAutopinyin") = "1" then rs("ClassSeo") = TimeToName()
			If Request.Form("oAutopinyin") = "2" then rs("ClassSeo") = Left(Pinyin(trim(request.form("DownName"))),200)
		Else
			rs("ClassSeo") = trim(Request.form("ClassSeo"))
		End If
		rs("SortID") = Request.Form("SortID")
		rs("SortPath") = Request.Form("SortPath")
		rs("FileUrl") =  Request.Form("FileUrl")
		rs("FileSize") =  Request.Form("FileSize")
		rs("Content") = RTrim(Request.Form("Content"))
		rs("SeoKeywords") = Trim(Request.Form("SeoKeywords"))
		rs("SeoDescription") = Trim(Request.Form("SeoDescription"))
		rs("AddTime") = Now()
		rs("UpdateTime") = Now()

		If EnableViews Then rs("ClickNumber") = trim(Request.Form("ClickNumber"))
		rs.update
		rs.close : set rs = nothing
		
		Set rs = server.CreateObject("adodb.recordset")
		sql = "select top 1 ID,ClassSeo from keyicms_Download order by ID desc"
		rs.Open sql, conn, 1, 1
		ID = rs("ID")
		rs.close : set rs = nothing
			
		'Call Html_To_Download(ID, "PC")
		If mStatus Then Call Html_To_Download(ID, "Mobile")
	End If
	If Result = "Modify" Then
		sql = "select * from keyicms_Download where ID="&ID
		rs.Open sql, conn, 1, 3
		rs("DownName") = Trim(Request.Form("DownName"))
		If Request.Form("ViewFlag") = 1 Then
			rs("ViewFlag") = Request.Form("ViewFlag")
		Else
			rs("ViewFlag") = 0
		End If
		if Request.Form("MobileFlag") = 1 then
			rs("MobileFlag") = Request.Form("MobileFlag")
		else
			rs("MobileFlag") = 0
		end if
		rs("SortID") = Request.Form("SortID")
		rs("SortPath") = Request.Form("SortPath")
		rs("FileUrl") =  Request.Form("FileUrl")
		rs("FileSize") =  Request.Form("FileSize")
		rs("Content") = RTrim(Request.Form("Content"))
		rs("SeoKeywords") = Trim(Request.Form("SeoKeywords"))
		rs("SeoDescription") = Trim(Request.Form("SeoDescription"))
		rs("UpdateTime") = Now()

		If EnableViews Then rs("ClickNumber") = trim(Request.Form("ClickNumber"))
		ID = rs("ID")
		rs.update
		rs.close : set rs = nothing

		'Call Html_To_Download(ID, "PC")
		If mStatus Then Call Html_To_Download(ID, "Mobile")
	End If
	Call SweetAlert("success", "操作成功", "设置成功，相关静态页面已更新！", "false", "location.replace('Ky_Download.Asp')")
Else
	If Result = "Modify" Then
		Set rs = server.CreateObject("adodb.recordset")
		sql = "select * from keyicms_Download where ID="& ID
		rs.Open sql, conn, 1, 1
		If rs.bof And rs.EOF Then
			response.Write ("数据库读取记录出错！")
			response.End
		End If
		DownName = rs("DownName")
		ViewFlag = rs("ViewFlag")
		MobileFlag = rs("MobileFlag")
		ClassSeo = rs("ClassSeo")&Separated&rs("id")
		SortID = rs("SortID")
		SortPath = rs("SortPath")
		GroupID = rs("GroupID")
		Exclusive = rs("Exclusive")
		Content = rs("Content")
		SeoKeywords = rs("SeoKeywords")
		SeoDescription = rs("SeoDescription")
		ClickNumber = rs("ClickNumber")
		FileUrl = rs("FileUrl")
		FileSize = rs("FileSize")
		rs.close : set rs = nothing
	Else
		ClickNumber = 10
	End If
End If
End Sub

Function DownSort(ID, ClassNum)
	Dim rs, sql
	set rs = Server.CreateObject("adodb.recordset")
	rs.open "select * from Keyicms_DownSort where ParentID="&ID&" Order by ID asc",conn,1,2
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
		Call DownSort(rs("ID"), ClassNum)
		ClassNum = ClassNum - 1
		rs.movenext
	Loop
	rs.close
	set rs = nothing
End Function
%>