<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="../Include/System_Language.asp" -->
<!--#include file="../Include/Version.Asp" -->
<!--#include file="CheckAdmin.Asp"-->
<!--#include file="../Html/html_Config.Asp"-->
<!--#include file="../Html/html_Function.Asp"-->
<!--#include file="../Html/html_To_Job.Asp"-->
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
dim ID, JobName, ViewFlag, MobileFlag, ClassSeo, JobAddress, JobNumber, Emolument, StartDate, EndDate, Requirement, ClickNumber
ID = request.QueryString("ID")
Call JobEdit()
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
<script charset="utf-8" src="KEditor_keyicms/kindeditor-min.js"></script>
<script charset="utf-8" src="KEditor_keyicms/lang/zh_CN.js"></script>
<script charset="utf-8" src="KEditor_keyicms/plugins/code/prettify.js"></script>
<script>
KindEditor.ready(function(K) {
	var editor = K.create('textarea[name="Requirement"]', {
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
	prettyPrint();

});
</script>
</head>

<body>
<!--#Include file="Keyicms_Top.Asp"-->
<div id="wrap">
	<!--#Include file="Keyicms_Meun.Asp"-->
 	<%
	If InStr(CompanyAdminPurview, "|6,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style=""color:red; font-size:14px;"")>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If
	%>
    <div class="main">
        <div class="row">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i> 当前位置</li>
                <li class="active"><a href="">功能模块</a></li>
                <%
				If SortPath<>"" Then
					Response.Write "<li>招聘信息列表 -> <a href='NewsList.Asp'>全部</a></li>"
					TextPath(SortID)
				Else
					Response.Write "<li class='active'>招聘信息列表 -> 全部</li>"
				End If
				%>
            </ol>	
        </div>	
        <div class="row ">
            <div class="col-lg-12 hideInIE8">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-edit"></i><span class="break"></span><strong>招聘管理</strong></h2> 
                    </div>
                    <div class="panel-body">
                        <a class="ShortCut1" href="Ky_JobEdit.Asp?Result=Add">  
                            <i class="fa fa-indent"></i>
                            <div class="title">添加招聘信息</div>
                        </a>					
                        <a class="ShortCut2" href="Ky_Job.Asp">  
                            <i class="fa fa-list"></i>
                            <div class="title">招聘信息列表</div>
                        </a>					
                    </div>
                </div>
            </div>
        </div> 
        <div class="row">
            <div class="col-md-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-gears"></i><span class="break"></span><strong><%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改<%End If%>招聘信息</strong></h2>
                    </div>
                    <div class="panel-body">
                    <form name="editForm" method="post" action="?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>" class="form-horizontal"role="form">
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">招聘名称：</label>
                            <div class="col-md-4">
                                <input name="JobName" type="text" id="JobName" class="form-control" value="<%=JobName%>" maxlength="100">
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
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">工作地点：</label>
                            <div class="col-md-5">
                                <input name="JobAddress" type="text" id="JobAddress" class="form-control" value="<%=JobAddress%>" maxlength="100">
                                <span class="help-block"></span>
                            </div>
                        </div> 
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">招聘人数：</label>
                            <div class="col-md-5">
                                <input name="JobNumber" type="text" id="JobNumber" class="form-control" value="<%=JobNumber%>" maxlength="100">
                                <span class="help-block"></span>
                            </div>
                        </div> 
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">薪资待遇：</label>
                            <div class="col-md-5">
                                <input name="Emolument" type="text" id="Emolument" class="form-control" value="<%=Emolument%>" maxlength="100">
                                <span class="help-block"></span>
                            </div>
                        </div> 
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">有效时间：</label>
                            <div class="col-md-2">
                                <input name="StartDate" type="text" id="StartDate" class="form-control" value="<% if StartDate="" then response.write now() else response.write (StartDate) end if%>" maxlength="18">
                                <span class="help-block">默认时间为3个月，可手动修改。</span>
                            </div>
                            <label class="col-md-1 control-label" style="text-align:center">→</label>
                            <div class="col-md-2">
                                <input name="EndDate" type="text" id="EndDate" class="form-control" value="<% if EndDate="" then response.write (DateAdd("m",3,now())) else response.write (EndDate) end if%>" maxlength="18">
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">职位要求：</label>
                            <div class="col-md-10">
                                <textarea name="Requirement" class="form-control" style="width:95%;height:400px;visibility:hidden;"><%=Server.HTMLEncode(Requirement)%></textarea>
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
                                <input name="submitSaveEdit" type="submit" id="submitSaveEdit" class="btn btn-danger btn-lg" value="保存招聘信息">
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
<script src="ke01/plugins/pace/pace.min.js"></script>
<script src="ke01/js/jquery.mmenu.min.js"></script>
<script src="ke01/js/core.min.js"></script>
<script src="ke01/plugins/jquery-cookie/jquery.cookie.min.js"></script>
<script src="ke01/js/demo.min.js"></script></body>
</html>
<%
Sub JobEdit()
Dim Action, rsCheckAdd, rs, sql
Action = request.QueryString("Action")
If Action = "SaveEdit" Then
	Set rs = server.CreateObject("adodb.recordset")
	If Len(Trim(request.Form("JobName")))<1 Then
		Call SweetAlert("warning", "友情提示", "请填写招聘职位名称！", "false", "history.back(-1)")
		response.End
	End If
	If Len(Trim(request.Form("JobAddress")))<1 Or Len(Trim(request.Form("Emolument")))<1 Then
		Call SweetAlert("warning", "友情提示", "请填写工作地点、薪水待遇！", "false", "history.back(-1)")
		response.End
	End If
	If Not IsNumeric(Trim(request.Form("JobNumber"))) Then
		Call SweetAlert("warning", "友情提示", "请正确填写职位数量！", "false", "history.back(-1)")
		response.End
	End If
	If Not (IsDate(Trim(request.Form("StartDate"))) Or IsDate(Trim(request.Form("EndDate")))) Then
		Call SweetAlert("warning", "友情提示", "请正确填写开始、结束日期！", "false", "history.back(-1)")
		response.End
	End If
	If Trim(request.Form("Requirement")) = "" Then
		Call SweetAlert("warning", "友情提示", "请填写职位要求！", "false", "history.back(-1)")
		response.End
	End If

	If Result = "Add" Then
		sql = "select * from keyicms_Job"
		rs.Open sql, conn, 1, 3
		rs.addnew
		rs("JobName") = Trim(Request.Form("JobName"))
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
		rs("ClassSeo") = TimeToName()
		rs("JobAddress") = Trim(Request.Form("JobAddress"))
		rs("JobNumber") = Trim(Request.Form("JobNumber"))
		rs("Emolument") = Trim(Request.Form("Emolument"))
		rs("StartDate") = Trim(Request.Form("StartDate"))
		rs("EndDate") = Trim(Request.Form("EndDate"))
		rs("Requirement") = Request.Form("Requirement")
		rs("AddTime") = Now()
		rs("UpdateTime") = Now()

		If EnableViews Then rs("ClickNumber") = trim(Request.Form("ClickNumber"))
		rs.update
		rs.close : set rs = nothing
		
		Set rs = server.CreateObject("adodb.recordset")
		sql = "select top 1 ID,ClassSeo from keyicms_Job order by ID desc"
		rs.Open sql, conn, 1, 1
		ID = rs("ID")
		rs.close : set rs = nothing
		
		Call Html_To_Job(ID, "PC")
		If mStatus Then Call Html_To_Job(ID, "Mobile")
	End If
	If Result = "Modify" Then
		sql = "select * from keyicms_Job where ID="&ID
		rs.Open sql, conn, 1, 3
		rs("JobName") = Trim(Request.Form("JobName"))
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
		rs("JobAddress") = Trim(Request.Form("JobAddress"))
		rs("JobNumber") = Trim(Request.Form("JobNumber"))
		rs("Emolument") = Trim(Request.Form("Emolument"))
		rs("StartDate") = Trim(Request.Form("StartDate"))
		rs("EndDate") = Trim(Request.Form("EndDate"))
		rs("Requirement") = Request.Form("Requirement")
		rs("UpdateTime") = Now()

		If EnableViews Then rs("ClickNumber") = trim(Request.Form("ClickNumber"))
		ID = rs("ID")
		rs.update
		rs.close : set rs = nothing
		
		Call Html_To_Job(ID, "PC")
		If mStatus Then Call Html_To_Job(ID, "Mobile")
	End If
	Call SweetAlert("success", "操作成功", "设置成功，相关静态页面已更新！", "false", "location.replace('Ky_Job.Asp')")
Else
	If Result = "Modify" Then
		Set rs = server.CreateObject("adodb.recordset")
		sql = "select * from keyicms_Job where ID="& ID
		rs.Open sql, conn, 1, 1
		JobName = rs("JobName")
		ViewFlag = rs("ViewFlag")
		MobileFlag = rs("MobileFlag")
		JobAddress = rs("JobAddress")
		JobNumber = rs("JobNumber")
		Emolument = rs("Emolument")
		StartDate = rs("StartDate")
		EndDate = rs("EndDate")
		Requirement = rs("Requirement")
		ClickNumber = rs("ClickNumber")
		rs.close : set rs = nothing
	Else
		ClickNumber = 10
	End If
End If
End Sub
%>