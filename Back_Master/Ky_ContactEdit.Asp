<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="../Include/system_language.asp" -->
<!--#include file="CheckAdmin.Asp"-->
<%
dim Result
Result = request.QueryString("Result")
dim ID, User, Qq, Phone, WeChat, Sequence 
ID = request.QueryString("ID")
call ContactEdit()
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
        <div class="row ">
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
            <div class="col-md-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-gears"></i><span class="break"></span><strong><%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改<%End If%>内置客服</strong></h2>
                    </div>
                    <div class="panel-body">
                    <form name="editForm" method="post" action="?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>" class="form-horizontal">
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">用户名：</label>
                            <div class="col-md-3">
                                <input name="User" type="text" id="User" class="form-control" value="<%=User%>">
                                <span class="help-block"></span>
                            </div>
                        </div> 
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">QQ：</label>
                            <div class="col-md-3">
                                <input name="Qq" type="text" id="Qq" class="form-control" value="<%=Qq%>">
                                <span class="help-block"></span>
                            </div>
                        </div> 
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">手机：</label>
                            <div class="col-md-3">
                                <input name="Phone" type="text" id="Phone" class="form-control" value="<%=Phone%>">
                                <span class="help-block"></span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">微信：</label>
                            <div class="col-md-3">
                                <input name="WeChat" type="text" id="WeChat" class="form-control" value="<%=WeChat%>">
                                <span class="help-block"></span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">排序：</label>
                            <div class="col-md-3">
                                <input name="Sequence" type="text" id="Sequence" class="form-control" value="<%=Sequence%>" onKeyDown="OnKeyDown()" onchange="OnChange(this, <%=Sequence%>)" />
                                <span class="help-block">数字越小排名越靠前。</span>
                            </div>
                        </div>
                        <div class="form-group">
                            <div class="col-md-10">
                                <input name="submitSaveEdit" type="submit" id="submitSaveEdit" class="btn btn-danger btn-lg" value="保存在线客服">
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
sub ContactEdit()
dim Action, rs, sql
Action = request.QueryString("Action")
if Action = "SaveEdit" then
    set rs = server.createobject("adodb.recordset")
    if Result = "Add" then
		If request.Form("User") = "" Then
			Call SweetAlert("warning", "友情提示", "请填写用户名！", "false", "history.back(-1)")
			response.end
		End If

		sql = "select * from keyicms_Contact"
		rs.open sql, conn, 1, 2
		rs.addnew
		rs("User") = trim(Request.Form("User"))
		rs("Qq") = trim(Request.Form("Qq"))
		rs("Phone") = trim(Request.Form("Phone"))
		rs("WeChat") = trim(Request.Form("WeChat"))
		rs("Sequence") = trim(Request.Form("Sequence"))
		rs("SortID") = 1
		rs.update
		rs.close
	end if
	if Result="Modify" then
		arr = split(replace(Request.Form("ID")," ", ""), ",")
		For i = 0 to Ubound(arr)
			sql = "select * from keyicms_Contact where ID="&arr(i)
			rs.open sql, conn, 1, 2
			rs("User") = trim(Request.Form("User_"&arr(i)))
			rs("Qq") = trim(Request.Form("Qq_"&arr(i)))
			rs("Phone") = trim(Request.Form("Phone_"&arr(i)))
			rs("WeChat") = trim(Request.Form("WeChat_"&arr(i)))
			If trim(Request.Form("Sequence_"&arr(i))) = "" Then
				rs("Sequence") = 0
			Else
				rs("Sequence") = trim(Request.Form("Sequence_"&arr(i)))
			End If
			rs.update
			rs.close
		Next
		
		sql = "select top 1 * from keyicms_Contact where SortID=1"
		rs.open sql, conn, 1, 2
		If Trim(Request.Form("Status1")) = "-1" Then
			rs("Status") = Trim(Request.Form("Status1"))
		Else
			rs("Status") = 0
		End If
		rs.update
		rs.close

	end if
	set rs = nothing
	Call SweetAlert("success", "操作成功", "设置成功！", "false", "location.replace('Ky_Contact.Asp')")
else
	if Result="Modify" then
		set rs = server.createobject("adodb.recordset")
		sql="select * from keyicms_Contact where ID="& ID
		rs.open sql,conn,1,1
		User = rs("User")
		Qq = rs("Qq")
		Phone = rs("Phone")
		WeChat = rs("WeChat")
		Sequence = rs("Sequence")
		rs.close : set rs = nothing
	else
		Sequence = 10
	end if
end if
end sub
%>