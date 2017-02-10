<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="../Include/Md5.asp"-->
<!--#include file="../Include/System_Language.asp" -->
<!--#include file="../Include/Version.Asp" -->
<!--#include file="CheckAdmin.Asp"-->
<%
'*******************************************************
'软件名称：科蚁企业网站内容管理系统（KEYICMS）
'软件开发：成都神笔天成网络科技有限公司
'网   址：http://www.keyicms.com
'本信息不会影响您网站的正常使用，无论免费用户或是收费用户请保留这里的信息.
'神笔(48)天成网络享有软件著作权，如果您擅自删除这里的版权信息,我们将追究法律责任.
'未经公司同意禁止更改以下代码.
'*******************************************************
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
</head>

<body>
<!--#Include file="Keyicms_Top.Asp"-->
<div id="wrap">
	<!--#Include file="Keyicms_Meun.Asp"-->
	<%
	If InStr(CompanyAdminPurview, "|21,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style=""color:red; font-size:14px;"")>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If

	Dim Result
	Result = request.QueryString("Result")
	Dim ID, AdminName, Working, Password, vPassword, UserName, Purview, Remark, AddTime
	ID = request.QueryString("ID")
	If ID = "" Then ID = 0
	Call AdminEdit()
	%>
    <div class="main">
        <div class="row ">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i><%=Keyicms_Lang_Copyright(14)%></li>
                <li><%=Keyicms_Lang_Copyright(2)%></li>
                <li class="active">管理员列表</li>
            </ol>	
        </div>
        <div class="row ">
            <div class="col-lg-12 hideInIE8">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-edit"></i><span class="break"></span><strong>管理员管理</strong></h2> 
                    </div>
                    <div class="panel-body">
                        <a class="ShortCut1" href="Ky_AdminEdit.Asp?Result=Add">  
                            <i class="fa fa-user"></i>
                            <div class="title">添加管理员</div>
                        </a>					
                        <a class="ShortCut2" href="Ky_Admin.Asp">  
                            <i class="fa fa-users"></i>
                            <div class="title">管理员列表</div>
                        </a>														
                        <a class="ShortCut4" href="Keyicms_ManageLog.Asp">  
                            <i class="fa fa-file-text-o"></i>
                            <div class="title">后台登录日志</div>
                        </a>														
                    </div>
                </div>
            </div>
        </div>
        <div class="row">
            <div class="col-md-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-gears"></i><span class="break"></span><strong><%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改<%End If%>管理员</strong></h2>
                    </div>
                    <form name="editForm" method="post" action="?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>" class="form-horizontal">
                    <div class="panel-body">
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">登录名称：</label>
                            <div class="col-md-4">
                                <input name="AdminName" type="text" id="AdminName" class="form-control" value="<%=AdminName%>" maxlength="16" <%if Result="Modify" then response.write ("readonly")%>>
                                <span class="help-block">3-10个字符</span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">是否生效：</label>
                            <div class="col-md-4">
                                <label class="switch switchtop switch-danger">
                                    <input name="Working" type="checkbox" class="switch-input"<%if Working then response.write ("checked")%> value="1">
                                    <span class="switch-label" data-on="yes" data-off="no"></span>
                                    <span class="switch-handle"></span>
                                </label>
                            </div>
                        </div>   
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">管理员密码：</label>
                            <div class="col-md-4">
                                <input name="Password" type="password" id="Password" maxlength="20" class="form-control">
                                <span class="help-block">6-16个字符（没填写则不修改密码）</span>
                            </div>
                        </div> 
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">确认密码：</label>
                            <div class="col-md-4">
                                <input name="vPassword" type="password" id="vPassword" maxlength="20" class="form-control">
                                <span class="help-block">6-16个字符</span>
                            </div>
                        </div> 
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">管理员名称：</label>
                            <div class="col-md-4">
                                <input name="UserName" type="text" id="UserName" class="form-control" value="<%=UserName%>">
                                <span class="help-block"></span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">操作权限：</label>
                            <div class="col-md-7">
							<%
                            set rs = server.CreateObject("adodb.recordset")
                            rs.open "select * from keyicms_System Order by ID asc" ,conn ,1 ,2
                            do until rs.eof
                            %>
                <label class="checkbox-inline" for="Purview<%=rs("ID")%>">
                    <input id="Purview<%=rs("ID")%>" name="Purview" type="checkbox" value="|<%=rs("ID")%>" <%if Instr(Purview,"|"&rs("ID")&",")>0 then response.write ("checked")%>><%=rs("SystemName")%>
                </label>
                            <%
                                rs.movenext
                            loop
                            rs.close : set rs = nothing
                            %>
                <label class="checkbox-inline" for="Purview7"><input id="Purview7" name="Purview" type="checkbox" value="|7"<%if Instr(Purview,"|7,")>0 then response.write ("checked")%>>（功能模块管理）</label>
                            </div>
                        </div> 
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input"></label>
                            <div class="col-md-7">
                <label class="checkbox-inline" for="Purview8"><input id="Purview8" name="Purview" type="checkbox" value="|8"<%if Instr(Purview,"|8,")>0 then response.write ("checked")%>>百度地图</label>
                <label class="checkbox-inline" for="Purview9"><input id="Purview9" name="Purview" type="checkbox" value="|9"<%if Instr(Purview,"|9,")>0 then response.write ("checked")%>>在线客服</label>
                <label class="checkbox-inline" for="Purview10"><input id="Purview10" name="Purview" type="checkbox" value="|10"<%if Instr(Purview,"|10,")>0 then response.write ("checked")%>>友情链接</label>
                <label class="checkbox-inline" for="Purview11"><input id="Purview11" name="Purview" type="checkbox" value="|11"<%if Instr(Purview,"|11,")>0 then response.write ("checked")%>>图文扩展</label>
                <label class="checkbox-inline" for="Purview12"><input id="Purview12" name="Purview" type="checkbox" value="|12"<%if Instr(Purview,"|12,")>0 then response.write ("checked")%>>幻灯片</label>
                <label class="checkbox-inline" for="Purview13"><input id="Purview13" name="Purview" type="checkbox" value="|13"<%if Instr(Purview,"|13,")>0 then response.write ("checked")%>>留言功能</label>
                <label class="checkbox-inline" for="Purview14"><input id="Purview14" name="Purview" type="checkbox" value="|14"<%if Instr(Purview,"|14,")>0 then response.write ("checked")%>>（企业应用）</label>
                            </div>
                        </div> 
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input"></label>
                            <div class="col-md-7">
                <label class="checkbox-inline" for="Purview15"><input id="Purview15" name="Purview" type="checkbox" value="|15"<%if Instr(Purview,"|15,")>0 then response.write ("checked")%>>网站系统设置</label>
                <label class="checkbox-inline" for="Purview16"><input id="Purview16" name="Purview" type="checkbox" value="|16"<%if Instr(Purview,"|16,")>0 then response.write ("checked")%>>网站参数设置</label>
                <label class="checkbox-inline" for="Purview17"><input id="Purview17" name="Purview" type="checkbox" value="|17"<%if Instr(Purview,"|17,")>0 then response.write ("checked")%>>手机网站设置</label>
                <label class="checkbox-inline" for="Purview18"><input id="Purview18" name="Purview" type="checkbox" value="|18"<%if Instr(Purview,"|18,")>0 then response.write ("checked")%>>网站导航设置</label>
                <label class="checkbox-inline" for="Purview19"><input id="Purview19" name="Purview" type="checkbox" value="|19"<%if Instr(Purview,"|19,")>0 then response.write ("checked")%>>模板列表</label>
                            </div>
                        </div> 
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input"></label>
                            <div class="col-md-7">
                <label class="checkbox-inline" for="Purview20"><input id="Purview20" name="Purview" type="checkbox" value="|20"<%if Instr(Purview,"|20,")>0 then response.write ("checked")%>>生成静态</label>
                <label class="checkbox-inline" for="Purview21"><input id="Purview21" name="Purview" type="checkbox" value="|21"<%if Instr(Purview,"|21,")>0 then response.write ("checked")%>>网站管理员管理</label>
                <label class="checkbox-inline" for="Purview22"><input id="Purview22" name="Purview" type="checkbox" value="|22"<%if Instr(Purview,"|22,")>0 then response.write ("checked")%>>后台登陆日志</label>
                            </div>
                        </div> 
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input"></label>
                            <div class="col-md-7">
                                <button type="button" id="OtherImage" class="btn btn-default btn-xs" onClick="CheckPurview()">全选</button>
                                <button type="button" id="OtherImage" class="btn btn-default btn-xs" onClick="CheckPurviewO()">反选</button>
                            </div>
                        </div> 
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">备注：</label>
                            <div class="col-md-4">
                            	<textarea name="Remark" rows="8" cols="80" id="Remark" class="form-control"><%=Remark%></textarea>
                            </div>
                        </div> 
                        <div class="form-group">
                            <div class="col-md-2 control-label"></div>
                            <div class="col-md-10">
                                <input name="submitSaveEdit" type="submit" id="submitSaveEdit" class="btn btn-danger btn-lg" value="保存管理员信息">
                                <input type="button" class="btn btn-danger btn-lg" value="<%=Keyicms_Lang_System(32)%>" onClick="history.back(-1)">
                            </div>	
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
sub AdminEdit()
	dim Action, rsCheckAdd, rs, sql
	Action = request.QueryString("Action")
	if Action="SaveEdit" then
		if len(trim(Request.Form("AdminName")))<3 or len(trim(Request.Form("AdminName")))>10  then
			Call SweetAlert("warning", "友情提示", "请填写管理员名称（字符数在3-10位之间）！", "false", "history.back(-1)")
			response.end
		end if
		If(Result="Modify" and Request.Form("Password")<>"") or Result="Add" Then
			if len(trim(Request.Form("Password")))<6 or len(trim(Request.Form("Password")))>16  then
				Call SweetAlert("warning", "友情提示", "请填写管理员密码（字符数在6-16位之间）！", "false", "history.back(-1)")
				response.end
			end if
		End If
		if Request.Form("Password")<>Request.Form("vPassword") then
			Call SweetAlert("warning", "友情提示", "两次输入的密码不同！", "false", "history.back(-1)")
			response.end
		end if
		if Request.Form("Purview")="" then
			Call SweetAlert("warning", "友情提示", "你还未选中任何一个权限！", "false", "history.back(-1)")
			response.end
		end if

		if Result="Add" then
			set rsCheckAdd = conn.execute("select AdminName from keyicms_Admin where AdminName='" & trim(Request.Form("AdminName")) & "'")
			if not (rsCheckAdd.bof and rsCheckAdd.eof) then
				Call SweetAlert("warning", "友情提示", trim(Request.Form("AdminName")) & "管理员名称已存在！", "false", "history.back(-1)")
				response.end
			end if
		
			set rs = server.createobject("adodb.recordset")
			sql = "select * from keyicms_Admin"
			rs.open sql,conn,1,3
			rs.addnew
			rs("AdminName") = trim(Request.Form("AdminName"))
			if Request.Form("Working") = 1 then
				rs("Working") = Request.Form("Working")
			else
				rs("Working") = 0
			end if
			rs("Password") = Md5(Request.Form("Password"))
			rs("UserName") = trim(Request.Form("UserName"))
			rs("AdminPurview") = replace(Request.Form("Purview"), " ", "")& ","
			rs("Remark") = trim(Request.Form("Remark"))
			rs("AddTime") = now()
		end if
		if Result="Modify" then
			set rs = server.createobject("adodb.recordset")
			sql = "select * from keyicms_Admin where ID="&ID
			rs.open sql,conn,1,3
			if Request.Form("Working") = 1 then
				rs("Working") = Request.Form("Working")
			else
				rs("Working") = 0
			end if
			If trim(Request.Form("Password")) <> "" Then
				rs("Password") = Md5(Request.Form("Password"))
			End If
			rs("UserName") = trim(Request.Form("UserName"))
			rs("AdminPurview") = replace(Request.Form("Purview"), " ", "")& ","
			rs("Remark") = trim(Request.Form("Remark"))
		end if
		rs.update
		rs.close : set rs = nothing
		Call SweetAlert("success", "操作成功", "设置成功！", "false", "location.replace('Ky_Admin.Asp')")
	else
		if Result="Modify" then
			set rs = server.createobject("adodb.recordset")
			sql = "select * from keyicms_Admin where ID="& ID
			rs.open sql,conn,1,1
			AdminName = rs("AdminName")
			Working = rs("Working")
			UserName = rs("UserName")
			Purview = rs("AdminPurview")
			Remark = rs("Remark")
			rs.close : set rs = nothing
		end if
	end if
end sub
%>