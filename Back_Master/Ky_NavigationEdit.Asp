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
'神笔(4)天成网络享有软件著作权，如果您擅自删除这里的版权信息,我们将追究法律责任.
'未经公司同意禁止更改以下代码.
'*******************************************************
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
	If InStr(CompanyAdminPurview, "|18,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style=""color:red; font-size:14px;"")>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If
	
	Dim Result
	Result = request.QueryString("Result")
	Dim ID, NavName, ViewFlag, MobileFlag, NavType, HtmlNavUrl, OutFlag, OutUrl, SelectClass, SelectID, Sequence, NavParentID
	ID = request.QueryString("ID")
	If ID <> "" Then ID = CInt(ID)
	
	Call SiteInfo()
	Call NavEdit()
	%>
    <div class="main">
        <div class="row ">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i> 当前位置</li>
                <li class="active">网站导航管理</li>
            </ol>	
        </div>	
        <div class="row">
            <div class="col-lg-12 hideInIE8">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-edit"></i><span class="break"></span><strong>网站导航管理</strong></h2> 
                    </div>
                    <div class="panel-body">
                        <a class="ShortCut1" href="Ky_NavigationEdit.Asp?Result=Add">  
                            <i class="fa fa-indent"></i>
                            <div class="title">添加导航</div>
                        </a>					
                        <a class="ShortCut2" href="Ky_Navigation.Asp">  
                            <i class="fa fa-th-list"></i>
                            <div class="title">导航管理</div>
                        </a>													
                    </div>
                </div>
            </div>
        </div>
        <div class="row">
            <div class="col-md-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-gears"></i><span class="break"></span><strong><%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改<%End If%>网站导航</strong></h2>
                    </div>
                    <div class="panel-body">
                    <form action="?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>" method="post" class="form-horizontal ">
                    	<div class="form-group">
                            <label class="col-md-2 control-label">导航名称：</label>
                            <div class="col-md-5">
                                <input name="NavName" type="text" id="NavName"  class="form-control" value="<%=NavName%>">
                                <label class="checkbox-inline" for="ViewFlag">
                                    <input type="checkbox" value="1" name="ViewFlag" id="ViewFlag" <%if ViewFlag then response.write ("checked")%>> 生效
                                </label>
                                <%If mStatus Then%><label class="checkbox-inline" for="MobileFlag">
                                    <input type="checkbox" value="1" name="MobileFlag" id="MobileFlag" <%if MobileFlag then response.write ("checked")%>> 手机显示
                                </label><%End If%>
                            </div>
                        </div>
                        <div class="form-group">
                        	<label class="col-md-2 control-label">位置：</label>
                            <div class="col-md-5">
                                <label for="NavType1" class="radio-inline">
                                    <input type="radio" value="1" name="NavType" id="NavType1" data-name="TabData" data-ch="NavType" data-value="1" <%if NavType = 1 then response.write ("checked")%>> 主导航
                                </label>
                                <label for="NavType2" class="radio-inline">
                                    <input type="radio" value="2" name="NavType" id="NavType2" data-name="TabData" data-ch="NavType" data-value="2" <%if NavType = 2 then response.write ("checked")%>> 底部导航
                                </label>
                            </div>
                        </div>
                        <div class="form-group" >
                        	<label class="col-md-2 control-label">上级导航：</label>
                            <div class="col-md-5">
                                <select id="select" name="ParentID1" class="form-control" data-div='NavType1' <%If NavType <> 1 Then response.Write "style='display:none'"%>>
                                <option value="0" <%if NavParentID = 0 then response.write ("selected")%>>顶级导航</option>
                                <%Call Navigation1(0, 0)%>
                                </select>

                                <select id="select" name="ParentID2" class="form-control" data-div='NavType2' <%If NavType <> 2 Then response.Write "style='display:none'"%>>
                                <option value="0" <%if NavParentID = 0 then response.write ("selected")%>>顶级导航</option>
                                <%Call Navigation2(0, 0)%>
                                </select>
                                <span class="help-block"></span>
                            </div>
                        </div>

                        <div class="form-group">
                            <label class="col-md-2 control-label">链接状态：</label>
                            <div class="col-md-9">
                                <label for="inline-radio2" class="radio-inline">
                                    <input type="radio" value="0" name="OutFlag" id="inline-radio2" data-name="TabData" data-ch="Out" data-value="2" <%if not OutFlag then response.write ("checked")%>> 内部链接
                                </label>
                                <label for="inline-radio1" class="radio-inline">
                                    <input type="radio" value="1" name="OutFlag" id="inline-radio1" data-name="TabData" data-ch="Out" data-value="1" <%if OutFlag then response.write ("checked")%>> 外部链接
                                </label>
                                <span class="help-block"></span>
                            </div>
                        </div>
                        <div class="form-group" data-div="Out2" <%If OutFlag Then response.Write "style='display:none'"%>>
                            <label class="col-md-2 control-label" for="select">导航选择</label>
                            <div class="col-md-5">
                                <select id="select" name="SelectClass" class="form-control" size="1">
                                <option value="">请选择...</option>
                                <option value="In_0" data="/" <%if SelectClass = "In" and SelectID = 0 then response.write ("selected")%>>首页</option>
                                <%If Conn.execute("select * from keyicms_System where DataForm=1")("ViewFlag") Then response.Write NavAbout()%>
                                <%If Conn.execute("select * from keyicms_System where DataForm=2")("ViewFlag") Then response.Write NavOne("News", 0)%>
                                <%If Conn.execute("select * from keyicms_System where DataForm=3")("ViewFlag") Then response.Write NavOne("Product", 0)%>
                                <%If Conn.execute("select * from keyicms_System where DataForm=4")("ViewFlag") Then response.Write NavOne("Case", 0)%>
                                <%If Conn.execute("select * from keyicms_System where DataForm=5")("ViewFlag") Then response.Write NavOne("Download", 0)%>
                                
                                <%If Conn.execute("select * from keyicms_System where DataForm=6")("ViewFlag") Then%>
                                <option value="Jo_0" data="/<%=JobDir%>" <%if SelectClass = "Jo" and SelectID = 0 then response.write ("selected")%>>
                                <%=Conn.execute("select * from keyicms_System where DataForm=6")("SystemName")%>
                                </option>
                                <%End If%>
                                
                                <option value="Co_0" data="/Contact" <%if SelectClass = "Co" and SelectID = 0 then response.write ("selected")%>>联系</option>
                                <option value="Me_0" data="/Message" <%if SelectClass = "Me" and SelectID = 0 then response.write ("selected")%>>留言</option>
                                </select>
                                <input name="SelectID" type="hidden">
                            </div>
                        </div>
                        <div class="form-group" data-div="Out2" <%If OutFlag Then response.Write "style='display:none'"%>>
                            <label class="col-md-2 control-label">静态页链接网址：</label>
                            <div class="col-md-5">
                                <input name="HtmlNavUrl" type="text" id="HtmlNavUrl"  class="form-control" value="<%=HtmlNavUrl%>" readonly>
                                <span class="help-block"></span>
                            </div>
                            </label>
                        </div>
                        <div class="form-group" data-div="Out1" <%If not OutFlag Then response.Write "style='display:none'"%>>
                            <label class="col-md-2 control-label">外部链接网址：</label>
                                <div class="col-md-5">
                                    <input name="OutUrl" type="text" id="OutUrl"  class="form-control" value="<%=OutUrl%>">
                                    <span class="help-block">请以“http://”开头</span>
                                </div>
                            </label>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label">排序：</label>
                                <div class="col-md-5">
                                    <input name="Sequence" type="text" id="Sequence"  class="form-control" value="<%=Sequence%>" onKeyDown="OnKeyDown()" onChange="OnChange(this, <%=Sequence%>)">
                                    <span class="help-block"></span>
                                </div>
                            </label>
                        </div>
                        <div class="form-group">
                            <div class="col-md-2 control-label"></div>
                            <div class="col-md-10">
                                <input name="submitSaveEdit" type="submit" id="submitSaveEdit" class="btn btn-danger btn-lg" value="保存导航信息">
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
<script language="javascript">
$("[name=SelectClass]").change(function(){
	HtmlNavUrl = $(this).children("option:selected").attr("data")
	$("[name=HtmlNavUrl]").val(HtmlNavUrl)
})
</script>
</body>
</html>
<%
Function NavAbout()
	TT = "Ab"
	Title = Conn.execute("select * from keyicms_System where DataForm=1")("SystemName")
	If TT = SelectClass and SelectID = 0 Then 
		NavAbout = "<option value='"&TT&"_0' data='/"&AboutDir&"' selected>"&Title&"</option>"
	Else
		NavAbout = "<option value='"&TT&"_0' data='/"&AboutDir&"'>"&Title&"</option>"
	End If
	set rs = server.CreateObject("adodb.recordset")
	rs.open "select * from keyicms_About Order by ID desc",Conn,1,2
	do until rs.eof
		HTMLUrl = "/" & AboutDir & "/" & rs("ClassSeo") & Separated & rs("ID") & "." & HTMLName
		If SelectClass = TT and SelectID = rs("ID") then
			NavAbout = NavAbout & "<option value='"&TT&"_"&rs("ID")&"' data='"&HTMLUrl&"' selected>"
		Else
			NavAbout = NavAbout & "<option value='"&TT&"_"&rs("ID")&"' data='"&HTMLUrl&"'>"
		End If
		NavAbout = NavAbout & "│├ "&rs("AboutName")&"</option>"
		rs.movenext
	loop
	rs.close : set rs = nothing
End Function

Function NavOne(Nav, ParentID)
	If Nav = "News" Then
		DataForm = "keyicms_NewsSort"
		Dir = NewsDir
		Title = Conn.execute("select * from keyicms_System where DataForm=2")("SystemName")
		TT = "Ne"
	End If
	If Nav = "Product" Then
		DataForm = "keyicms_ProductSort"
		Dir = ProDir
		Title = Conn.execute("select * from keyicms_System where DataForm=3")("SystemName")
		TT = "Pr"
	End If
	If Nav = "Case" Then
		DataForm = "keyicms_CaseSort"
		Dir = CaseDir
		Title = Conn.execute("select * from keyicms_System where DataForm=4")("SystemName")
		TT = "Ca"
	End If
	If Nav = "Download" Then
		DataForm = "keyicms_DownSort"
		Dir = DownDir
		Title = Conn.execute("select * from keyicms_System where DataForm=5")("SystemName")
		TT = "Do"
	End If
	If TT = SelectClass and SelectID = 0 Then 
		NavOne = "<option value='"&TT&"_0' data='/"&Dir &"' selected>"&Title&"</option>"
	Else
		NavOne = "<option value='"&TT&"_0' data='/"&Dir &"'>"&Title&"</option>"
	End If
    dim ClassNum
	set rs = server.CreateObject("adodb.recordset")
	rs.open "select * from "&DataForm&" where ParentID = 0 Order by ID desc",Conn,1,2
	do until rs.eof
		HTMLUrl = "/" & Dir & "/" & rs("FolderName")
		If SelectClass = TT and SelectID = rs("ID") then
			NavOne = NavOne & "<option value='"&TT&"_"&rs("ID")&"' data='"&HTMLUrl&"' selected>"
		Else
			NavOne = NavOne & "<option value='"&TT&"_"&rs("ID")&"' data='"&HTMLUrl&"'>"
		End If
		NavOne = NavOne & "│├ "&rs("SortName")&"</option>"
		NavOne = NavOne & NavSub(DataForm, rs("ID"), TT, Dir, 1)
		rs.movenext
	loop
	rs.close : set rs = nothing
End Function

Function NavSub(DataForm, ParentID, TT, Dir, ClassNum)
	dim rs, sql
	set rs = server.CreateObject("adodb.recordset")
	sql = "select * from "&DataForm&" where ParentID = "&ParentID&" Order by ID desc"
	rs.open sql, Conn, 1, 2
	do until rs.eof
		HTMLUrl = "/" & Dir & "/" & rs("FolderName")	
		If SelectClass = TT and SelectID = rs("ID") then
			NavSub = NavSub & "<option value='"&TT&"_"&rs("ID")&"' data='"&HTMLUrl&"' selected>"
		Else
			NavSub = NavSub & "<option value='"&TT&"_"&rs("ID")&"' data='"&HTMLUrl&"'>"
		End If
		For ClassN = 1 to ClassNum
			NavSub = NavSub & "│"
		Next
		NavSub = NavSub & "│├ "&rs("SortName")&"</option>"

		ClassNum = ClassNum + 1
		NavSub = NavSub & NavSub(DataForm, rs("ID"), TT, Dir, ClassNum)
		ClassNum = ClassNum - 1

		rs.movenext
	loop
	rs.close : set rs = nothing
End Function

sub NavEdit()
dim Action, rs, sql
Action=request.QueryString("Action")
if Action="SaveEdit" then
	set rs = server.createobject("adodb.recordset")
	if len(trim(request.Form("NavName")))<2 then
		Call SweetAlert("warning", "友情提示", "请填写导航名称并保持至少在一个汉字以上！", "false", "history.back(-1)")
		response.end
	end If
	If Request.Form("OutFlag") = 1  Then
		if trim(Request.Form("OutUrl")) = "" then
			Call SweetAlert("warning", "友情提示", "请填写外链接网址！", "false", "history.back(-1)")
			response.end
		end if
	Else
		if Request.Form("SelectClass") = "" or trim(Request.Form("HtmlNavUrl")) = "" then
			Call SweetAlert("warning", "友情提示", "请选择导航并静态链接网址不能为空！", "false", "history.back(-1)")
			response.end
		end if
	End If

	if Request.Form("NavType") = 1 then
		ParentID = Request.Form("ParentID1")
	else
		ParentID = Request.Form("ParentID2")
	end if
	If ParentID = "0" Then
		NavPath = "0,"
	Else
		NavPath = Conn.execute("Select * from keyicms_Navigation where ID="&ParentID)("NavPath")
	End If

	if Result="Add" Then
		sql="select * from keyicms_Navigation"
		rs.open sql,Conn,1,3
		rs.addnew
		rs("NavName") = trim(Request.Form("NavName"))
		if Request.Form("ViewFlag") = 1 then
			rs("ViewFlag") = Request.Form("ViewFlag")
		else
			rs("ViewFlag") = 0
		end if
		if Request.Form("MobileFlag") = 1 then
			rs("MobileFlag") = Request.Form("MobileFlag")
		else
			rs("MobileFlag")= 0
		end if
		
		rs("NavType") = Request.Form("NavType")
		rs("ParentID") = ParentID
		rs("NavPath") = NavPath & rs("ID") & ","

		if Request.Form("OutFlag") = 1 then
			rs("OutFlag") = Request.Form("OutFlag")
			rs("OutUrl") = Request.Form("OutUrl")
		else
			rs("OutFlag") = 0
			rs("HtmlNavUrl") = trim(Request.Form("HtmlNavUrl"))
			
			SelectClass = split(trim(Request.Form("SelectClass")),"_")
			rs("SelectClass") = SelectClass(0)
			rs("SelectID") = SelectClass(1)
		end if
		
		If rs("Sequence") = "" then
			rs("Sequence") = 99
		else
			rs("Sequence") = Request.Form("Sequence")
		end if
		rs("AddTime") = now()
	end if
	if Result="Modify" Then
		sql = "select * from keyicms_Navigation where ID="&ID
		rs.open sql,Conn,1,3
		fromID = rs("ID")
		fromParentID = rs("ParentID")
		fromSortPath = rs("NavPath")
		rs.close

		If CInt(fromParentID) <> CInt(ParentID) Then
			toID = ParentID
			toParentID = Conn.execute("select * from keyicms_Navigation where ID="&ParentID)("ParentID")
			toSortPath= NavPath

			if toID = "" or toParentID = "" or toSortPath = "" then
				Call SweetAlert("success", "操作成功", "请选择移动的目标位置！", "false", "history.back(-1);")
			end if
			if fromParentID = 0 then
				Call SweetAlert("success", "操作成功", "一级分类无法被移动！", "false", "history.back(-1);")
			end if
			if fromSortPath = toSortPath then
				Call SweetAlert("success", "操作成功", "当前选择的移动类别和目标位置相同，操作无效！", "false", "history.back(-1);")
			end if
			if Instr(toSortPath, fromSortPath)>0 or fromParentID = toID then
				Call SweetAlert("success", "操作成功", "不能将类别移动到本类或下属类里，操作无效！", "false", "history.back(-1);")
			end if
			fromParentSortPath = Conn.execute("select * from keyicms_Navigation where ID=(select ParentID from keyicms_Navigation where ID="&ID&")")("NavPath")
			Conn.execute("update keyicms_Navigation set NavPath='"&toSortPath&"'+Mid(NavPath,Len('"&fromParentSortPath&"')+1) where Instr(NavPath,'"&fromSortPath&"')>0")
			Conn.execute("update keyicms_Navigation set ParentID='"&toID&"' where ID="&fromID)
			Conn.execute("update keyicms_Navigation set NavType="&Request.Form("NavType")&" where Instr(NavPath,'"&toSortPath&"')>0")
		End If

		sql = "select * from keyicms_Navigation where ID="&ID
		rs.open sql,Conn,1,3
		rs("NavName") = trim(Request.Form("NavName"))
		if Request.Form("ViewFlag") = 1 then
			rs("ViewFlag") = Request.Form("ViewFlag")
		else
			rs("ViewFlag") = 0
		end if
		if Request.Form("ViewFlag") = 1 then
			rs("ViewFlag") = Request.Form("ViewFlag")
		else
			rs("ViewFlag") = 0
		end if
		if Request.Form("MobileFlag") = 1 then
			rs("MobileFlag") = Request.Form("MobileFlag")
		else
			rs("MobileFlag")= 0
		end if

		if Request.Form("OutFlag") = 1 then
			rs("OutFlag") = Request.Form("OutFlag")
			rs("OutUrl") = Request.Form("OutUrl")
		else
			rs("OutFlag") = 0
			rs("HtmlNavUrl") = trim(Request.Form("HtmlNavUrl"))
			
			SelectClass = split(trim(Request.Form("SelectClass")),"_")
			rs("SelectClass") = SelectClass(0)
			rs("SelectID") = SelectClass(1)
		end if
		rs("Sequence") = Request.Form("Sequence")
	end if
	rs.update
	rs.close : set rs = nothing
	Call SweetAlert("success", "操作成功", "导航设置成功！", "false", "location.replace('Ky_Navigation.Asp')")
else
	if Result="Modify" then
		set rs = server.createobject("adodb.recordset")
		sql = "select * from keyicms_Navigation where ID="&ID
		rs.open sql,Conn,1,1
		NavName = rs("NavName")
		ViewFlag = rs("ViewFlag")
		MobileFlag = rs("MobileFlag")
		NavType = rs("NavType")
		NavParentID = rs("ParentID")
		OutFlag = rs("OutFlag")
		OutUrl = rs("OutUrl")
		HtmlNavUrl = rs("HtmlNavUrl")
		SelectClass = rs("SelectClass")
		SelectID = rs("SelectID")
		Sequence = rs("Sequence")
		rs.close : set rs = nothing
	else
		HtmlNavUrl = "/"
		Sequence = 99
		NavType = 1
		OutFlag = 0
	end if
end if
end sub

Function Navigation1(ParentID, ClassNum)
	dim rs, sql
	set rs = Server.CreateObject("adodb.recordset")
	sql = "select * from keyicms_Navigation where ParentID="&ParentID&" and NavType=1 Order by Sequence asc, ID Desc"
	rs.open sql, Conn, 1, 2
	Do Until rs.eof
		If rs("ID") <> ID Then
			response.Write "<option value='"&rs("ID")&"' "
			If NavParentID = rs("ID") Then response.Write "selected"
			response.Write ">"

			For ClassN = 1 to ClassNum
				response.Write "│"
			Next
			response.Write "│├ "&rs("NavName")
			response.Write "</option>"
			Call Navigation1(rs("ID"), ClassNum + 1)
		End If
		rs.movenext
	Loop
	rs.close : set rs = nothing
End Function

Function Navigation2(ParentID, ClassNum)
	dim rs, sql
	set rs = Server.CreateObject("adodb.recordset")
	sql = "select * from keyicms_Navigation where ParentID="&ParentID&" and NavType=2 Order by Sequence asc, ID Desc"
	rs.open sql, Conn, 1, 2
	Do Until rs.eof
		If rs("ID") <> ID Then
			response.Write "<option value='"&rs("ID")&"' "
			If NavParentID = rs("ID") Then response.Write "selected"
			response.Write ">"

			For ClassN = 1 to ClassNum
				response.Write "│"
			Next
			response.Write "│├ "&rs("NavName")
			response.Write "</option>"
			Call Navigation2(rs("ID"), ClassNum + 1)
		End If
		rs.movenext
	Loop
	rs.close : set rs = nothing
End Function
%>