<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="../Include/System_Language.asp" -->
<!--#include file="../Include/Version.Asp" -->
<!--#include file="CheckAdmin.Asp"-->
<!--#include file="../Html/html_Config.Asp"-->
<!--#include file="../Html/html_Function.Asp"-->
<!--#include file="../Html/html_To_About.Asp"-->
<%
Dim Result
Result = request.QueryString("Result")
Dim ID, AboutName, ClassSeo, ViewFlag, MobileFlag, Content, SeoKeywords, SeoDescription, Sequence
ID = request.QueryString("ID")
Call AboutEdit()
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
	var editor = K.create('textarea[name="Content"]', {
		cssPath : 'KEditor_keyicms/plugins/code/prettify.css',
		uploadJson : 'KEditor_keyicms/asp/upload_json.asp',
		fileManagerJson : 'KEditor_keyicms/asp/file_manager_json.asp',
		allowFileManager : true,
		afterBlur : function() {
			this.sync();
			K.ctrl(document, 13, function() {
				K('form[name=example]')[0].submit();
			});
			K.ctrl(this.edit.doc, 13, function() {
				K('form[name=example]')[0].submit();
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
	If InStr(CompanyAdminPurview, "|1,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style=""color:red; font-size:14px;"")>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If
	%>
    <div class="main">
        <div class="row">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i> <%=Keyicms_Lang_Copyright(14)%></li>
                <li><%=Keyicms_Lang_Copyright(2)%></li>
                <li class="active">学院信息管理</li>
            </ol>	
        </div>	
        <div class="row">
            <div class="col-lg-12 hideInIE8">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-edit"></i><span class="break"></span><strong>学院信息管理</strong></h2> 
                    </div>
                    <div class="panel-body">
                        <a class="ShortCut1" href="Ky_AboutEdit.Asp?Result=Add">  
                            <i class="fa fa-sign-in"></i>
                            <div class="title">添加学院信息</div>
                        </a>					
                        <a class="ShortCut2" href="Ky_About.Asp">  
                            <i class="fa fa-list"></i>
                            <div class="title">学院信息列表</div>
                        </a>				
                    </div>
                </div>
            </div>
        </div>
        <div class="row">
            <div class="col-md-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-gears"></i><span class="break"></span><strong><%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改<%End If%>学院信息</strong></h2>
                    </div>
                    <div class="panel-body">
                    <form name="editForm" method="post" action="?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>" class="form-horizontal ">
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">学院信息名称：</label>
                            <div class="col-md-4">
                                <input name="AboutName" type="text" id="AboutName" class="form-control" value="<%=AboutName%>">
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
                            <label class="col-md-2 control-label" for="text-input">信息内容：</label>
                            <div class="col-md-10">
                                <textarea name="Content" class="form-control" style="width:95%;height:400px;visibility:hidden;"><%=Server.HTMLEncode(Content)%></textarea>
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
                                <input name="submitSaveEdit" type="submit" id="submitSaveEdit" class="btn btn-danger btn-lg" value="保存学院信息">
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
Sub AboutEdit()
    Dim Action, rsCheckAdd, rs, sql
    Action = request.QueryString("Action")
    If Action = "SaveEdit" Then
        Set rs = server.CreateObject("adodb.recordset")
        If Len(Trim(request.Form("AboutName")))<1 Then
			Call SweetAlert("warning", "友情提示", "请填写信息标题！", "false", "history.back(-1)")
            response.End
        End If
        If Trim(request.Form("Content")) = "" Then
			Call SweetAlert("warning", "友情提示", "请填写信息内容！", "false", "history.back(-1)")
            response.End
        End If

        If Result = "Add" Then
            sql = "select * from keyicms_About"
            rs.Open sql, conn, 1, 3
            rs.addnew
			rs("AboutName")=trim(Request.Form("AboutName"))
			If Request.Form("ClassSeo") = "" then
				If Request.Form("oAutopinyin") = "1" then rs("ClassSeo") = TimeToName()
				If Request.Form("oAutopinyin") = "2" then rs("ClassSeo") = Left(Pinyin(trim(request.form("AboutName"))),200)
			Else
				rs("ClassSeo") = trim(Request.form("ClassSeo"))
			End If
            rs("AboutName") = Trim(Request.Form("AboutName"))

            rs("ViewFlag") = 1	'默认显示

            If Request.Form("MobileFlag") = 1 Then
                rs("MobileFlag") = Request.Form("ViewFlag")
            Else
                rs("MobileFlag") = 0
            End If
			rs("SeoKeywords") = trim(Request.Form("SeoKeywords"))
			rs("SeoDescription") = trim(Request.Form("SeoDescription"))
            rs("Content") = RTrim(Request.Form("Content"))
			rs("Sequence") = trim(Request.Form("Sequence"))
            rs("AddTime") = Now()
            rs("UpdateTime") = Now()
            rs.update
			rs.close : set rs = nothing
		  
			set rs = server.createobject("adodb.recordset")
            sql = "select top 1 ID,ClassSeo from keyicms_About order by ID desc"
            rs.Open sql, conn, 1, 1
            ID = rs("ID")
			rs.close : set rs = nothing
            
			'Call Html_To_About(ID, "PC")
			If mStatus Then Call Html_To_About(ID, "Mobile")
        End If
        If Result = "Modify" Then
            sql = "select * from keyicms_About where ID="&ID
            rs.Open sql, conn, 1, 3
            rs("AboutName") = Trim(Request.Form("AboutName"))

            rs("ViewFlag") = 1

            If Request.Form("MobileFlag") = 1 Then
                rs("MobileFlag") = Request.Form("ViewFlag")
            Else
                rs("MobileFlag") = 0
            End If
			rs("SeoKeywords") = trim(Request.Form("SeoKeywords"))
			rs("SeoDescription") = trim(Request.Form("SeoDescription"))
            rs("Content") = Request.Form("Content")
			rs("Sequence") = trim(Request.Form("Sequence"))
            rs("UpdateTime") = Now()
	  		ID = rs("ID")
            rs.update
			rs.close : set rs = nothing
					
			'Call Html_To_About(ID, "PC")
			If mStatus Then Call Html_To_About(ID, "Mobile")
        End If    
		Call SweetAlert("success", "操作成功", "设置成功，相关内容已更新！", "false", "location.replace('Ky_About.Asp')")
    Else
        If Result = "Modify" Then
            Set rs = server.CreateObject("adodb.recordset")
            sql = "select * from keyicms_About where ID="& ID
            rs.Open sql, conn, 1, 1
            AboutName = rs("AboutName")
            ViewFlag = rs("ViewFlag")
            MobileFlag = rs("MobileFlag")
	        ClassSeo=rs("ClassSeo")&Separated&rs("ID")
            Content = rs("Content")
			SeoKeywords = rs("SeoKeywords")
			SeoDescription = rs("SeoDescription")
			Sequence = rs("Sequence")
			rs.close : set rs = nothing
		Else
			Sequence = 99
        End If
    End If
End Sub
%>