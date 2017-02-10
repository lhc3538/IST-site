<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="../Include/system_language.asp" -->
<!--#include file="CheckAdmin.Asp"-->
<!--#include file="../Html/html_Config.Asp"-->
<!--#include file="../Html/html_Function.Asp"-->
<!--#include file="../Html/html_To_Product.Asp"-->
<%

Dim Result
Result = request.QueryString("Result")
Dim ID, ProductName, ViewFlag, ClassSeo, SortName, SortID, SortPath
Dim ProductNo, Maker, CommendFlag, NewFlag, MobileFlag, SeoTitle, SeoKeywords, SeoDescription
Dim SmallPic, BigPic, OtherPic, Content, ClickNumber, Sequence
ID = request.QueryString("ID")
Call ProductEdit()
Call SiteInfo()
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
	K('#SmallImage').click(function() {
		editor.loadPlugin('image', function() {
			editor.plugin.imageDialog({
				imageUrl : K('#SmallPic').val(),
				clickFn : function(url, title, width, height, border, align) {
					K('#SmallPic').val(url);
					editor.hideDialog();
				}
			});
		});
	});
	K('#BigImage').click(function() {
		editor.loadPlugin('image', function() {
			editor.plugin.imageDialog({
				imageUrl : K('#BigPic').val(),
				clickFn : function(url, title, width, height, border, align) {
					K('#BigPic').val(url);
					editor.hideDialog();
				}
			});
		});
	});
	K('#OtherImage').click(function() {
		editor.loadPlugin('multiimage', function() {
			var aVal=""
			editor.plugin.multiImageDialog({
				
				clickFn : function(urlList) {
					var div = K('#OtherPic');
					aVal=div.val()
					K.each(urlList, function(i, data) {
						if (div.val() == "")
							aVal =aVal + data.url
						else
							aVal =aVal + "|" +  data.url

					div.val(aVal);
					});
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
	If InStr(CompanyAdminPurview, "|3,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style=""color:red; font-size:14px;"")>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If
	%>
    <div class="main">
        <div class="row">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i> <%=Keyicms_Lang_Copyright(14)%></li>
                <li class="active"><a href="">功能模块</a></li>
                <li class="active">图片管理</li>
            </ol>	
        </div>
        <div class="row ">
            <div class="col-lg-12 hideInIE8">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-edit"></i><span class="break"></span><strong>图片管理</strong></h2> 
                    </div>
                    <div class="panel-body">
                        <a class="ShortCut1" href="Ky_ProductEdit.Asp?Result=Add">  
                            <i class="fa fa-sign-in"></i>
                            <div class="title">添加图片</div>
                        </a>					
                        <a class="ShortCut2" href="Ky_Product.Asp">  
                            <i class="fa fa-th"></i>
                            <div class="title">图片列表</div>
                        </a>	
                        <a class="ShortCut3" href="Ky_ProductSort.Asp">  
                            <i class="fa fa-sitemap"></i>
                            <div class="title">图片分类</div>
                        </a>													
                    </div>
                </div>
            </div>
        </div>
        <div class="row">
			<div class="col-md-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-gears"></i><span class="break"></span><strong><%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改<%End If%>图片</strong></h2>
                    </div>
                    <div class="panel-body">
                    <form name="editForm" method="post" action="?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>" class="form-horizontal">
                        <div class="form-group">
                            <label class="col-md-2 control-label">图片名称：</label>
                            <div class="col-md-4">
                                <input name="ProductName" type="text" id="ProductName" class="form-control" value="<%=ProductName%>" maxlength="100">
                            </div>
                        </div>  
                        <%If EnableSeo = True Then%>
                        <div class="form-group">
                            <label class="col-md-2 control-label">关键词：</label>
                            <div class="col-md-6">
                                <input name="SeoKeywords" type="text" id="SeoKeywords" class="form-control" value="<%=SeoKeywords%>" maxlength="250">
                                <span class="help-block">SEO：多个关键词请用“,”半角逗号分开</span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label">描述：</label>
                            <div class="col-md-6">
                                <input name="SeoDescription" type="text" id="SeoDescription" class="form-control" value="<%=SeoDescription%>" maxlength="250">
                                <span class="help-block">SEO：请填写描述</span>
                            </div>
                        </div>
                        <%End If%>
                        <div class="form-group">
                            <label class="col-md-2 control-label">图片类别：</label>
                            <div class="col-md-5">
                                <select size="1" class="form-control" name="SortSelect" id="SortSelect" onChange="AddSort()">
                                <option value="" valuepath="">请选择...</option>
                                <%Call ProductSort(0, 0)%>
                                </select>
                                <span class="help-block">
                                    <input name="SortID" type="hidden" id="SortID" value="<%=SortID%>" readonly class="form-control">
                                    <input name="SortPath" type="hidden" id="SortPath" value="<%=SortPath%>" readonly class="form-control">
                                </span>
                            </div>
                        </div>

                        <%
						If Result = "Modify" Then
							Set rs = server.CreateObject("adodb.recordset")
							sql = "select * from keyicms_Product where ID="& ID
							rs.Open sql, conn, 1, 1
							If rs.bof And rs.EOF Then
								response.Write ("<center>数据库记录读取错误！</center>")
								response.End
							End If
							If rs("Attribute")<>"" And rs("Attribute_value")<>"" Then
								Attribute_1 = Split(rs("Attribute"), "§§§")
								Attribute_value_1 = Split(rs("Attribute_value"), "§§§")
								Num_1 = UBound(Attribute_1) + 1
							Else
								Num_1 = 0
							End If
							rs.Close
							Set rs = Nothing
						Else
							Num_1 = 0
						End If
						%>
                        
                        <div class="form-group">
                            <label class="col-md-2 control-label">大图：</label>
                            <div class="col-md-5">
                            	<div class="input-group">
                                    <input type="text" id="BigPic" name="BigPic" class="form-control" value="<%=BigPic%>" readonly>
                                    <span class="input-group-btn">
                                    <button type="button" id="BigImage" class="btn btn-default">上传图片</button>
                                    </span>
                                </div>
                                <span class="help-block">图片大图主要是做内页显示使用</span>
                            </div>
                        </div>                     
                        <div class="form-group">
                            <label class="col-md-2 control-label">小图：</label>
                            <div class="col-md-5"> 
                            	<div class="input-group">
                                    <input type="text" id="SmallPic" name="SmallPic" class="form-control" value="<%=SmallPic%>" readonly>
                                    <span class="input-group-btn">
                                    <button type="button" id="SmallImage" class="btn btn-default">上传图片</button>
                                    </span>
                                </div>
                                <span class="help-block">小图可以直接使用大图地址：<a href="javascript:void(0)" onClick="document.editForm.SmallPic.value=document.editForm.BigPic.value">点击复制上方大图地址</a></span>
                            </div>
                        </div>  
                        <div class="form-group">
                            <label class="col-md-2 control-label">更多图片：</label>
                            <div class="col-md-5">
                                <textarea rows="5" cols="80" name="OtherPic" id="OtherPic" class="form-control"><%=OtherPic%></textarea>                                
                                <span class="input-group-btn">
                                <button type="button" id="OtherImage" class="btn btn-default">批量上传</button>
                                </span>
                                <span class="help-block">上传多张图片</span>
                            </div>
                        </div>  
                        <div class="form-group">
                            <label class="col-md-2 control-label">图片详细介绍：</label>
                                <div class="col-md-10">
                                <textarea name="Content" style="width:95%;height:400px;visibility:hidden;" class="form-control"><%=Server.HTMLEncode(Content)%></textarea>
                                <span class="help-block"></span>
                            </div>
                        </div>
                        <%If EnableViews = True Then%>  
                        <div class="form-group">
                            <label class="col-md-2 control-label">浏览数：</label>
                            <div class="col-md-4">
                                <input type="text" name="ClickNumber" id="ClickNumber" class="form-control" value="<%=ClickNumber%>" onKeyDown="OnKeyDown()" onChange="OnChange(this, <%=ClickNumber%>)" />
                                <span class="help-block">可以手动输入浏览数量。</span>
                            </div>
                        </div> 
                        <%End If%>
                        <div class="form-group">
                            <label class="col-md-2 control-label">排序：</label>
                            <div class="col-md-4">
                                <input name="Sequence" type="text" id="Sequence" class="form-control" value="<%=Sequence%>" onKeyDown="OnKeyDown()" onChange="OnChange(this, <%=Sequence%>)">
                                <span class="help-block">数字越小排名越靠前。</span>
                            </div>
                        </div>
                        <div class="form-group">
                            <div class="col-md-2 control-label"></div>
                            <div class="col-md-10">
                                <input name="submitSaveEdit" type="submit" id="submitSaveEdit" class="btn btn-danger btn-lg" value="保存图片信息">
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
Sub ProductEdit()
Dim Action, rsRepeat, rs, sql
Action = request.QueryString("Action")
If Action = "SaveEdit" Then
	Set rs = server.CreateObject("adodb.recordset")
	If Len(Trim(request.Form("ProductName")))<1 Then
		Call SweetAlert("warning", "友情提示", "请填写图片名称！", "false", "history.back(-1)")
		response.End
	End If
	If Request.Form("SortID") = "" And Request.Form("SortPath") = "" Then
		Call SweetAlert("warning", "友情提示", "请选择所属分类！", "false", "history.back(-1)")
		response.End
	End If
	If LTrim(request.Form("SmallPic")) = "" Then
		Call SweetAlert("warning", "友情提示", "请上传图片小图！", "false", "history.back(-1)")
		response.End
	End If
	If LTrim(request.Form("BigPic")) = "" Then
		Call SweetAlert("warning", "友情提示", "请上传图片大图！", "false", "history.back(-1)")
		response.End
	End If
	If LTrim(request.Form("Content")) = "" Then
		Call SweetAlert("warning", "友情提示", "请填写图片详细介绍！", "false", "history.back(-1)")
		response.End
	End If

	If Result = "Add" Then
		Set rsRepeat = conn.Execute("select ProductNo from keyicms_Product where ProductNo='" & Trim(Request.Form("ProductNo")) & "'")
		rsRepeat.Close : Set rsRepeat = Nothing

		sql = "select * from keyicms_Product"
		rs.Open sql, conn, 1, 3
		rs.addnew
		rs("ProductName") = Trim(Request.Form("ProductName"))
		If Request.Form("ClassSeo") = "" then
			If Request.Form("oAutopinyin") = "1" then rs("ClassSeo") = TimeToName()
			If Request.Form("oAutopinyin") = "2" then rs("ClassSeo") = Left(Pinyin(trim(request.form("ProductName"))),200)
		Else
			rs("ClassSeo") = trim(Request.form("ClassSeo"))
		End If
		
		rs("ViewFlag") = 1

		If Request.Form("CommendFlag") = 1 Then
			rs("CommendFlag") = Request.Form("CommendFlag")
		Else
			rs("CommendFlag") = 0
		End If
		If Request.Form("NewFlag") = 1 Then
			rs("NewFlag") = Request.Form("NewFlag")
		Else
			rs("NewFlag") = 0
		End If
		If Request.Form("MobileFlag") = 1 Then
			rs("MobileFlag") = Request.Form("MobileFlag")
		Else
			rs("MobileFlag") = 0
		End If
		
		rs("SortID") = Request.Form("SortID")
		rs("SortPath") = Request.Form("SortPath")
		rs("ProductNo") = Trim(Request.Form("ProductNo"))
		rs("Maker") = Trim(Request.Form("Maker"))
		rs("SmallPic") = Trim(Request.Form("SmallPic"))
		rs("BigPic") = Trim(Request.Form("BigPic"))
		rs("OtherPic") = Trim(Request.Form("OtherPic"))
		rs("Content") = RTrim(Request.Form("Content"))
		rs("AddTime") = Now()
		rs("UpdateTime") = Now()
		Num_1 = CheckStr(Request.Form("Num_1"), 1)
		If Num_1 = "" Then Num_1 = 0
		If Num_1>0 Then
			For i = 1 To Num_1
				If CheckStr(Request.Form("Attribute"&i), 0)<>"" And CheckStr(Request.Form("Attribute"&i&"_value"), 0)<>"" Then
					If Attribute = "" Then
						Attribute = CheckStr(Request.Form("Attribute"&i), 0)
						Attribute_value = CheckStr(Request.Form("Attribute"&i&"_value"), 0)
					Else
						Attribute = Attribute&"§§§"&CheckStr(Request.Form("Attribute"&i), 0)
						Attribute_value = Attribute_value&"§§§"&CheckStr(Request.Form("Attribute"&i&"_value"), 0)
					End If
				End If
			Next
		End If
		rs("Attribute") = Attribute
		rs("Attribute_value") = Attribute_value
		rs("SeoTitle") = Trim(Request.Form("SeoTitle"))
		rs("SeoKeywords") = Trim(Request.Form("SeoKeywords"))
		rs("SeoDescription") = Trim(Request.Form("SeoDescription"))
		
		If EnableViews Then rs("ClickNumber") = trim(Request.Form("ClickNumber"))
		rs("Sequence") = trim(Request.Form("Sequence"))
		rs.update
		rs.close : set rs = nothing
	  
		set rs = server.createobject("adodb.recordset")
		sql = "select top 1 ID,ClassSeo from keyicms_Product order by ID desc"
		rs.Open sql, conn, 1, 1
		ID = rs("ID")
		rs.close : set rs = nothing

		Call html_To_Product(ID, "PC")
		If mStatus Then Call html_To_Product(ID, "Mobile")
	End If
	If Result = "Modify" Then
		sql = "select * from keyicms_Product where ID="&ID
		rs.Open sql, conn, 1, 3
		rs("ProductName") = Trim(Request.Form("ProductName"))

		rs("ViewFlag") = 1

		If Request.Form("CommendFlag") = 1 Then
			rs("CommendFlag") = Request.Form("CommendFlag")
		Else
			rs("CommendFlag") = 0
		End If
		If Request.Form("NewFlag") = 1 Then
			rs("NewFlag") = Request.Form("NewFlag")
		Else
			rs("NewFlag") = 0
		End If
		If Request.Form("MobileFlag") = 1 Then
			rs("MobileFlag") = Request.Form("MobileFlag")
		Else
			rs("MobileFlag") = 0
		End If

		rs("SortID") = Request.Form("SortID")
		rs("SortPath") = Request.Form("SortPath")
		rs("ProductNo") = Trim(Request.Form("ProductNo"))
		rs("Maker") = Trim(Request.Form("Maker"))
		rs("SmallPic") = Trim(Request.Form("SmallPic"))
		rs("BigPic") = Trim(Request.Form("BigPic"))
		rs("OtherPic") = Trim(Request.Form("OtherPic"))
		rs("Content") = RTrim(Request.Form("Content"))
		rs("UpdateTime") = Now()
		Num_1 = CheckStr(Request.Form("Num_1"), 1)
		If Num_1 = "" Then Num_1 = 0
		If Num_1 > 0 Then
			For i = 0 To Num_1 - 1
				If CheckStr(Request.Form("Attribute"&i), 0)<>"" And CheckStr(Request.Form("Attribute"&i&"_value"), 0)<>"" Then
					If Attribute = "" Then
						Attribute = CheckStr(Request.Form("Attribute"&i), 0)
						Attribute_value = CheckStr(Request.Form("Attribute"&i&"_value"), 0)
					Else
						Attribute = Attribute&"§§§"&CheckStr(Request.Form("Attribute"&i), 0)
						Attribute_value = Attribute_value&"§§§"&CheckStr(Request.Form("Attribute"&i&"_value"), 0)
					End If
				End If
			Next
		End If
		rs("Attribute") = Attribute
		rs("Attribute_value") = Attribute_value
		rs("SeoTitle") = Trim(Request.Form("SeoTitle"))
		rs("SeoKeywords") = Trim(Request.Form("SeoKeywords"))
		rs("SeoDescription") = Trim(Request.Form("SeoDescription"))
		
		If EnableViews Then rs("ClickNumber") = trim(Request.Form("ClickNumber"))
		rs("Sequence") = trim(Request.Form("Sequence"))

		ID = rs("ID")
		rs.update
		rs.close : set rs = nothing
		
		'Call html_To_Product(ID, "PC")
		If mStatus Then Call html_To_Product(ID, "Mobile")
	End If
	Call SweetAlert("success", "操作成功", "设置成功，相关页面已更新！", "false", "location.replace('Ky_Product.Asp')")
Else
	If Result = "Modify" Then
		Set rs = server.CreateObject("adodb.recordset")
		sql = "select * from keyicms_Product where ID="& ID
		rs.Open sql, conn, 1, 1
		If rs.bof And rs.EOF Then
			response.Write ("<center>数据库记录读取错误！</center>")
			response.End
		End If
		ProductName = rs("ProductName")
		ViewFlag = rs("ViewFlag")
		ClassSeo = rs("ClassSeo")&Separated&rs("ID")
		SortName = SortText(rs("SortID"))
		SortID = rs("SortID")
		SortPath = rs("SortPath")
		ProductNo = rs("ProductNo")
		Maker = rs("Maker")
		CommendFlag = rs("CommendFlag")
		NewFlag = rs("NewFlag")
		MobileFlag = rs("MobileFlag")
		SmallPic = rs("SmallPic")
		BigPic = rs("BigPic")
		OtherPic = rs("OtherPic")
		Content = rs("Content")
		SeoTitle = rs("SeoTitle")
		SeoKeywords = rs("SeoKeywords")
		SeoDescription = rs("SeoDescription")
		ClickNumber = rs("ClickNumber")
		Sequence = rs("Sequence")
		rs.close : set rs = nothing
	Else
		ClickNumber = 10
		Sequence = 10
		randomize timer
		ProductNo = Hour(now) & Minute(now) & Second(now) & "-" & int(900*rnd)+100
	End If
End If
End Sub

Function SortText(ID)
    Dim rs, sql
    Set rs = server.CreateObject("adodb.recordset")
    sql = "Select * From keyicms_ProductSort where ID="&ID
    rs.Open sql, conn, 1, 1
    SortText = rs("SortName")
	rs.close : set rs = nothing
End Function

Function ProductSort(ID, ClassNum)
	Dim rs
	set rs = Server.CreateObject("adodb.recordset")
	rs.open "select * from Keyicms_ProductSort where ParentID="&ID&" Order by ID asc", conn, 1, 2
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
		Call ProductSort(rs("ID"), ClassNum)
		ClassNum = ClassNum - 1
		rs.movenext
	Loop
	rs.close : set rs = nothing
End Function
%>