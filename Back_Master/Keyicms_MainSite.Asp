<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="../Include/System_Language.asp" -->
<!--#include file="../Include/Version.Asp" -->
<!--#include file="CheckAdmin.Asp"-->
<!--#include file="../Html/html_Config.Asp"-->
<%
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
	K('#image1').click(function() {
		editor.loadPlugin('image', function() {
			editor.plugin.imageDialog({
				showRemote : true,
				imageUrl : K('#Ico').val(),
				clickFn : function(url, title, width, height, border, align) {
					K('#Ico').val(url);
					editor.hideDialog();
				}
			});
		});
	});
	K('#image2').click(function() {
		editor.loadPlugin('image', function() {
			editor.plugin.imageDialog({
				showRemote : true,
				imageUrl : K('#SiteLogo').val(),
				clickFn : function(url, title, width, height, border, align) {
					K('#SiteLogo').val(url);
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
    <!-- start: Content -->
	<%
	If InStr(CompanyAdminPurview, "|15,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style='color:red; font-size:14px;')>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If
	%>
    <div class="main">
        <div class="row ">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i> <%=Keyicms_Lang_Copyright(14)%></li>
				<li><%=Keyicms_Lang_Copyright(2)%></li>
                <li class="active"><%=Keyicms_Lang_System(10)%></li>
            </ol>	
        </div>	 
        <div class="row">
            <div class="col-md-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-gears"></i><span class="break"></span><strong><%=Keyicms_Lang_System(10)%></strong></h2>
                    </div>
                    <div class="panel-body">
                    <form action="?Action=Save" method="post" class="form-horizontal ">
                        <div class="panel-heading">
                            <h2><i class="fa fa-adjust"></i><span class="break"></span><strong><%=Keyicms_Lang_MainSite(22)%></strong></h2>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label"><%=Keyicms_Lang_MainSite(15)%>：</label>
                            <div class="col-md-6">
                                <label class="switch switchtop switch-danger">
                                    <input type="checkbox" class="switch-input" name="SiteStatus" value="-1" id="SiteStatus" <%if SiteStatus then response.write ("checked")%> />
                                    <span class="switch-label" data-on="开启" data-off="关闭"></span>
                                    <span class="switch-handle"></span>
                                </label>
                                 <span class="help-block"></span>
                            </div>
                        </div> 
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input"><%=Keyicms_Lang_MainSite(18)%>：</label>
                            <div class="col-md-6">
                                <textarea name="Explain" id="Explain" class="form-control" rows="2" cols="80"><%=Explain%></textarea>
                                <span class="help-block"></span>
                            </div>
                        </div>

                        <div class="panel-heading">
                            <h2><i class="fa fa-cog"></i><span class="break"></span><strong><%=Keyicms_Lang_MainSite(20)%></strong></h2>
                        </div>    
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input"><%=Keyicms_Lang_MainSite(0)%>：</label>
                            <div class="col-md-6">
                                <input name="SiteTitle" type="text" id="SiteTitle"  class="form-control" value="<%=SiteTitle%>">
                                <span class="help-block">请填写网站标题，这里的信息显示在浏览器顶部以及搜索引擎搜索出的网站主页标题</span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input"><%=Keyicms_Lang_MainSite(1)%>：</label>
                            <div class="col-md-5">
                                <input type="text" id="SiteUrl" name="SiteUrl"  class="form-control" value="<%=SiteUrl%>"> 
                                <span class="help-block">请填写您的网站网址如：www.keyicms.com 请不要带[http://]</span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input"><%=Keyicms_Lang_MainSite(2)%>：</label>
                            <div class="col-md-5">
                                <input type="text" id="IcpNumber" name="IcpNumber"  class="form-control" value="<%=IcpNumber%>"> 
                                <span class="help-block">已备案网站请填写您在工信部申请的的备案号码, 格式：蜀ICP备11017223号</span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input"><%=Keyicms_Lang_MainSite(3)%>：</label>
                            <div class="col-md-6"> 
                                <div class="input-group">
                                    <input name="SiteLogo" type="text" id="SiteLogo" value="<%=SiteLogo%>"  class="form-control" />
                                    <span class="input-group-btn">
                                    	<button type="button" id="image2" class="btn btn-default"><%=Keyicms_Lang_System(27)%></button>
                                    </span>
                                </div>
                                <span class="help-block">请上传网站LOGO图标，建议PNG格式，大小根据实际模板样式确定</span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input"><%=Keyicms_Lang_MainSite(4)%>：</label>
                            <div class="col-md-6"> 
                                <div class="input-group">
                                    <input name="Ico" type="text" id="Ico" value="<%=Ico%>"  class="form-control" size="60" />
                                    <span class="input-group-btn">
                                    	<button type="button" id="image1" class="btn btn-default"><%=Keyicms_Lang_System(27)%></button>
                                    </span>
                                </div>
                                <span class="help-block">//注 32*32像素 ico格式</span>
                            </div>
                        </div>
                        <div class="form-group">
                        	<label class="col-md-2 control-label" for="text-input"><%=Keyicms_Lang_MainSite(5)%>：</label>
                            <div class="col-md-6">
                                <textarea name="Stats" id="Stats"  style="padding:5px;" rows="3" cols="80" class="form-control" ><%=Stats%></textarea>
                                <span class="help-block">可填写统计代码等第三方接口JS代码，显示在网站底部</span>
                        	</div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input"><%=Keyicms_Lang_MainSite(6)%>：</label>
                            <div class="col-md-6">
								<textarea name="Keywords"  id="Keywords" class="form-control"  style="padding:5px;" rows="3" cols="80"><%=Keywords%></textarea>
                                <span class="help-block">填写keywords（网站关键词），多个关键词请用“,”分割。</span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input"><%=Keyicms_Lang_MainSite(7)%>：</label>
                            <div class="col-md-6">
                                <textarea name="Descriptions" id="Descriptions" class="form-control" style="padding:5px;" rows="3" cols="80"><%=Descriptions%></textarea>
                                <span class="help-block">填写Descriptions（网站描述），多个关键词请用“,”分割。</span>
                            </div>
                        </div>

                        <div class="panel-heading">
                            <h2><i class="fa  fa-phone"></i><span class="break"></span><strong><%=Keyicms_Lang_MainSite(21)%></strong></h2>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input"><%=Keyicms_Lang_MainSite(8)%>：</label>
                            <div class="col-md-5">
                                <input name="ComName" type="text" id="ComName"   class="form-control" value="<%=ComName%>">
                                 <span class="help-block"></span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input"><%=Keyicms_Lang_MainSite(9)%>：</label>
                            <div class="col-md-5">
                                <input name="Address" type="text" id="Address"  class="form-control" value="<%=Address%>">
                                <span class="help-block"></span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input"><%=Keyicms_Lang_MainSite(10)%>：</label>
                            <div class="col-md-5">
                                <input name="Contacts" type="text" id="Contacts" class="form-control" value="<%=Contacts%>">
                                <span class="help-block"></span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input"><%=Keyicms_Lang_MainSite(11)%>：</label>
                            <div class="col-md-5">
                                <input name="Telephone" type="text" id="Telephone" class="form-control" value="<%=Telephone%>">
                                <span class="help-block"></span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input"><%=Keyicms_Lang_MainSite(12)%>：</label>
                            <div class="col-md-5">
                                <input name="Mobile" type="text" id="Mobile" class="form-control" value="<%=Mobile%>">
                                <span class="help-block"></span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input"><%=Keyicms_Lang_MainSite(13)%>：</label>
                            <div class="col-md-5">
                                <input name="Fax" type="text" id="Fax" class="form-control" value="<%=Fax%>">
                                <span class="help-block"></span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input"><%=Keyicms_Lang_MainSite(14)%>：</label>
                            <div class="col-md-5">
                                <input name="Email" type="text" id="Email" class="form-control" value="<%=Email%>">
                                <span class="help-block"></span>
                            </div>
                        </div>
                        
                        <div class="panel-heading">
                            <h2><i class="fa fa-file-text-o"></i><span class="break"></span><strong><%=Keyicms_Lang_MainSite(22)%></strong></h2>
                        </div> 
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input"><%=Keyicms_Lang_MainSite(19)%>：</label>
                            <div class="col-md-6">
                                <textarea name="SiteDetail" id="SiteDetail" class="form-control" rows="6" cols="100"><%=SiteDetail%></textarea>
                                <span class="help-block">请填写学院简介，300字以内</span>
                            </div>
                        </div>
                        
                        <div class="form-group">
                            <div class="col-md-2 control-label"></div>
                            <div class="col-md-10">
                                <input name="submitSaveEdit" type="submit" id="submitSaveEdit" class="btn btn-danger btn-lg" value="<%=Keyicms_Lang_MainSite(24)%>">
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
	
    If Len(Trim(request.Form("SiteTitle")))<4 Then
		Call SweetAlert("warning", "操作失败！", "请详细填写您的网站标题并保持至少在两个汉字以上！", "false", "history.back()")
        response.End
    End If
    If Len(Trim(request.Form("SiteUrl")))<9 Then
		Call SweetAlert("warning", "操作失败！", "请详细填写您的学院网址！", "false", "history.back()")
        response.End
    End If
    If Len(Trim(request.Form("ComName")))<4 Then
 		Call SweetAlert("warning", "操作失败！", "请详细填写您的学院名称并保持至少在两个汉字以上！", "false", "history.back()")
       	response.End
    End If
    If Len(Trim(request.Form("Address")))<4 Then
		Call SweetAlert("warning", "操作失败！", "请详细填写您的学院地址并保持至少在两个汉字以上！", "false", "history.back()")
        response.End
    End If
    If Len(Trim(request.Form("Telephone")))<8 Then
		Call SweetAlert("warning", "操作失败！", "请详细填写正确的联系电话！", "false", "history.back()")
        response.End
    End If
	If Len(Trim(request.Form("Mobile")))<11 Then
		Call SweetAlert("warning", "操作失败！", "请详细填写正确的手机号码！", "false", "history.back()")
        response.End
    End If
	If Len(Trim(request.Form("Contacts")))<2 Then
		Call SweetAlert("warning", "操作失败！", "请详细填联系人并保持至少在两个汉字以上！", "false", "history.back()")
        response.End
    End If
    If Len(Trim(request.Form("Fax")))<10 Then
		Call SweetAlert("warning", "操作失败！", "请详细填写传真号码并保持至少在10个字符以上！", "false", "history.back()")
        response.End
    End If
    If Len(Trim(request.Form("Email")))<6 Then
		Call SweetAlert("warning", "操作失败！", "请详细填写电子邮箱地址并保持至少在6个字符以上！", "false", "history.back()")
        response.End
    End If
    If Len(Trim(request.Form("Keywords")))>250 Then
		Call SweetAlert("warning", "操作失败！", "请详细填写网站关键词并保持在250个字符以内！", "false", "history.back()")
        response.End
    End If
    If Len(Trim(request.Form("Descriptions")))>250 Then
		Call SweetAlert("warning", "操作失败！", "请详细填写网站描述并保持在250个字符以内！", "false", "history.back()")
        response.End
    End If
	If Len(Trim(request.Form("Stats")))>1000 Then
		Call SweetAlert("warning", "操作失败！", "请详细填写统计代码并保持在1000个字符以内！", "false", "history.back()")
        response.End
    End If
    Call SiteInfo()
    Dim rs, sql
    Set rs = server.CreateObject("adodb.recordset")
    sql = "select top 1 * from keyicms_Site"
    rs.Open sql, conn, 1, 3
	If Clng(rs("SiteStatus")) <> Clng(Request.Form("SiteStatus")) Then IsStatus = "注意你已改变网站的状态，请重新生成PC全站静态。"
	
    rs("SiteTitle") = Trim(Request.Form("SiteTitle"))
    rs("SiteUrl") = Trim(Request.Form("SiteUrl"))
    rs("ComName") = Trim(Request.Form("ComName"))
    rs("Address") = Trim(Request.Form("Address"))
    rs("Telephone") = Trim(Request.Form("Telephone"))
	rs("Mobile") = Trim(Request.Form("Mobile"))
    rs("Contacts") = Trim(Request.Form("Contacts"))
    rs("Fax") = Trim(Request.Form("Fax"))
    rs("Email") = Trim(Request.Form("Email"))
    rs("Keywords") = Trim(Request.Form("Keywords"))
    rs("Descriptions") = Trim(Request.Form("Descriptions"))
	rs("Stats") = Trim(Request.Form("Stats"))
    rs("IcpNumber") = Trim(Request.Form("IcpNumber"))
    rs("SiteLogo") = Trim(Request.Form("SiteLogo"))
    rs("Ico") = Trim(Request.Form("ico"))
	If Trim(Request.Form("SiteStatus")) = "-1" Then
		rs("SiteStatus") = Trim(Request.Form("SiteStatus"))
	Else
		rs("SiteStatus") = 0
	End If
    rs("Explain") = Trim(Request.Form("Explain"))
    rs("SiteDetail") = Trim(Request.Form("SiteDetail"))
    rs.update
    rs.Close
    Set rs = Nothing
	
	If Trim(Request.Form("SiteStatus")) = "-1" Then
		Call DoDelHtml("/Maintain.html")
	Else
		Explain = "<br/><br/><p style='color:red; font-size:14px; text-align:center'>"&Trim(Request.Form("Explain"))&"</p>"
		Call WriteToUTF("/Maintain.html", Explain)
	End If
	Call SweetAlert("success", "操作成功", Keyicms_Lang_MainSite(25)&" "&IsStatus, "false", "location.replace('Keyicms_MainSite.Asp')")
End Function

Function ViewSiteInfo()
    Dim rs, sql
    Set rs = server.CreateObject("adodb.recordset")
    sql = "select top 1 * from keyicms_Site"
    rs.Open sql, conn, 1, 1
    If rs.bof And rs.EOF Then
        response.Write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>"&Keyicms_Lang_System(28)&"</font></div>")
        response.End
    Else
		SiteTitle = rs("SiteTitle")
		Keywords = rs("Keywords")
		Descriptions = rs("Descriptions")
		Stats = rs("Stats")
		SiteUrl = rs("SiteUrl")
		ComName = rs("ComName")
		Address = rs("Address")
		Telephone = rs("Telephone")
		Mobile = rs("Mobile")
		Contacts = rs("Contacts")
		Fax = rs("Fax")
		Email = rs("Email")
		IcpNumber = rs("IcpNumber")
		SiteLogo = rs("SiteLogo")
		SiteDetail = rs("SiteDetail")
		SiteStatus = rs("SiteStatus")
		Explain = rs("Explain")
        rs.Close : Set rs = Nothing
    End If
End Function

%>