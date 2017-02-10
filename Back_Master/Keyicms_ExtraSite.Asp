<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../Include/Version.asp" -->
<!--#include file="../Include/system_language.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
Select Case request.QueryString("Action")
	Case "Save"
		SaveSiteInfo
	Case "SaveConst"
		SaveConstInfo
	Case Else
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
</head>

<body>
<!--#Include file="Keyicms_Top.Asp"-->

<div id="wrap">
	<!--#Include file="Keyicms_Meun.Asp"-->
	<%
	If InStr(CompanyAdminPurview, "|16,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style=""color:red; font-size:9pt;"")>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If
	%>
    <!-- start: Content -->
    <div class="main">
        <div class="row ">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i> <%=Keyicms_Lang_Copyright(14)%></li>
                <li><%=Keyicms_Lang_Copyright(2)%></li>
                <li class="active"><%=Keyicms_Lang_System(2)%></li>
            </ol>
        </div>	
		<div class="row">
            <div class="col-md-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-gears"></i><span class="break"></span><strong><%=Keyicms_Lang_System(2)%></strong></h2>
                    </div>
                    <div class="panel-body">
                    <form action="?Action=SaveConst" method="post" name="ConstForm" class="form-horizontal ">
                        <div class="panel-heading">
                            <h2><i class="fa  fa-exclamation-circle"></i><span class="break"></span><strong><%=Keyicms_Lang_sys_Parameters(31)%></strong></h2>
                        </div> 
                        <div class="form-group">
                            <label class="col-md-2 control-label"><%=Keyicms_Lang_sys_Parameters(2)%>：</label>
                            <div class="col-md-10">
                                <label class="radio-inline" for="EnableCopy1">
                                    <input name="EnableCopy" id="EnableCopy1" type="radio" value="True" <%If EnableCopy = True Then%> checked<%End If%>> <%=Keyicms_Lang_System(30)%>
                                </label>
                                <label class="radio-inline" for="EnableCopy2">
                                    <input name="EnableCopy" id="EnableCopy2" type="radio" value="False" <%If EnableCopy = False Then%> checked<%End If%>> <%=Keyicms_Lang_System(31)%> 
                                </label>
                                <span class="help-block"></span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label"><%=Keyicms_Lang_sys_Parameters(3)%>：</label>
                            <div class="col-md-10">
                                <label class="radio-inline" for="EnableBackstageCode1">
                                    <input name="EnableBackstageCode" id="EnableBackstageCode1" type="radio" value="True" <%If EnableBackstageCode = True Then%> checked<%End If%>> <%=Keyicms_Lang_System(30)%>
                                </label>
                                <label class="radio-inline" for="EnableBackstageCode2">
                                    <input name="EnableBackstageCode" id="EnableBackstageCode2" type="radio" value="False" <%If EnableBackstageCode = False Then%> checked<%End If%>> <%=Keyicms_Lang_System(31)%> 
                                </label>
                                <span class="help-block"></span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label"><%=Keyicms_Lang_sys_Parameters(6)%>：</label>
                            <div class="col-md-10">
                                <label class="radio-inline" for="EnableSiteManageCode1">
                                    <input name="EnableSiteManageCode" id="EnableSiteManageCode1" type="radio" value="True" <%If EnableSiteManageCode = True Then%> checked<%End If%>> <%=Keyicms_Lang_System(30)%>
                                </label>
                                <label class="radio-inline" for="EnableSiteManageCode2">
                                    <input name="EnableSiteManageCode" id="EnableSiteManageCode2" type="radio" value="False" <%If EnableSiteManageCode = False Then%> checked<%End If%>> <%=Keyicms_Lang_System(31)%>
                                </label>
                                <span class="help-block"></span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label"><%=Keyicms_Lang_sys_Parameters(7)%>：</label>
                            <div class="col-md-3">
                                <input name="SiteManageCode" type="text" id="SiteManageCode" class="form-control" value="<%=SiteManageCode%>" maxlength="20"> 
                                <span class="help-block"><font color="red">*</font>请填写后台登陆认证码</span>
                            </div>
                        </div>
                        
                        <div class="panel-heading">
                            <h2><i class="fa fa-file-text"></i><span class="break"></span><strong><%=Keyicms_Lang_sys_Parameters(33)%></strong></h2>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label"><%=Keyicms_Lang_sys_Parameters(16)%>：</label>
                            <div class="col-md-10">
                                <label class="radio-inline" for="ClassSeoISPY1">
                                    <input name="ClassSeoISPY" id="ClassSeoISPY1" type="radio" value="True" <%If ClassSeoISPY = True Then%> checked<%End If%>> <%=Keyicms_Lang_System(30)%>
                                </label>
                                <label class="radio-inline" for="ClassSeoISPY2">
                                    <input name="ClassSeoISPY" id="ClassSeoISPY2" type="radio" value="False" <%If ClassSeoISPY = False Then%> checked<%End If%>> <%=Keyicms_Lang_System(31)%>
                                </label>
                                <span class="help-block">如您启用汉字转换拼音功能，则选项必须为开启状态。</span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label"><%=Keyicms_Lang_sys_Parameters(17)%>：</label>
                            <div class="col-md-4">
                                <input name="Separated" type="text" id="Separated" class="form-control" value="<%=Separated%>">
                                <span class="help-block">如：News-1.html中的“-”即为分隔符</span>
                            </div>
                        </div>

                        <div class="panel-heading">
                            <h2><i class="fa  fa-columns"></i><span class="break"></span><strong>同步邮箱通知</strong></h2>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label">是否启用同步邮箱通知：</label>
                            <div class="col-md-4">
                                <label class="radio-inline" for="JMailDisplay1" id="Show">
                                    <input name="JMailDisplay" id="JMailDisplay1" type="radio" value="True" <%If JMailDisplay = True Then%> checked<%End If%>> <%=Keyicms_Lang_System(30)%>
                                </label>
                                <label class="radio-inline" for="JMailDisplay2" id="Hide">
                                    <input name="JMailDisplay" id="JMailDisplay2" type="radio" value="False" <%If JMailDisplay = False Then%> checked<%End If%>> <%=Keyicms_Lang_System(31)%>
                                </label>
                                <span class="help-block">开启此功能后，客户的留言会发送到管理员设置好的信箱中。</span>
                            </div>
                        </div>
                        <div class="form-group" name="ShowFlag" <%If JMailDisplay = False Then response.Write "style='display:none'"%>>
                            <label class="col-md-2 control-label">SMTP服务器：</label>
                            <div class="col-md-4">
                                <input name="JMailSMTP" type="text" id="JMailSMTP" class="form-control" value="<%=JMailSMTP%>">
                                <span class="help-block"></span>
                            </div>
                        </div>
                        <div class="form-group" name="ShowFlag" <%If JMailDisplay = False Then response.Write "style='display:none'"%>>
                            <label class="col-md-2 control-label">SMTP服务器用户名：</label>
                            <div class="col-md-4">
                                <input name="JMailUser" type="text" id="JMailUser" class="form-control" value="<%=JMailUser%>">
                                <span class="help-block"></span>
                            </div>
                        </div>
                        <div class="form-group" name="ShowFlag" <%If JMailDisplay = False Then response.Write "style='display:none'"%>>
                            <label class="col-md-2 control-label">SMTP服务器密码：</label>
                            <div class="col-md-4">
                                <input name="JMailPass" type="text" id="JMailPass" class="form-control" value="<%=JMailPass%>">
                                <span class="help-block"></span>
                            </div>
                        </div>
                        <div class="form-group" name="ShowFlag" <%If JMailDisplay = False Then response.Write "style='display:none'"%>>
                            <label class="col-md-2 control-label">发送人：</label>
							<div class="col-md-4">
                            	<input name="JMailName" type="text" id="JMailName" class="form-control" value="<%=JMailName%>">
                                <span class="help-block"></span>
                            </div>
                        </div>
                        <div class="form-group" name="ShowFlag" <%If JMailDisplay = False Then response.Write "style='display:none'"%>>
                            <label class="col-md-2 control-label">接收信箱：</label>
                            <div class="col-md-4">
								<input name="JMailInFrom" type="text" id="JMailInFrom" class="form-control" value="<%=JMailInFrom%>">
                                <span class="help-block">暂不支持QQ邮箱接收</span>
                            </div>
                        </div>
                        <div class="form-group" name="ShowFlag" <%If JMailDisplay = False Then response.Write "style='display:none'"%>>
                            <label class="col-md-2 control-label">发送信箱：</label>
                            <div class="col-md-4">
								<input name="JMailOutFrom" type="text" id="JMailOutFrom" class="form-control" value="<%=JMailOutFrom%>">
                                <span class="help-block"></span>
                            </div>
                        </div>
                        <div class="form-group" name="ShowFlag" <%If JMailDisplay = False Then response.Write "style='display:none'"%>>
                            <label class="col-md-2 control-label">反馈标题头：</label>
                            <div class="col-md-4">
                          		<input name="JMailTitle" type="text" id="JMailTitle" class="form-control" value="<%= JMailTitle %>">
                                <span class="help-block"></span>
                            </div>
                        </div>

                        <div class="panel-heading">
                            <h2><i class="fa  fa-columns"></i><span class="break"></span><strong><%=Keyicms_Lang_sys_Parameters(34)%></strong></h2>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label"><%=Keyicms_Lang_sys_Parameters(28)%>：</label>
                            <div class="col-md-10">
                                <label class="radio-inline" for="EnableViews1">
                                    <input name="EnableViews" id="EnableViews1" type="radio" value="True" <%If EnableViews = True Then%> checked<%End If%>> <%=Keyicms_Lang_System(30)%>
                                </label>
                                <label class="radio-inline" for="EnableViews2">
                                    <input name="EnableViews" id="EnableViews2" type="radio" value="False" <%If EnableViews = False Then%> checked<%End If%>> <%=Keyicms_Lang_System(31)%>
                                </label>
                                 <span class="help-block"></span>
                            </div>
                        </div>  
                        <div class="form-group">
                            <label class="col-md-2 control-label"><%=Keyicms_Lang_sys_Parameters(29)%>：</label>
                            <div class="col-md-10">
                                <label class="radio-inline" for="EnableSeo1">
                                    <input name="EnableSeo" id="EnableSeo1" type="radio" value="True" <%If EnableSeo = True Then%> checked<%End If%>> <%=Keyicms_Lang_System(30)%>
                                </label>
                                <label class="radio-inline" for="EnableSeo2">
                                    <input name="EnableSeo" id="EnableSeo2" type="radio" value="False" <%If EnableSeo = False Then%> checked<%End If%>> <%=Keyicms_Lang_System(31)%>
                                </label>
                                 <span class="help-block"></span>
                            </div>
                        </div>                                


                        <div class="form-group">
                            <div class="col-md-2 control-label"></div>
                            <div class="col-md-10">
                                <input name="submitSaveEdit" type="submit" id="submitSaveEdit" class="btn btn-danger btn-lg" value="<%=Keyicms_Lang_sys_Parameters(37)%>">
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
Function SaveConstInfo()
    Call SiteInfo()

    Set ASM = server.CreateObject("Adod" & "b.St" & "ream")
    ASM.Type = 2
    ASM.mode = 3
    ASM.charset = "utf-8"
    ASM.Open
    ASM.WriteText "<!--#include file=""Charset.Asp"" -->" & vbCrLf
    ASM.WriteText "<" & "%" & vbCrLf


    ASM.WriteText "Const EnableCopy = " & ReplaceConstChar(Trim(request("EnableCopy"))) & "" & vbCrLf
	ASM.WriteText "Const EnableBackstageCode = " & ReplaceConstChar(Trim(request("EnableBackstageCode"))) & "" & vbCrLf
    ASM.WriteText "Const EnableSiteManageCode = " & ReplaceConstChar(Trim(request("EnableSiteManageCode"))) & "" & vbCrLf
    ASM.WriteText "Const SiteManageCode = " & Chr(34) & ReplaceConstChar(Trim(request("SiteManageCode"))) & Chr(34) & "" & vbCrLf
    
    ASM.WriteText "Const ClassSeoISPY = " & ReplaceConstChar(Trim(request("ClassSeoISPY"))) & "" & vbCrLf
    ASM.WriteText "Const Separated = " & Chr(34) & ReplaceConstChar(Trim(request("Separated"))) & Chr(34) & "" & vbCrLf

    ASM.WriteText "Const JMailDisplay = " & ReplaceConstChar(Trim(request("JMailDisplay"))) & "" & vbCrLf
    ASM.WriteText "Const JMailSMTP = " & Chr(34) & ReplaceConstChar(Trim(request("JMailSMTP"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const JMailUser = " & Chr(34) & ReplaceConstChar(Trim(request("JMailUser"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const JMailPass = " & Chr(34) & ReplaceConstChar(Trim(request("JMailPass"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const JMailName = " & Chr(34) & ReplaceConstChar(Trim(request("JMailName"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const JMailInFrom = " & Chr(34) & ReplaceConstChar(Trim(request("JMailInFrom"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const JMailOutFrom = " & Chr(34) & ReplaceConstChar(Trim(request("JMailOutFrom"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const JMailTitle = " & Chr(34) & ReplaceConstChar(Trim(request("JMailTitle"))) & Chr(34) & "" & vbCrLf
	
    ASM.WriteText "Const EnableViews = " & ReplaceConstChar(Trim(request("EnableViews"))) & "" & vbCrLf
    ASM.WriteText "Const EnableSeo = " & ReplaceConstChar(Trim(request("EnableSeo"))) & "" & vbCrLf
	If Trim(request("mStatus")) = "" Then
		ASM.WriteText "Const mStatus = False" & vbCrLf
	Else
		ASM.WriteText "Const mStatus = " & ReplaceConstChar(Trim(request("mStatus"))) & "" & vbCrLf
	End If
	
    ASM.WriteText "Const SysRootDir = " & Chr(34) & "/" & Chr(34) & vbCrLf
    ASM.WriteText "Const HTMLName = " & Chr(34) & "html" & Chr(34) & vbCrLf
    ASM.WriteText "Const KeyicmsV = " & Chr(34) & "V1.06 RC" & Chr(34) & vbCrLf
	
    ASM.WriteText "%" & ">"
    ASM.SaveToFile Server.mappath("../Include/Const.Asp"), 2
    ASM.flush
    ASM.Close
    Set ASM = Nothing
	
	Call SweetAlert("success", "操作成功", Keyicms_Lang_sys_Parameters(38), "false", "location.replace('Keyicms_ExtraSite.Asp')")
End Function

Sub FsoDelHtml(HTMLName)
    Dim Fso, FsoOut, File
    Set Fso = Server.CreateObject("Scripting.FileSystemObject")
    Set FsoOut = Fso.GetFolder(Server.Mappath(SysRootDir))
    For Each File In FsoOut.Files
        If Right(File.Name, Len(File.Name) - InStrRev(File.Name, ".")) = HTMLName And HTMLName <> "asp" Then
            Response.Write "<span style=""color:red; padding-left: 18px"">" & File.Name & "</span> "&Keyicms_Lang_sys_Parameters(39)&"…<br />"
            Fso.DeleteFile File.Path, True
        End If
    Next
    Set FsoOut = Nothing
    Set Fso = Nothing
End Sub
%>