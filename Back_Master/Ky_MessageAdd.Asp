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
'神笔(30)天成网络享有软件著作权，如果您擅自删除这里的版权信息,我们将追究法律责任.
'未经公司同意禁止更改以下代码.
'*******************************************************
Dim ID, MesName, Content
Dim mLinkman, mTelephone, mEmail, AddTime
ID = request.QueryString("ID")
conn.Execute "update Keyicms_Message set SeeFlag = -1 where ID="&ID
Call MesEdit()
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
	If InStr(CompanyAdminPurview, "|13,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style=""color:red; font-size:14px;"")>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If
	%>
    <div class="main">
        <div class="row">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i> 当前位置</li>
                <li>企业应用</li>
                <li class="active">查看留言信息</li>
            </ol>	
        </div>	
        <div class="row">
            <div class="col-md-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-gears"></i><span class="break"></span><strong>查看留言信息</strong></h2>
                    </div>
                    <form name="editForm" method="post" action="Ky_JobEdit.Asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>" class="form-horizontal">
                    <div class="panel-body">
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">留言标题：</label>
                            <div class="col-md-4">
								<p class="form-control-static"><%=MesName%></p>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">留言人：</label>
                            <div class="col-md-4">
								<p class="form-control-static"><%=mLinkman%></p>
                            </div>
                        </div> 
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">联系电话：</label>
                            <div class="col-md-4">
								<p class="form-control-static"><%=mTelephone%></p>
                            </div>
                        </div> 
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">电子信箱：</label>
                            <div class="col-md-4">
								<p class="form-control-static"><%=mEmail%></p>
                            </div>
                        </div> 
                        <div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">留言时间：</label>
                            <div class="col-md-4">
								<p class="form-control-static"><%=AddTime%></p>
                            </div>
                        </div> 
						<div class="form-group">
                            <label class="col-md-2 control-label" for="text-input">留言内容：</label>
                            <div class="col-md-4">
								<p class="form-control-static"><%=Content%></p>
                            </div>
                        </div>
                    </div>
                        <div class="form-group">
                            <div class="col-md-2 control-label"></div>
                            <div class="col-md-10">
                                <input type="button" class="btn btn-danger btn-lg" value="<%=Keyicms_Lang_System(32)%>" onClick="history.back(-1)">
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
Sub MesEdit()
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from keyicms_Message where ID="& ID
rs.Open sql, conn, 1, 1
MesName = rs("MesName")
Content = ReStrReplace(rs("Content"))
mLinkman = rs("Linkman")
mTelephone = rs("Telephone")
mEmail = rs("Email")
AddTime = rs("AddTime")
rs.Close
Set rs = Nothing
End Sub
%>