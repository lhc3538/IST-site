<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../Include/system_language.asp" -->
<!--#include file="CheckAdmin.asp"-->
<!--#include file="../Include/Version.asp" -->
<%
m_SQL = "select count(*) from keyicms_Admin"
set rs = conn.Execute(m_SQL)
m_ManageNumber = rs(0)
rs.Close
set rs = Nothing

m_SQL = "select count(*) from keyicms_Message"
set rs = conn.Execute(m_SQL)
m_Message = rs(0)
rs.Close
set rs = Nothing

m_SQL = "select count(*) from keyicms_About"
set rs = conn.Execute(m_SQL)
m_About = rs(0)
rs.Close
set rs = Nothing

m_SQL = "select count(*) from keyicms_Download"
set rs = conn.Execute(m_SQL)
m_Download = rs(0)
rs.Close
set rs = Nothing

m_SQL = "select count(*) from keyicms_FriendLink"
set rs = conn.Execute(m_SQL)
m_FriendLink = rs(0)
rs.Close
set rs = Nothing

m_SQL = "select count(*) from keyicms_Job"
set rs = conn.Execute(m_SQL)
m_Job = rs(0)
rs.Close
set rs = Nothing

m_SQL = "select count(*) from keyicms_Case"
set rs = conn.Execute(m_SQL)
m_Case = rs(0)
rs.Close
set rs = Nothing

m_SQL = "select count(*) from keyicms_News"
set rs = conn.Execute(m_SQL)
m_News = rs(0)
rs.Close
set rs = Nothing

m_SQL = "select count(*) from keyicms_Product"
set rs = conn.Execute(m_SQL)
m_Product = rs(0)
rs.Close
set rs = Nothing

m_SQL = "select count(*) from keyicms_App"
set rs = conn.Execute(m_SQL)
m_App = rs(0)
rs.Close
set rs = Nothing

m_SQL = "select count(*) from keyicms_Template"
set rs = conn.Execute(m_SQL)
m_Template = rs(0)
rs.Close
set rs = Nothing
%>

<!DOCTYPE html>
<html lang="zh-cn">
<head>
<meta charset="utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="description" content="<%=Keyicms_Lang_Copyright(13)%>">
<meta name="author" content="<%=Keyicms_Lang_Copyright(3)%>">
<link href="favicon.ico" rel="shortcut icon" />
<title><%=Keyicms_Lang_Copyright(1)%> <%=Str_Soft_Version%></title>
<!--#Include file="Keyicms_css.Asp"-->
</head>

<body>
<!--#Include file="Keyicms_Top.Asp"-->
<div id="wrap">
	<!--#Include file="Keyicms_Meun.Asp"-->

    <!-- start: Content -->
    <div class="main">
        <div class="row ">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i> 当前位置</li>
                <li class="active">后台首页</li>
            </ol>	
        </div>	
        
        <div class="row ">
            <div class="col-lg-12 hideInIE8">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-edit"></i><span class="break"></span>快捷菜单</h2> 
                        <div class="panel-actions">
                            <a href="#" class="btn-minimize"><i class="fa fa-chevron-up"></i></a>
                            <a href="#" class="btn-close"><i class="fa fa-times"></i></a>
                        </div>
                    </div>
                    <div class="panel-body">
                    	<%If Conn.execute("select * from keyicms_System where DataForm=3")("ViewFlag") Then%>
                        <a class="ShortCut1" href="Ky_ProductEdit.Asp?Result=Add">  
                            <i class="fa fa-inbox"></i>
                            <div class="title">添加学院图片</div>
                        </a>
                        <%End If%>					
                    	<%If Conn.execute("select * from keyicms_System where DataForm=2")("ViewFlag") Then%>
                        <a class="ShortCut2" href="Ky_NewsEdit.Asp?Result=Add">  
                            <i class="fa fa-book"></i>
                            <div class="title">添加新闻等信息</div>
                        </a>					
                        <%End If%>					
                    	<%If Conn.execute("select * from keyicms_System where DataForm=4")("ViewFlag") Then%>
                        <a class="ShortCut3" href="Ky_CaseEdit.Asp?Result=Add">  
                            <i class="fa fa-picture-o"></i>
                            <div class="title">添加资源扩展</div>
                        </a>					
                        <%End If%>					
                        <a class="ShortCut4" href="Ky_App.Asp">  
                            <i class="fa fa-suitcase"></i>
                            <div class="title">网站应用</div>
                        </a>					
                    </div>
                </div>
            </div><!--/col-->
        </div>	

        <div class="row">
            <div class="col-lg-3 col-md-4">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-bar-chart-o"></i>信息统计</h2>
                    </div>
                    <div class="panel-body">
                        <table id="user" class="table table-bordered table-striped" style="clear: both">
                        <tr>         
                            <td width="65%">管理员</td>
                            <td width="35%"><%=m_ManageNumber%>个</td>
                        </tr>
                        <%
						set rs = server.CreateObject("adodb.recordset")
						rs.open "select * from keyicms_System where ViewFlag Order by ID asc", conn, 1, 2
						do until rs.eof
							DataForm = Conn.execute("select * from keyicms_DataForm where ID="&rs("DataForm"))("DataForm")
							m_rs = Conn.execute("select count(*) from "&DataForm)(0)
							response.Write "<tr>"         
							response.Write "	<td>"&rs("SystemName")&"</td>"
							response.Write "	<td>"&m_rs&"条</td>"
							response.Write "</tr>"  
							rs.movenext
						loop
						rs.close : set rs = nothing
						%>
                        <tr>         
                            <td>网站应用</td>
                            <td><%=m_App%>条</td>
                        </tr> 
                        <tr>         
                            <td>在线留言</td>
                            <td><%=m_Message%>条</td>
                        </tr>                                        
                        </table>
                    </div>
                </div>	
            </div>
                    
            <div class="col-md-6">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <ul class="nav tab-menu nav-tabs" id="myTab">
                            <li class=""><a href="#info">系统基本信息</a></li>
                        </ul>
                    </div>
                    <div class="panel-body">
                        <div id="myTabContent" class="tab-content">
                            <div class="tab-pane" id="info">
							<%
                            set rs = server.createobject("adodb.recordset") 
                            sql = "select top 1 * from keyicms_AdminLog where [id] < (select max(id) from keyicms_AdminLog) order by [id] desc" 
                            rs.open sql,conn,1,1
							If rs.eof Then
								LoginTime = "暂无"
								LoginIP = "暂无"
							Else
								LoginTime = rs("LoginTime")
								LoginIP = rs("LoginIP")
							End If
                            rs.close : set rs = nothing

                            set rs1 = server.createobject("adodb.recordset") 
                            sql = "select top 1 * from keyicms_AdminLog order by id desc " 
                            rs1.open sql,conn,1,1
                            %>														
                            <table id="user" class="table table-bordered table-striped" style="clear: both">
                            <tr>         
                                <td width="35%">登陆角色</td>
                                <td width="65%"><%=session("UserName")%></td>
                            </tr>
                            <tr>         
                                <td>本次登陆时间</td>
                                <td><%=rs1("LoginTime")%></td>
                            </tr>  
                            <tr>         
                                <td>本次登陆IP地址</td>
                                <td><%=rs1("LoginIP")%> </td>
                            </tr>
                            <tr>         
                                <td>最近一次登陆时间</td>
                                <td><%=LoginTime%></td>
                            </tr> 
                            <tr>         
                                <td>最近一次登陆IP地址</td>
                                <td><%=LoginIP%></td>
                            </tr>  
                            <tr>         
                                <td>服务器名称</td>
                                <td><%=Request.ServerVariables("SERVER_NAME")%> (IP:<%=Request.ServerVariables("LOCAL_ADDR")%>) 端口:<%=Request.ServerVariables("SERVER_PORT")%></td>
                            </tr>
                            <tr>         
                                <td>服务器脚本引擎</td>
                                <td><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %> </td>
                            </tr> 
                            <tr>         
                                <td>站点物理路径</td>
                                <td><%=request.ServerVariables("APPL_PHYSICAL_PAth")%></td>
                            </tr> 
                            <tr>         
                                <td>IIS版本</td>
                                <td><%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
                            </tr> 
                            <tr>         
                                <td>脚本超时时间</td>
                                <td><%=Server.ScriptTimeout%> s</td>
                            </tr>                                       
                            </table>
                            <%
                            rs1.close : set rs1 = nothing
                            %>
                            </div>
                        </div>
                    </div>
                </div>				
            </div>              
        </div>
    </div>
    <!-- end: Content -->
	<!--#Include file="Keyicms_End.Asp"-->				
</div>
<!--#Include file="Keyicms_Js.Asp"-->	
</body>
</html>