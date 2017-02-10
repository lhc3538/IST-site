<!--#include file="../Include/Const.Asp"-->
<!--#include file="../Include/ConnSiteData.Asp"-->
<!--#include file="../Include/Md5.Asp"-->
<!--#include file="../Include/Keyicms_safe.asp"-->
<%
Dim LoginName, LoginPassword, AdminName, Password, AdminPurview, Working, UserName, rs, sql
LoginName = ReplaceBadChar(Trim(request("UserName")))
LoginPassword = Md5(ReplaceBadChar(request("PassWord")))
CheckCode = LCase(ReplaceBadChar(Trim(Request("CheckCode"))))
AdminLoginCode = Trim(Request("AdminLoginCode"))

Dim ErrMsg(6)
ErrMsg(0) = 0 '成功登陆！

ErrMsg(1) = 1 '您在管理登录停留的时间过长，导致验证码失效。请重新返回登录页面进行登录。
ErrMsg(2) = 2 '您输入的验证码和系统产生的不一致，请重新输入。
ErrMsg(3) = 3 '您输入的后台管理认证码不对，请重新输入。
ErrMsg(4) = 4 '无此帐号，请返回检查！
ErrMsg(5) = 5 '密码错误，请返回检查！
ErrMsg(6) = 6 '帐号被禁用！

If EnableBackstageCode = True Then
	If Trim(Session("CheckCode")) = "" Then
		response.Write ErrMsg(1)
		response.End
	End If
	If CheckCode <> Session("CheckCode") Then
		response.Write ErrMsg(2)
		response.End
	End If
End If
	
If EnableSiteManageCode = True And AdminLoginCode <> SiteManageCode Then
	response.Write ErrMsg(3)
	response.End
End If

Set rs = server.CreateObject("adodb.recordset")
sql = "select * from keyicms_Admin where AdminName='"&LoginName&"'"
rs.Open sql, conn, 1, 3
If rs.EOF Then
    response.Write ErrMsg(4)
    response.End
Else
    AdminName = rs("AdminName")
    Password = rs("Password")
    AdminPurview = rs("AdminPurview")
    Working = rs("Working")
    UserName = rs("UserName")
End If
If LoginPassword<>Password Then
    response.Write ErrMsg(5)
    response.End
End If
If Not Working Then
    response.Write ErrMsg(6)
    response.End
End If
If LoginName = AdminName And LoginPassword = Password Then
    rs("LastLoginTime") = Now()
    rs("LastLoginIP") = Request.ServerVariables("Remote_Addr")
    rs.update
    rs.Close
    Set rs = Nothing
    session("AdminName")=AdminName
    session("UserName")=UserName
    session("AdminPurview")=AdminPurview
    session("LoginSystem")="Succeed"
    Response.Cookies("CompanyZY")("AdminName") = AdminName
    Response.Cookies("CompanyZY")("AdminPurview") = AdminPurview
    Response.Cookies("CompanyZY")("LoginSystem") = "Succeed"
    Response.Cookies("CompanyZY")("AdminLoginCode") = AdminLoginCode
    Response.Cookies("CompanyZY")("LoginIP") = Request.ServerVariables("Remote_Addr")
    Dim LoginIP, LoginTime, LoginSoft
    LoginIP = Request.ServerVariables("Remote_Addr")
    LoginSoft = Request.ServerVariables("http_user_agent")
    LoginTime = Now()
	
    Set rs = server.CreateObject("adodb.recordset")
    sql = "select * from keyicms_AdminLog"
    rs.Open sql, conn, 1, 3
    rs.addnew
    rs("AdminName") = AdminName
    rs("UserName") = UserName
    rs("LoginIP") = LoginIP
    rs("LoginSoft") = LoginSoft
    rs("LoginTime") = LoginTime
    rs.update
    rs.Close : Set rs = Nothing
	
    response.Write ErrMsg(0)
    response.End
End If
%>