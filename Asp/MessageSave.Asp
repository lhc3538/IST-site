<!--#include file="../Include/Const.Asp"-->
<!--#include file="../Include/System_Language.asp" -->
<!--#include file="../Include/ConnSiteData.Asp"-->
<!--#include file="../Include/Keyicms_safe.Asp" -->
<%
Dim rs, sql
Dim MesName, Content, mLinkman, mTelephone, mEmail, CheckCode
MesName = Trim(request.Form("MesName"))
Content = Trim(request.Form("Content"))
mLinkman = Trim(request.Form("Linkman"))
mTelephone = Trim(request.Form("Telephone"))
mEmail = Trim(request.Form("Email"))
CheckCode = Trim(request.Form("CheckCode"))

Dim ErrMessage, ErrMsg(6), FindErr(6), i
ErrMsg(0) = "请正确填写留言主题。"
ErrMsg(1) = "请填写留言内容。"
ErrMsg(2) = "请正确填写联系电话。"
ErrMsg(3) = "请正确填写电子邮箱。"
ErrMsg(4) = "验证码不能为空，请返回检查。"
ErrMsg(5) = "您在页面停留的时间过长，导致验证码失效。\n请返回并刷新【"&Keyicms_Lang_Title(8)&"】页面！"
ErrMsg(6) = "您输入的验证码和系统产生的不一致，请重新输入。"

If Len(MesName)>100 Or Len(MesName) = 0 Then
    FindErr(0) = True
End If
If Len(Content)<1 Then
    FindErr(1) = True
End If
If Len(mTelephone)>50 Or Len(mTelephone) = 0 Then
    FindErr(2) = True
End If
If Len(mEmail)=0 Or Not IsValidEmail(mEmail) Then
    FindErr(3) = True
End If
If CheckCode = "" Then
    FindErr(4) = True
End If
If Trim(Session("CheckCode")) = "" Then
    FindErr(5) = True
End If
If CheckCode <> Session("CheckCode") Then
    FindErr(6) = True
End If
For i = 0 To UBound(FindErr)
    If FindErr(i) = True Then
        ErrMessage = ErrMessage+ErrMsg(i) + "\n"
    End If
Next
If Not (ErrMessage = "" Or IsNull(ErrMessage)) Then
	response.Write "<script language='javascript'>alert('"&ErrMessage&"');history.go(-1);</script/>"
    response.End
End If

Set rs = server.CreateObject("adodb.recordset")
sql = "select * from keyicms_Message"
rs.Open sql, conn, 1, 3
rs.addnew
rs("MesName") = StrReplace(MesName)
rs("Content") = StrReplace(Content)
rs("Linkman") = StrReplace(mLinkman)
rs("Telephone") = StrReplace(mTelephone)
rs("Email") = mEmail
rs("AddTime") = Now()
rs.update
rs.Close : Set rs = Nothing

Cont = "<strong>"&MesName&"</strong>"
Cont = Cont & "<p>姓名："&mLinkman&"</p>"
Cont = Cont & "<p>手机："&mTelephone&"</p>"
Cont = Cont & "<p>邮箱："&mEmail&"</p>"
Cont = Cont & "<p>内容："&Content&"</p>"

If JMailDisplay = True Then
    Call ShopSendEmail(JMailInFrom, JMailTitle, Cont)
Else
	response.Write "<script>alert('留言提交成功！');history.back();</script>"
End If

Function ShopSendEmail(strToEmail, strTitle, strSubject)
    On Error Resume Next
    Set jmail = Server.CreateObject("JMail.Message")
    jmail.Charset = "gb2312"
    jmail.ContentType = "text/html"
    jmail.from = JMailOutFrom
    jmail.silent = True
    jmail.Logging = True
    jmail.FromName = JMailName
    jmail.mailserverusername = JMailUser
    jmail.mailserverpassword = JMailPass
    jmail.AddRecipient strToEmail
    jmail.body = strSubject
    jmail.Subject = strTitle
    jmail.Priority = 1
    If Not jmail.Send (JMailSMTP) Then
        response.Write "<script>alert('留言提交成功，发送邮箱失败！');history.back();</script>"
    Else
        response.Write "<script>alert('留言提交成功，发送邮箱成功！');history.back();</script>"
    End If
    Set jmail = Nothing
End Function
%>