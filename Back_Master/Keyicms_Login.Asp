<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/System_Language.asp"-->
<!--#include file="../Include/Version.asp" -->
<!--#include file="../Include/Keyicms_safe.asp"-->
<%
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
%>
<!DOCTYPE html>
<html>
<head>
<title><%=Keyicms_Lang_Copyright(1)%> - <%=Str_Soft_Version%></title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta charset="utf-8">
<link href="ke01/css/bootstrap.css" rel="stylesheet">
<link href="ke01/css/lock.css" rel="stylesheet">
<link href="ke01/css/animate.css" rel="stylesheet">
<link href="ke01/css/font-awesome.min.css" rel="stylesheet">

<link rel="stylesheet" type="text/css" href="SweetAlert/sweet-alert.css">
<script src="SweetAlert/sweet-alert.min.js"></script>
</head>

<body onLoad="">
    <div class="logo">
        <h4><a href="#"><img src="ke01/img/adminlogo.png" alt=""></a></h4>
    </div>
    <div class="avatar">
        <img src="ke01/img/one.png" alt="">
    </div> 
    <form name="Login" onSubmit="return CheckForm()" action="JavaScript:void(0)" method="post">
    <div class="lock-holder">
    	<div class="lock-div">
            <div class="lock-group">   
                <div class="form-group pull-left input-username">
                    <div class="input-group">
                        <input name="UserName" type="text" class="form-control " placeholder="用户名">
                        <span class="input-group-addon"><i class="fa fa-user"></i></span>    
                    </div>
                </div>
                <i class="fa fa-ellipsis-h dot-left"></i>
                <i class="fa fa-ellipsis-h dot-right"></i>
                <div class="form-group pull-right input-password">
                    <div class="input-group">
                        <input name="PassWord" type="password" class="form-control " placeholder="密码" >
                        <span class="input-group-addon"><i class="fa fa-key"></i></span>
                    </div>
                </div>
            </div> 
            <%If EnableBackstageCode = True and EnableSiteManageCode = True Then%>
            <div class="lock-group">   
                <div class="form-group pull-left input-code">
                    <div class="input-group">
                        <input name="AdminLoginCode" type="password" class="form-control " placeholder="认证码" >
                        <span class="input-group-addon"><i class="fa fa-lock" ></i></span>
                    </div> 
                </div>
                <i class="fa fa-ellipsis-h dot-left"></i>
                <i class="fa fa-ellipsis-h dot-right"></i>
                <div class="form-group pull-right input-LoginCode">
                   <div class="input-group code">
                        <input name="CheckCode" type="text" class="form-control " placeholder="验证码">
                        <span class="input-group-addon"><a href="javascript:refreshimg()" title="看不清楚，换个图片。"><img name="CodeImg" src="../Include/CheckCode/CheckCode.asp"/></a></span>     
                    </div>
                </div> 
            </div>
            <%End If%>
        </div>
        <%If EnableSiteManageCode = True and EnableBackstageCode = False Then%>
        <div class="lock-group"> 
        	<i class="fa fa-ellipsis-h dot"></i>
        	<div class="form-group pull">  
                <div class="input-group">
                    <input name="AdminLoginCode" type="password" class="form-control " placeholder="认证码" >
                    <span class="input-group-addon"><i class="fa fa-lock"></i></span>
                </div>
            </div>
        </div>
        <%End If%>
        <%If EnableSiteManageCode = False and EnableBackstageCode = True Then%>
        <div class="lock-group"> 
        	<i class="fa fa-ellipsis-h dot"></i>
        	<div class="form-group pull">  
                <div class="input-group">
                    <input name="CheckCode" type="text" class="form-control " placeholder="验证码" >
                    <span class="input-group-addon"><a href="javascript:refreshimg()" title="看不清楚，换个图片。"><img name="CodeImg" src="../Include/CheckCode/CheckCode.asp"/></a></span>     
                </div>
            </div>
        </div>
        <%End If%>
    </div> 
    <div class="submit">
        <button type="submit" onClick=""  class="btn btn-primary btn-lg">登陆后台</button>   
    </div>
    </form>
    
    <div class="mask_div" id="bodyMask"></div>
    <div class="layerDiv" id="browser">
        <h2>噢，您是否知道您正在使用的浏览器无法完全支持我们的页面？</h2>
        <h3>很抱歉，由于采用了HTML5，当前浏览器无法完美呈现该页面。</h3>
        <div><SPAN>如果你还不知道什么是HTML5标准，请看<A href="http://zh.wikipedia.org/wiki/HTML5" target="_blank">维基百科</A>。</SPAN>建议您使用以下浏览器的最新版本：</div>
        <ul>
            <li><A title="谷歌Chrome浏览器" class="ico icoChrome" href="http://www.google.cn/chrome/intl/zh-CN/landing_chrome.html" target="_blank"></A></li>
            <li><A title="Opera浏览器" class="ico icoOpera" href="http://www.opera.com/" target="_blank"></A></li>
            <li><A title="火狐浏览器" class="ico icoFirefox" href="http://www.firefox.com.cn/download/" target="_blank"></A></li>
            <li><A title="Safar浏览器" class="ico icoSafar" href="http://www.apple.com.cn/safari/download/" target="_blank"></A></li>
        </ul>
        <p><a class="ico icoGt" href="JavaScript:void(0);">知道了，但请还是打开页面</a></p>
    </div>
</body>

<!--<script src="ke01/js/jquery.js"></script>-->
<script src="ke01/js/jquery-1.11.1.min.js" type="text/javascript"></script>
<script src="ke01/js/bootstrap.min.js"></script>
<script src="ke01/js/bootstrap-progressbar.js"></script>
<script src="ke01/js/loding.js" type="text/javascript"></script>
<script type="text/javascript">
function CheckForm(e){
	var UserName, PassWord, CheckCode, AdminLoginCode
	UserName = document.Login.UserName.value
	PassWord = document.Login.PassWord.value
	
	if(UserName == "") {
		sweetAlert("操作失败", "请输入用户名！", "error");
		document.Login.UserName.focus();
		return false;
	}
	if(PassWord == "") {
		sweetAlert("操作失败", "请输入密码！", "error");
		document.Login.PassWord.focus();
		return false;
	}
	<%If EnableSiteManageCode = True Then%>
	AdminLoginCode = document.Login.AdminLoginCode.value
	if (AdminLoginCode == "") {
		sweetAlert("操作失败", "请输入您的管理验证码！", "error");
		document.Login.AdminLoginCode.focus();
		return(false);
	}
	<%End If%>
	<%If EnableBackstageCode = True Then%>
	CheckCode = document.Login.CheckCode.value
	if (CheckCode == "") {
		sweetAlert("操作失败", "请输入您的验证码！", "error");
		document.Login.CheckCode.focus();
		return(false);
	}
	<%End If%>
	
	$.ajax({
		type: "Get",
		url: "CheckLogin.Asp",  
		data: "UserName="+UserName+"&PassWord="+PassWord+"&CheckCode="+CheckCode+"&AdminLoginCode="+AdminLoginCode,
		cache: false,
		beforeSend: function(){

		},
		success: function(ErrMsg){
			if(ErrMsg == 1){
				sweetAlert("操作失败", "您在管理登录停留的时间过长，导致验证码失效。请重新登录。", "error");
			}
			if(ErrMsg == 2){
				sweetAlert("操作失败", "您输入的验证码和系统产生的不一致，请重新输入。", "error");
			}
			if(ErrMsg == 3){
				sweetAlert("操作失败", "您输入的后台管理认证码不对，请重新输入。", "error");
			}
			if(ErrMsg == 4){
				sweetAlert("操作失败", "无此帐号，请检查！", "error");
			}
			if(ErrMsg == 5){
				sweetAlert("操作失败", "密码错误，请检查！", "error");
			}
			if(ErrMsg == 6){
				sweetAlert("操作失败", "帐号被禁用！", "error");
			}
			if(ErrMsg != 0){refreshimg();}

			if(ErrMsg == 0){
				$('.pull-left,.dot-left').addClass('animated fadeOutRight')
				$('.pull-right,.dot-right').addClass('animated fadeOutLeft')
				$('.btn-submit,.pull,.dot').addClass('animated fadeOutUp')
			
				setTimeout(function(){
					$('.avatar').addClass('avatar-top');
					$('.submit').html('<div class="progress"><div class="progress-bar progress-bar-success" aria-valuetransitiongoal="100"></div></div>');
					$('.progress .progress-bar').progressbar(); 
				},500);
			
				setTimeout(function () {
					window.location.href = 'Keyicms_Index.Asp';
				},1500);
			}
		}
	})
}
</script>
<script>
$(window).load(function(){
	NProgress.done();
	if (!window.applicationCache) {
		$("#bodyMask").css({height:$(document).height()});
		$("#bodyMask").show();
		$("#browser").show();
		$(".icoGt").click(function(){$("#bodyMask").hide();$("#browser").hide()});
	}
})
</script>
<script language="javascript" src="ke01/js/Admin.js"></script>
</html>