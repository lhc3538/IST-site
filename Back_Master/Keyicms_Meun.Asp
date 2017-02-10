<!-- start: Header -->
<div class="navbar" role="navigation">
    <!--logo start-->
    <div class="profile">
        <div class="logo"><a href="Keyicms_Index.Asp"><img src="ke01/img/logo.png" alt=""></a></div>
    </div>
    <!--logo end--> 
    
    <div class="container-fluid">
        <ul class="nav navbar-nav navbar-actions navbar-left">
            <li class="visible-md visible-lg"><a href="JavaScript:void(0)" id="main-menu-toggle"><i class="fa fa-bars"></i></a></li>
            <li class="visible-xs visible-sm"><a href="JavaScript:void(0)" id="sidebar-menu"><i class="fa fa-bars"></i></a></li>
        </ul>
        <ul class="nav navbar-nav navbar-right">
            <li class="dropdown visible-md visible-lg">
                <a href="index.html#" class="dropdown-toggle" data-toggle="dropdown"><i class="fa fa-wrench"></i><span class="badge">!</span></a>
                <ul class="dropdown-menu">
                    <li class="dropdown-menu-header text-center">
                        <strong>功能菜单</strong>
                    </li>
                    <li><a href="../" target="_blank"><i class="fa fa-home"></i>前台首页<span class="label label-info"></span></a></li>
                    <%If mStatus Then%><li><a href="../m" target="_blank"><i class="fa fa-refresh"></i>手机网站首页</a></li><%End If%>
                    <li><a href="./"><i class="fa  fa-list-alt"></i>后台首页<span class="label label-success"></span></a></li>
                    <li><a href="javascript:AdminOut()" target="_top"><i class="fa fa-power-off"></i>安全退出<span class="label label-warning"></span></a></li>
                </ul>
            </li>
            <li><a href="javascript:AdminOut()" target="_top"><i class="fa fa-power-off"></i></a></li>
        </ul>
    </div>
</div>
<!-- end: Header -->

<div id="nav">   
    <!--navigation start-->
    <ul class="navigation">
        <li><a href="Keyicms_Index.Asp" <%=Currentmenu("Keyicms_Index")%>><i class="fa fa-home"></i><span>后台首页</span></a></li>
        <li class="sub">
            <a href="#" <%=Currentmenu("Keyicms_ExtraSite.Asp, Keyicms_MainSite, Keyicms_MobileSite")%>><i class="fa fa-cogs"></i><span>核心设置</span></a>
            <ul class="navigation-sub">
            	<li><a href="Keyicms_ExtraSite.Asp"> <i class="fa fa-home"></i>    <span>网站系统设置</span></a></li>
                <li><a href="Keyicms_MainSite.Asp">  <i class="fa fa-gear"></i>    <span>网站参数设置</span></a></li>
            </ul>
        </li>
        <li class="sub">
            <a href="#" <%=Currentmenu("Ky_About, Ky_News, Ky_Product, Ky_Case, Ky_Down, Ky_Job, Ky_System, Ky_Navigation")%>><i class="fa fa-list-alt"></i><span>功能模块</span></a>
            <ul class="navigation-sub">
            	<li><a href="Ky_System.Asp"> <i class="fa fa-list-alt"></i> <span>功能模块管理</span></a></li>
            	<%
                set rs = server.CreateObject("adodb.recordset")
				rs.open "select * from keyicms_System where ViewFlag Order by ID asc", conn, 1, 2
				do until rs.eof
					response.Write "<li><a href='"&rs("SystemPath")&"'> <i class='fa fa-clipboard'></i> <span>"&rs("SystemName")&"</span></a></li>"
					rs.movenext
				loop
				rs.close : set rs = nothing
                %>
                <li><a href="Ky_Navigation.Asp">     <i class="fa fa-sitemap"></i> <span>网站导航设置</span></a></li>
            </ul>
        </li>
        <li><a href="Ky_App.Asp" <%=Currentmenu("Ky_App, Ky_FriendLink, Ky_ad, Ky_Slide, Ky_Message, Ky_Map, Ky_Contact")%>><i class="fa fa-briefcase"></i><span>网站应用</span></a></li>
        
        <li class="sub">
            <a href="#" <%=Currentmenu("Ky_Admin, Keyicms_ManageLog")%>><i class="fa fa-puzzle-piece"></i><span>管理员管理</span></a>
            <ul class="navigation-sub">
                <li><a href="Ky_Admin.Asp">                <i class="fa fa-user"></i>     <span>网站管理员管理</span></a></li>
                <li><a href="Ky_AdminEdit.Asp?Result=Add"> <i class="fa fa-users"></i>    <span>添加网站管理员</span></a></li>
                <li><a href="Keyicms_ManageLog.Asp">       <i class="fa fa-calendar"></i> <span>后台登陆日志</span></a></li>
            </ul>
        </li>
    </ul>
    <!--navigation end--> 
</div>
<%
Function Currentmenu(str)
Currentstr = "Class='active'"
If Instr(Request.ServerVariables("URL"),str) Then Currentmenu = Currentstr
If Instr(str,", ") > 0 Then
	cArr = split(str,", ")
	For i = 0 to UBound(cArr)
		if Instr(Request.ServerVariables("URL"),cArr(i)) > 0 then
			Currentmenu = Currentstr
		end if
	Next
End If
End Function
%>