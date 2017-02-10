<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/SiteData.asp" -->

<%
Dim NewsDataShow__MMColParam
NewsDataShow__MMColParam = "1"
If (Request("MM_EmptyValue") <> "") Then 
  NewsDataShow__MMColParam = Request("MM_EmptyValue")
End If

Dim NewsDataShow
Dim NewsDataShow_numRows

Set NewsDataShow = Server.CreateObject("ADODB.Recordset")
NewsDataShow.ActiveConnection = MM_SiteData_STRING
NewsDataShow.Source = "SELECT ID, AddTime, NewsName, SortID, ViewFlag FROM keyicms_News WHERE SortID = " + Replace(NewsDataShow__MMColParam, "'", "''") + " ORDER BY AddTime DESC"
NewsDataShow.CursorType = 0
NewsDataShow.CursorLocation = 2
NewsDataShow.LockType = 1
NewsDataShow.Open()

NewsDataShow_numRows = 0

Dim NoticeShow__MMColParam
NoticeShow__MMColParam = "2"
If (Request("MM_EmptyValue") <> "") Then 
  NoticeShow__MMColParam = Request("MM_EmptyValue")
End If

Dim NoticeShow
Dim NoticeShow_numRows

Set NoticeShow = Server.CreateObject("ADODB.Recordset")
NoticeShow.ActiveConnection = MM_SiteData_STRING
NoticeShow.Source = "SELECT AddTime, ID, NewsName, SortID, ViewFlag FROM keyicms_News WHERE SortID = " + Replace(NoticeShow__MMColParam, "'", "''") + " ORDER BY AddTime DESC"
NoticeShow.CursorType = 0
NoticeShow.CursorLocation = 2
NoticeShow.LockType = 1
NoticeShow.Open()

NoticeShow_numRows = 0

Dim StudentShow__MMColParam
StudentShow__MMColParam = "3"
If (Request("MM_EmptyValue") <> "") Then 
  StudentShow__MMColParam = Request("MM_EmptyValue")
End If

Dim StudentShow
Dim StudentShow_numRows

Set StudentShow = Server.CreateObject("ADODB.Recordset")
StudentShow.ActiveConnection = MM_SiteData_STRING
StudentShow.Source = "SELECT AddTime, ID, NewsName, SortID, ViewFlag FROM keyicms_News WHERE SortID = " + Replace(StudentShow__MMColParam, "'", "''") + " ORDER BY AddTime DESC"
StudentShow.CursorType = 0
StudentShow.CursorLocation = 2
StudentShow.LockType = 1
StudentShow.Open()

StudentShow_numRows = 0

Dim DownloadShow
Dim DownloadShow_numRows

Set DownloadShow = Server.CreateObject("ADODB.Recordset")
DownloadShow.ActiveConnection = MM_SiteData_STRING
DownloadShow.Source = "SELECT AddTime, DownName, ID, ViewFlag FROM keyicms_Download"
DownloadShow.CursorType = 0
DownloadShow.CursorLocation = 2
DownloadShow.LockType = 1
DownloadShow.Open()

DownloadShow_numRows = 0

Dim ShowPicture
Dim ShowPicture_numRows

Set ShowPicture = Server.CreateObject("ADODB.Recordset")
ShowPicture.ActiveConnection = MM_SiteData_STRING
ShowPicture.Source = "SELECT ID, ProductName, SmallPic, ViewFlag FROM keyicms_Product ORDER BY AddTime DESC"
ShowPicture.CursorType = 0
ShowPicture.CursorLocation = 2
ShowPicture.LockType = 1
ShowPicture.Open()

ShowPicture_numRows = 0

Dim ShowSlide
Dim ShowSlide_numRows

Set ShowSlide = Server.CreateObject("ADODB.Recordset")
ShowSlide.ActiveConnection = MM_SiteData_STRING
ShowSlide.Source = "SELECT Pic, SlideName, State, Url FROM keyicms_Slide ORDER BY State ASC"
ShowSlide.CursorType = 0
ShowSlide.CursorLocation = 2
ShowSlide.LockType = 1
ShowSlide.Open()

ShowSlide_numRows = 0

Dim ShowNavFa__MMColParam
ShowNavFa__MMColParam = "0"
If (Request("MM_EmptyValue") <> "") Then 
  ShowNavFa__MMColParam = Request("MM_EmptyValue")
End If

Dim ShowNavFa
Dim ShowNavFa_numRows

Set ShowNavFa = Server.CreateObject("ADODB.Recordset")
ShowNavFa.ActiveConnection = MM_SiteData_STRING
ShowNavFa.Source = "SELECT HtmlNavUrl, NavName, OutUrl, ViewFlag FROM keyicms_Navigation WHERE ParentID = " + Replace(ShowNavFa__MMColParam, "'", "''") + " ORDER BY Sequence ASC"
ShowNavFa.CursorType = 0
ShowNavFa.CursorLocation = 2
ShowNavFa.LockType = 1
ShowNavFa.Open()

ShowNavFa_numRows = 0

Dim ShowContact
Dim ShowContact_numRows

Set ShowContact = Server.CreateObject("ADODB.Recordset")
ShowContact.ActiveConnection = MM_SiteData_STRING
ShowContact.Source = "SELECT Phone, QQ, User, WeChat FROM keyicms_Contact ORDER BY Sequence ASC"
ShowContact.CursorType = 0
ShowContact.CursorLocation = 2
ShowContact.LockType = 1
ShowContact.Open()

ShowContact_numRows = 0

Dim ShowSite
Dim ShowSite_numRows

Set ShowSite = Server.CreateObject("ADODB.Recordset")
ShowSite.ActiveConnection = MM_SiteData_STRING
ShowSite.Source = "SELECT * FROM keyicms_Site"
ShowSite.CursorType = 0
ShowSite.CursorLocation = 2
ShowSite.LockType = 1
ShowSite.Open()

ShowSite_numRows = 0

Dim ShowFriendLink
Dim ShowFriendLink_numRows

Set ShowFriendLink = Server.CreateObject("ADODB.Recordset")
ShowFriendLink.ActiveConnection = MM_SiteData_STRING
ShowFriendLink.Source = "SELECT LinkFace, LinkName, LinkType, LinkUrl, ViewFlag FROM keyicms_FriendLink"
ShowFriendLink.CursorType = 0
ShowFriendLink.CursorLocation = 2
ShowFriendLink.LockType = 1
ShowFriendLink.Open()

ShowFriendLink_numRows = 0

Dim ShowExp
Dim ShowExp_numRows

Set ShowExp = Server.CreateObject("ADODB.Recordset")
ShowExp.ActiveConnection = MM_SiteData_STRING
ShowExp.Source = "SELECT CaseName, Content, SortID, ViewFlag FROM keyicms_Case ORDER BY Sequence ASC"
ShowExp.CursorType = 0
ShowExp.CursorLocation = 2
ShowExp.LockType = 1
ShowExp.Open()

ShowExp_numRows = 0

Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 13
Repeat1__index = 0
NewsDataShow_numRows = NewsDataShow_numRows + Repeat1__numRows

Dim Repeat4__numRows
Dim Repeat4__index

Repeat4__numRows = 13
Repeat4__index = 0
DownloadShow_numRows = DownloadShow_numRows + Repeat4__numRows

Dim Repeat6__numRows
Dim Repeat6__index

Repeat6__numRows = -1
Repeat6__index = 0
ShowSlide_numRows = ShowSlide_numRows + Repeat6__numRows

Dim Repeat5__numRows
Dim Repeat5__index

Repeat5__numRows = -1
Repeat5__index = 0
ShowPicture_numRows = ShowPicture_numRows + Repeat5__numRows

Dim Repeat7__numRows
Dim Repeat7__index

Repeat7__numRows = -1
Repeat7__index = 0
ShowContact_numRows = ShowContact_numRows + Repeat7__numRows

Dim Repeat9__numRows
Dim Repeat9__index

Repeat9__numRows = -1
Repeat9__index = 0
ShowFriendLink_numRows = ShowFriendLink_numRows + Repeat9__numRows

Dim Repeat10__numRows
Dim Repeat10__index

Repeat10__numRows = -1
Repeat10__index = 0
ShowExp_numRows = ShowExp_numRows + Repeat10__numRows

Dim Repeat8__numRows
Dim Repeat8__index

Repeat8__numRows = -1
Repeat8__index = 0
ShowFriendLink_numRows = ShowFriendLink_numRows + Repeat8__numRows

Dim Repeat3__numRows
Dim Repeat3__index

Repeat3__numRows = 13
Repeat3__index = 0
NewsDataShow_numRows = NewsDataShow_numRows + Repeat3__numRows

Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = 13
Repeat2__index = 0
NewsDataShow_numRows = NewsDataShow_numRows + Repeat2__numRows
%>
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title><%=(ShowSite.Fields.Item("SiteTitle").Value)%></title>
<meta name="keywords" content="<%=(ShowSite.Fields.Item("Keywords").Value)%>" />
<meta name="description" content="<%=(ShowSite.Fields.Item("Descriptions").Value)%>" />
<link href="/Template/PC/Default/css/css.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="/Template/PC/Default/js/jquery-1-7-2.js"></script>
<script type="text/javascript" src="/Template/PC/Default/js/navScroll.js"></script>
<script type="text/javascript">
<!--
$(function(){
	$(".nav > ul > li").ScrollNav({
		scrollObj:".nav > .scrollObj",
		speed:"fast"
	});
});

function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>
<script src="/Template/PC/Default/js/jquery.kinMaxShow-1.1.min.js" type="text/javascript" charset="utf-8"></script>
</head>

<body>
<div class="main_div1">
    <div class="top">
        <div class="header">
            <img src="<%=(ShowSite.Fields.Item("SiteLogo").Value)%>" class="logo"/>
            <div class="head">
            <div class="headTxt">
                <a href="javascript:;"onclick="this.style.behavior='url(#default#homepage)';this.setHomePage(location.href);">设为首页</a>
                <a href="javascript:;" onClick="window.external.AddFavorite(location.href,document.title)">加入收藏</a> 
            </div>
            <form name="searchForm" method="post" onSubmit="return CheckForm()" action="JavaScript:void(0)" class="search">
                <input type="text" name="key" id="key" value="" />
                <input type="submit" name="button" id="button" value="" />
            </form>
            </div>
            <script>
            function CheckForm(){
                if (document.searchForm.key.value=='')
                {
                    alert('搜索内容不能为空');
                    return false
                }
                else
                {
                    window.location.href = '/Search/Index.Asp?key='+document.searchForm.key.value
                }
            }
            </script>
        </div>
    </div>
    <div class="gridNav">
        <div class="nav">
            <ul class="ulBox" id="ulBox">  <!-- 显示导航栏 -->
            	<%
				Dim ShowVavChild__MMColParam
				Dim ShowVavChild
				Dim ShowVavChild_numRows
				%>

				<% for nav_i = 1 to 8 %>
                	<li><a href="<%=(ShowNavFa.Fields.Item("HtmlNavUrl").Value)%><%=(ShowNavFa.Fields.Item("OutUrl").Value)%>"><%=(ShowNavFa.Fields.Item("NavName").Value)%></a>	
					<ul>
						<%
						ShowVavChild__MMColParam = cstr(nav_i)
						If (Request("MM_EmptyValue") <> "") Then 
						  ShowVavChild__MMColParam = Request("MM_EmptyValue")
						End If
						
						Set ShowVavChild = Server.CreateObject("ADODB.Recordset")
						ShowVavChild.ActiveConnection = MM_SiteData_STRING
						ShowVavChild.Source = "SELECT HtmlNavUrl, NavName, OutUrl, ViewFlag FROM keyicms_Navigation WHERE ParentID = " + Replace(ShowVavChild__MMColParam, "'", "''") + " ORDER BY Sequence ASC"
						ShowVavChild.CursorType = 0
						ShowVavChild.CursorLocation = 2
						ShowVavChild.LockType = 1
						ShowVavChild.Open()
						
						ShowVavChild_numRows = 0
						%>

                        <% While ( NOT ShowVavChild.EOF ) 
							if (ShowVavChild.Fields.Item("ViewFlag").Value) = -1 then %> 
                    		<li><a href=<%=(ShowVavChild.Fields.Item("HtmlNavUrl").Value)%><%=(ShowVavChild.Fields.Item("OutUrl").Value)%>><%=(ShowVavChild.Fields.Item("NavName").Value)%></a></li>
                 		<%	end if 
						ShowVavChild.MoveNext()
						Wend %>
						
                    </ul>
					</li>				
				<% ShowNavFa.MoveNext() 
				next %>

				<script type="text/javascript" language="javascript">
                var nav = document.getElementById("ulBox");
                var links = nav.getElementsByTagName("li");
                var lilen = nav.getElementsByTagName("a"); //判断地址
                var currenturl = document.location.href;
                var last = 0;
                for (var i = 0;i < links.length;i++)
                {
                	var linkurl =  lilen[i].getAttribute("href");
					if(currenturl.indexOf(linkurl) != -1) {last = i;}
                }
				last = 0;	//首页导航高亮
				links[last].className = "current";  //高亮代码样式
				if($(links[last]).length>0) $(links[last]).parent().parent("li").addClass("current")
                </script>
            </ul>
            <div class="scrollObj"></div>
        </div>
    </div>
    
    <div class="slide-bg banner">
        <div class="slide-wp">
          <div id="slides" class="slides">
            <% While ((Repeat6__numRows <> 0) AND (NOT ShowSlide.EOF)) %>
              <div><a class="a-jd opa js-log-login" href="<%=(ShowSlide.Fields.Item("Url").Value)%>" title="<%=(ShowSlide.Fields.Item("SlideName").Value)%>"><img class="slideImg" src="<%=(ShowSlide.Fields.Item("Pic").Value)%>" alt="<%=(ShowSlide.Fields.Item("SlideName").Value)%>"></a></div>
              <% 
				  Repeat6__index=Repeat6__index+1
				  Repeat6__numRows=Repeat6__numRows-1
				  ShowSlide.MoveNext()
				Wend
				%>
</div>
        </div>
    </div>

</div>

<div class="main_div2">
	<div class="main">
        <div class="box disciplines">
            <div class="bar">
            	<div class="tit tit01">
                    <h5><a href="/ShowNewsList.asp?SortID=1&PageNum=1">小组动态</a></h5>
                    <span>Events</span>
                </div>
            </div>
            <div class="cont">
            	<ul>
                  <% While ((Repeat1__numRows <> 0) AND (NOT NewsDataShow.EOF))
				  		if  (NewsDataShow.Fields.Item("ViewFlag").Value) = -1 then %>
                    		<li><a href="/ShowDetail.asp?ID=<%=(NewsDataShow.Fields.Item("ID").Value)%>" title="<%=(NewsDataShow.Fields.Item("NewsName").Value)%>" target="_blank"><%=(NewsDataShow.Fields.Item("NewsName").Value)%></a><span><%=(NewsDataShow.Fields.Item("AddTime").Value)%></span></li>
                  <% end if  
				  	Repeat1__index=Repeat1__index+1
					Repeat1__numRows=Repeat1__numRows-1
					NewsDataShow.MoveNext()
					Wend  %>
                
                </ul>
            </div>
        </div>
        
        <div class="box notice">
            <div class="bar">
            	<div class="tit tit01">
                    <h5><a href="/ShowNewsList.asp?SortID=2&PageNum=1">通知公告</a></h5>
                    <span>Notice</span>
                </div>
            </div>
            <div class="cont">
            	<ul>
				                    
                  <% While ((Repeat2__numRows <> 0) AND (NOT NoticeShow.EOF))
				  		if (NoticeShow.Fields.Item("ViewFlag").Value) = -1 then %>
                   			<li><a href="/ShowDetail.asp?ID=<%=(NoticeShow.Fields.Item("ID").Value)%>" title="<%=(NoticeShow.Fields.Item("NewsName").Value)%>" target="_blank"><%=(NoticeShow.Fields.Item("NewsName").Value)%></a><span><%=(NoticeShow.Fields.Item("AddTime").Value)%></span></li>
                    <% 	end if 
					  Repeat2__index=Repeat2__index+1
					  Repeat2__numRows=Repeat2__numRows-1
					  NoticeShow.MoveNext()
					Wend
					%>
				</ul>
            </div>
        </div>

    	<div class="clear"></div>
            
        <div class="box download">
            <div class="bar">
            	<div class="tit tit01">
                    <h5><a href="/ShowDownloadList.asp?SortID=1&PageNum=1">资料下载</a></h5>
                    <span>Download</span>
                </div>
            </div>
            <div class="cont">
                <ul>
                  <% While ((Repeat4__numRows <> 0) AND (NOT DownloadShow.EOF))
				  		if (DownloadShow.Fields.Item("ViewFlag").Value) = -1 then %>
                  <li><a href="/ShowDownloadDetail.asp?ID=<%=(DownloadShow.Fields.Item("ID").Value)%>" title=<%=(DownloadShow.Fields.Item("DownName").Value)%> target="_blank"><%=(DownloadShow.Fields.Item("DownName").Value)%></a><span><%=(DownloadShow.Fields.Item("AddTime").Value)%></span></li>
                    <% 	end if
					  Repeat4__index=Repeat4__index+1
					  Repeat4__numRows=Repeat4__numRows-1
					  DownloadShow.MoveNext()
					Wend
					%>
				</ul>
			</div>
        </div>
        
        <div class="box student">
            <div class="bar">
            	<div class="tit tit01">
                    <h5><a href="/ShowNewsList.asp?SortID=3&PageNum=1">网安新闻</a></h5>
                    <span>News</span>
                </div>
            </div>
            <div class="cont">
            	<ul>
                  <% While ((Repeat3__numRows <> 0) AND (NOT StudentShow.EOF)) 
						if (StudentShow.Fields.Item("ViewFlag").Value) = -1 then %>
                    		<li><a href="ShowDetail.asp?ID=<%=(StudentShow.Fields.Item("ID").Value)%>" title="<%=(StudentShow.Fields.Item("NewsName").Value)%>" target="_blank"><%=(StudentShow.Fields.Item("NewsName").Value)%></a><span><%=(StudentShow.Fields.Item("AddTime").Value)%></span></li>
                    <% 	end if 
					  Repeat3__index=Repeat3__index+1
					  Repeat3__numRows=Repeat3__numRows-1
					  StudentShow.MoveNext()
					Wend
					%>
				</ul>
            </div>
        </div>
        
        <div class="box explorer">
            <div class="bar">
            	<div class="tit tit01">
                    <h5>资源扩展</h5>
                    <span>Explorer</span>
                </div>
            </div>
            <div class="cont">
            	
                <p><a href="/Explorer/vedio.html" target="_blank"> <img src="/uploadfile/image/20150801/20150801153797629762.ico" title="精品课程视频" /> </a></p>
                <p><a href="/Explorer/book.html" target="_blank"> <img src="/uploadfile/image/20150801/20150801153797629762.ico" title="专业书籍" /> </a></p>
                <p>
                <form enctype="application/x-www-form-urlencoded" name="formJump" target="_blank" id="formJump">
                  <select style="WIDTH: 256px" onChange="MM_jumpMenu('parent',this,0)" name="menu1"> 
					  <% 
						While ((Repeat10__numRows <> 0) AND (NOT ShowExp.EOF)) 
						if (ShowExp.Fields.Item("SortID").Value) = 1 and (ShowExp.Fields.Item("ViewFlag").Value) = -1 then
						%>
                      		<option value="<%=(ShowExp.Fields.Item("Content").Value)%>"><%=(ShowExp.Fields.Item("CaseName").Value)%></option> 
						<% 
						  end if
						  Repeat10__index=Repeat10__index+1
						  Repeat10__numRows=Repeat10__numRows-1
						  ShowExp.MoveNext()
						Wend
						%>
                  </select>
                </form>
                </p> 
				<p>           
                <form enctype="application/x-www-form-urlencoded" name="formJump" target="_blank" id="formJump">
                  <select style="WIDTH: 256px" onChange="MM_jumpMenu('parent',this,0)" name="menu1"> 
					  <% 
					    ShowExp.MoveFirst()
						While ((Repeat10__numRows <> 0) AND (NOT ShowExp.EOF)) 
						if (ShowExp.Fields.Item("SortID").Value) = 2 and (ShowExp.Fields.Item("ViewFlag").Value) = -1 then
						%>
                      		<option value="<%=(ShowExp.Fields.Item("Content").Value)%>"><%=(ShowExp.Fields.Item("CaseName").Value)%></option> 
						<% 
						  end if
						  Repeat10__index=Repeat10__index+1
						  Repeat10__numRows=Repeat10__numRows-1
						  ShowExp.MoveNext()
						Wend
						%>
                  </select>
                </form>
                </p>
                <p>
                <form enctype="application/x-www-form-urlencoded" name="formJump" target="_blank" id="formJump">
                  <select style="WIDTH: 256px" onChange="MM_jumpMenu('parent',this,0)" name="menu1"> 
					  <% 
					    ShowExp.MoveFirst()
						While ((Repeat10__numRows <> 0) AND (NOT ShowExp.EOF)) 
						if (ShowExp.Fields.Item("SortID").Value) = 3 and (ShowExp.Fields.Item("ViewFlag").Value) = -1 then
						%>
                      		<option value="<%=(ShowExp.Fields.Item("Content").Value)%>"><%=(ShowExp.Fields.Item("CaseName").Value)%></option> 
						<% 
						  end if
						  Repeat10__index=Repeat10__index+1
						  Repeat10__numRows=Repeat10__numRows-1
						  ShowExp.MoveNext()
						Wend
						%>
                  </select>
                </form>
				</p>
            </div>
        </div>
    </div>
	
    <div class="chanpin">
        <div class="cbar">
            <img src="/Template/PC/Default/image/ProTitle.png" />
            <a href="/ShowPicture.asp">more...</a>
        </div>
        <div class="Cp">
            <div class="Button Prev"></div>
            <div class="Button Next"></div>
            <div class="CpBox">
            	<ul>
                  <% While ((Repeat5__numRows <> 0) AND (NOT ShowPicture.EOF))
				  		if (ShowPicture.Fields.Item("ViewFlag").Value) = -1 then %>
						  <li>
							  <a href="/ShowPicture.asp?ID=<%=(ShowPicture.Fields.Item("ID").Value)%>" title="<%=(ShowPicture.Fields.Item("ProductName").Value)%>"><img src="<%=(ShowPicture.Fields.Item("SmallPic").Value)%>" alt="<%=(ShowPicture.Fields.Item("ProductName").Value)%>"/></a>
							  <p class="txt"><%=(ShowPicture.Fields.Item("ProductName").Value)%></p>
						  </li>
                    <% 	end if
					  Repeat5__index=Repeat5__index+1
					  Repeat5__numRows=Repeat5__numRows-1
					  ShowPicture.MoveNext()
					Wend
					%>
				</ul>
            </div>
        </div>
    </div>
	
    <div class="link">
        <div class="bar">
            <span>友情链接</span>
        </div>
        <div class="cont img">
          <% 
			While ((Repeat8__numRows <> 0) AND (NOT ShowFriendLink.EOF)) 
				if (ShowFriendLink.Fields.Item("LinkType").Value) = -1 and (ShowFriendLink.Fields.Item("ViewFlag").Value) = -1 then
			%>
            <a href='<%=(ShowFriendLink.Fields.Item("LinkUrl").Value)%>' title="<%=(ShowFriendLink.Fields.Item("LinkName").Value)%>" target='_blank'><img src='<%=(ShowFriendLink.Fields.Item("LinkFace").Value)%>' /></a>
            <% 
				end if
			  Repeat8__index=Repeat8__index+1
			  Repeat8__numRows=Repeat8__numRows-1
			  ShowFriendLink.MoveNext()
			Wend
			%>
</div>
        <div class="cont txt">
          <% 
		  	ShowFriendLink.MoveFirst()	'将数据库指针移到第一条
			While ((Repeat9__numRows <> 0) AND (NOT ShowFriendLink.EOF)) 
				if (ShowFriendLink.Fields.Item("LinkType").Value) = 0 and (ShowFriendLink.Fields.Item("ViewFlag").Value) = -1 then
			%>
          			<a href='<%=(ShowFriendLink.Fields.Item("LinkUrl").Value)%>' title="<%=(ShowFriendLink.Fields.Item("LinkName").Value)%>" target='_blank'><%=(ShowFriendLink.Fields.Item("LinkName").Value)%></a>
            <% 
				end if
			  Repeat9__index=Repeat9__index+1
			  Repeat9__numRows=Repeat9__numRows-1
			  ShowFriendLink.MoveNext()
			Wend
			%>
</div>
    </div>
</div>

<div class="footer">
	<div class="foot">
    	<img src="<%=(ShowSite.Fields.Item("Ico").Value)%>" height="125" />
        
        <p>Copyright © <%=(ShowSite.Fields.Item("ComName").Value)%> All rights reserved</p>
        <p>网站备案号：<%=(ShowSite.Fields.Item("IcpNumber").Value)%></p>
        <p>联系方式：<%=(ShowSite.Fields.Item("Mobile").Value)%></p>
		<p>地址：<%=(ShowSite.Fields.Item("Address").Value)%></p>
        <p>Powered by <a href="<%=(ShowSite.Fields.Item("SiteUrl").Value)%>" target="_blank"><%=(ShowSite.Fields.Item("Contacts").Value)%></a> Inc.</p>
    </div>
</div>

<!--在线客服代码-->
<link href="/Template/PC/Default/File/Online/online.css" rel="stylesheet" type="text/css" />
<script src="/Template/PC/Default/File/Online/online.js"></script>

<div id="online_qq_layer">
	<div id="online_qq_tab">
		<div class="online_icon">
			<a title id="floatShow" style=" display:block;" href="javascript:void(0);">&nbsp;</a>
			<a title id="floatHide" style=" display:none;" href="javascript:void(0);">&nbsp;</a>
		</div>
	</div>
	<div id="onlineService">
		<div class="online_windows overz">
			<div class="online_w_top"></div>
			<div class="online_w_c overz">
                <img src="<%=(ShowSite.Fields.Item("Ico").Value)%>" />
                <% 
				While ((Repeat7__numRows <> 0) AND (NOT ShowContact.EOF)) 
				%>
                  <div class="online_bar collapse">
                    <h2><a href="JavaScript:void(0)"><%=(ShowContact.Fields.Item("User").Value)%></a></h2>
                    <div class="online_content overz">
                      <ul class="overz">
                        <li><a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=<%=(ShowContact.Fields.Item("QQ").Value)%>&site=qq&menu=yes" class="qq_icon">QQ交谈</a></li>
                        <li>电话：<%=(ShowContact.Fields.Item("Phone").Value)%></li>
                        <li>微信号：<%=(ShowContact.Fields.Item("WeChat").Value)%></li>
				      </ul>
			        </div>
                  </div>
                  <% 
				  Repeat7__index=Repeat7__index+1
				  Repeat7__numRows=Repeat7__numRows-1
				  ShowContact.MoveNext()
				Wend
				%>
		  </div>
            <div class="online_w_bottom"></div>
        </div>
    </div>
</div>
<!--在线客服代码END-->

<link href="/Template/PC/Default/File/360/css.css" rel="stylesheet" type="text/css" />
<script src="/Template/PC/Default/File/360/jquery.slides.min.js"></script> 
<script>
$(function() {
	$('#slides').slidesjs({
		play:{
			active: false,
			effect: "fade",
			auto: true,
			interval: 4000
		},
		effect: {
			fade: {
			speed: 1500,
			crossfade: true
			}
		},
		pagination: {
			active: true
		},
		navigation:{
			active: false
		}
	});
});
</script>
<script>
$(document).ready(function(e) {
	www = 210
	$('.Next').click(function(){
		obj = $(this).siblings('.CpBox').find('ul')
		linum = obj.find('li').length;
		w = linum * www;
		if($(obj).is(':animated')){
			$(obj).stop(true,true);
		}
		if(linum > 5){
			Prol = parseInt($(obj).css('left'));
			
			if(Prol != (w - (www * 5)) * -1){
				$(obj).animate({left: Prol - www + 'px'},500);
			}
			else{//交换图片显示时
				$(obj).animate({left: 0},500)
			}
		}
	})

	$('.Prev').click(function(){
		obj = $(this).siblings('.CpBox').find('ul')
		linum = obj.find('li').length;
		w = linum * www;
		if($(obj).is(':animated')){
			$(obj).stop(true,true);
		}
		if(linum > 5){
			Prol = parseInt($(obj).css('left'));
			
			if(Prol == 0){
				$(obj).animate({left: (w - (www * 5)) * -1},500);
			}
			else{
				$(obj).animate({left: Prol + www},500);				
			}
		}
	})    

	$(".CpBox li, .case ul li").hover(
		function(){
			$(this).children().stop(true,false);
			$(this).children(".txt").animate({top:-35},250);	 
		}, 
		function() {
			$(this).children(".txt").animate({top:0},250);	 
		}
	);
});
</script>
</body>
</html>
<%
NewsDataShow.Close()
Set NewsDataShow = Nothing

NoticeShow.Close()
Set NoticeShow = Nothing

StudentShow.Close()
Set StudentShow = Nothing

DownloadShow.Close()
Set DownloadShow = Nothing

ShowPicture.Close()
Set ShowPicture = Nothing

ShowSlide.Close()
Set ShowSlide = Nothing

ShowNavFa.Close()
Set ShowNavFa = Nothing

ShowVavChild.Close()
Set ShowVavChild = Nothing

ShowExp.Close()
Set ShowExp = Nothing

ShowFriendLink.Close()
Set ShowFriendLink = Nothing

ShowSite.Close()
Set ShowSite = Nothing

ShowContact.Close()
Set ShowContact = Nothing
%>
