<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/SiteData.asp" -->

<%
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
%>

<%
Dim ShowSort
Dim ShowSort_numRows

Set ShowSort = Server.CreateObject("ADODB.Recordset")
ShowSort.ActiveConnection = MM_SiteData_STRING
ShowSort.Source = "SELECT ID, SortName, ViewFlag FROM keyicms_NewsSort"
ShowSort.CursorType = 0
ShowSort.CursorLocation = 2
ShowSort.LockType = 1
ShowSort.Open()

ShowSort_numRows = 0
%>
<%
Dim ShowDetail__MMColParam
ShowDetail__MMColParam = "1"
If (Request.QueryString("ID") <> "") Then 
  ShowDetail__MMColParam = Request.QueryString("ID")
End If
%>
<%
Dim ShowDetail
Dim ShowDetail_numRows

Set ShowDetail = Server.CreateObject("ADODB.Recordset")
ShowDetail.ActiveConnection = MM_SiteData_STRING
ShowDetail.Source = "SELECT AboutName, AddTime, Content, ViewFlag FROM keyicms_About WHERE ID = " + Replace(ShowDetail__MMColParam, "'", "''") + ""
ShowDetail.CursorType = 0
ShowDetail.CursorLocation = 2
ShowDetail.LockType = 1
ShowDetail.Open()

ShowDetail_numRows = 0
%>
<%
Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = -1
Repeat2__index = 0
ShowSort_numRows = ShowSort_numRows + Repeat2__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=(ShowSite.Fields.Item("SiteTitle").Value)%></title>
<meta name="keywords" content="<%=(ShowSite.Fields.Item("Keywords").Value)%>" />
<meta name="description" content="<%=(ShowSite.Fields.Item("Descriptions").Value)%>" />
<link href="/Template/PC/Default/css/css.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="/Template/PC/Default/js/jquery-1-7-2.js"></script>
<script type="text/javascript" src="/Template/PC/Default/js/navScroll.js"></script>
<script type="text/javascript">
$(function(){
	$(".nav > ul > li").ScrollNav({
		scrollObj:".nav > .scrollObj",
		speed:"fast"
	});
});
</script>
<script src="/Template/PC/Default/js/jquery.kinMaxShow-1.1.min.js" type="text/javascript" charset="utf-8"></script>
<script type="text/javascript">
$(function(){
	$("#kinMaxShow").kinMaxShow();
});
</script>
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
            <form name="searchForm" method="post" onsubmit="return CheckForm()" action="JavaScript:void(0)" class="search">
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
				last =1;
				links[last].className = "current";  //高亮代码样式
				if($(links[last]).length>0) $(links[last]).parent().parent("li").addClass("current")
                </script>
            </ul>
            <div class="scrollObj"></div>
        </div>
    </div>
	<div class="main_div3">
		<div class="main">
			<div class="left">
				<div class="box">
					<div class="bar">
						<div class="tit tit03">详细内容</div>
					</div>
					<div class="cont">
						<div class="article">
							<% if (ShowDetail.Fields.Item("ViewFlag").Value) = -1 then %>
							<div class='artName'><%=(ShowDetail.Fields.Item("AboutName").Value)%></div>
							<div class='artInfo'>
								添加时间：<%=(ShowDetail.Fields.Item("AddTime").Value)%>&nbsp;&nbsp;
							</div>
							<div class='artCont'><%=(ShowDetail.Fields.Item("Content").Value)%></div>
							<% else %>
							<div class='artName'></div>
							<div class='artInfo'>
								&nbsp;&nbsp;
							</div>
							<div class='artCont'>内容失效</div>
							<% end if %>
						</div>
					</div>
				</div>
			</div>
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

</body>
</html>
<%
ShowSort.Close()
Set ShowSort = Nothing
%>
<%
ShowDetail.Close()
Set ShowDetail = Nothing
%>
<%
ShowSite.Close()
Set ShowSite = Nothing
%>
<%
ShowNavFa.Close()
Set ShowNavFa = Nothing
%>
<%
ShowVavChild.Close()
Set ShowVavChild = Nothing
%>