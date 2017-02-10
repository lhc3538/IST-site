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
ShowSort.Source = "SELECT ID, SortName, ViewFlag FROM keyicms_DownSort"
ShowSort.CursorType = 0
ShowSort.CursorLocation = 2
ShowSort.LockType = 1
ShowSort.Open()

ShowSort_numRows = 0
%>
<%
Dim ShowList__MMColParam	'用于记录SortID
ShowList__MMColParam = "1"
If (Request.QueryString("SortID") <> "") Then 
  ShowList__MMColParam = Request.QueryString("SortID")
End If
%>
<%
Dim ShowDisciplines__MMColParam
ShowDisciplines__MMColParam = ShowList__MMColParam
If (Request("MM_EmptyValue") <> "") Then 
  ShowDisciplines__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim ShowDisciplines
Dim ShowDisciplines_numRows

Set ShowDisciplines = Server.CreateObject("ADODB.Recordset")
ShowDisciplines.ActiveConnection = MM_SiteData_STRING
ShowDisciplines.Source = "SELECT * FROM keyicms_Download WHERE SortID = " + Replace(ShowDisciplines__MMColParam, "'", "''") + " ORDER BY AddTime DESC"
ShowDisciplines.CursorType = 0
ShowDisciplines.CursorLocation = 2
ShowDisciplines.LockType = 1
ShowDisciplines.Open()

ShowDisciplines_numRows = 0
%>

<%
Dim ShowList
Dim ShowList_numRows

Set ShowList = Server.CreateObject("ADODB.Recordset")
ShowList.ActiveConnection = MM_SiteData_STRING
ShowList.Source = "SELECT AddTime, ID, DownName, SortID, ViewFlag FROM keyicms_Download WHERE SortID = " + Replace(ShowList__MMColParam, "'", "''") + " ORDER BY AddTime DESC"
ShowList.CursorType = 0
ShowList.CursorLocation = 3
ShowList.LockType = 1
ShowList.PageSize = 20	'每页20条
ShowList.Open()

ShowList_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 13
Repeat1__index = 0
ShowDisciplines_numRows = ShowDisciplines_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = -1
Repeat2__index = 0
ShowSort_numRows = ShowSort_numRows + Repeat2__numRows
%>
<%
Dim Repeat3__numRows
Dim Repeat3__index

Repeat3__numRows = 20
Repeat3__index = 0
ShowList_numRows = ShowList_numRows + Repeat3__numRows
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
				links[last].className = "current";  //高亮代码样式
				if($(links[last]).length>0) $(links[last]).parent().parent("li").addClass("current")
                </script>
            </ul>
            <div class="scrollObj"></div>
        </div>
    </div>
	<div class="main_div3">
		<div class="main">
			<div class="right">
				<div class="box">
					<div class="bar">
						<div class="tit tit03">下载分类</div>
					</div>
					<div class="cont">
						<ul class="sort">
                          <% Dim NextTitle
						   While ((Repeat2__numRows <> 0) AND (NOT ShowSort.EOF)) 
						   		if cstr(ShowSort.Fields.Item("ID").Value) = cstr(ShowList__MMColParam) then	'确定下一栏的标题
									NextTitle = (ShowSort.Fields.Item("SortName").Value)
								end if
								if (ShowSort.Fields.Item("ViewFlag").Value) = -1 then							%>
                          			<li><a href="/ShowDownloadList.asp?SortID=<%=(ShowSort.Fields.Item("ID").Value)%>&PageNum=1"><%=(ShowSort.Fields.Item("SortName").Value)%></a></li>
                            <% 	end if
							  Repeat2__index=Repeat2__index+1
							  Repeat2__numRows=Repeat2__numRows-1
							  ShowSort.MoveNext()
							Wend
							%>
</ul>
					</div>
				</div>
				<div class="box">
					<div class="bar">
						<div class="tit tit03">最新 <%=NextTitle%></div>
					</div>
					<div class="cont">
						<ul class="list">
                          <% While ((Repeat1__numRows <> 0) AND (NOT ShowDisciplines.EOF)) 
								if (ShowDisciplines.Fields.Item("ViewFlag").Value) = -1 then	%>
                          			<li><a href='/ShowDownloadDetail.asp?ID=<%=(ShowDisciplines.Fields.Item("ID").Value)%>' title='<%=(ShowDisciplines.Fields.Item("DownName").Value)%>'><%=(ShowDisciplines.Fields.Item("DownName").Value)%></a></li>
                            <% 	end if
							  Repeat1__index=Repeat1__index+1
							  Repeat1__numRows=Repeat1__numRows-1
							  ShowDisciplines.MoveNext()
							Wend
							%>
					  </ul>
					</div>
				</div>
			</div>
			
			<div class="left">
				<div class="box">
					<div class="bar">
						<div class="tit tit03">全部 <%=NextTitle%></div>
					</div>
					<div class="cont">
						<ul class='list'>
                          <% 
						  Dim  PageNum
						  PageNum = 1
						  if Request.QueryString("PageNum") <> "" then
						  	PageNum = Request.QueryString("PageNum")
						  end if

						  ShowList.AbsolutePage =  PageNum	'将记录指针移到第PageNum页的第一行
							While ((Repeat3__numRows <> 0) AND (NOT ShowList.EOF)) 
								if (ShowList.Fields.Item("ViewFlag").Value) = -1 then
							%>
                          			<li><a href="/ShowDownloadDetail.asp?ID=<%=(ShowList.Fields.Item("ID").Value)%>" title="<%=(ShowList.Fields.Item("DownName").Value)%>"><%=(ShowList.Fields.Item("DownName").Value)%></a> <span><%=(ShowList.Fields.Item("AddTime").Value)%></span></li>
                            <% 	end if
							  Repeat3__index=Repeat3__index+1
							  Repeat3__numRows=Repeat3__numRows-1
							  ShowList.MoveNext()
							Wend
							%>
							</ul>
						<div class="Page">
						<% if PageNum <> 1 then %>
							<a href='/ShowDownloadList.asp?SortID=<%=ShowList__MMColParam%>&PageNum=<%=PageNum-1%>'>上一页</a>
						<% end if %>    
							
							<a href="" Class="selected"><%=PageNum%></a>
							
						<% if cstr(PageNum) <> cstr(ShowList.PageCount) then %>
							<a href='/ShowDownloadList.asp?SortID=<%=ShowList__MMColParam%>&PageNum=<%=PageNum+1%>'>下一页</a>
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
ShowDisciplines.Close()
Set ShowDisciplines = Nothing
%>
<%
ShowList.Close()
Set ShowList = Nothing
%>
<%
ShowSort.Close()
Set ShowSort = Nothing
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