<%
server.scripttimeout = 999
Call SiteInfo()

Function ReadFromUTF(FilePath,CharSet) 'FilePath为要读取的模版文件路径；Charset是编码
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	if fso.FileExists(server.MapPath(FilePath)) then
		Dim str
		Set stm = server.CreateObject("adodb.stream")
		stm.Type = 2
		stm.Mode = 3
		stm.Charset = CharSet
		stm.Open
		stm.loadfromfile Server.MapPath(FilePath)
		str = stm.readtext
		stm.Close
		Set stm = Nothing
		ReadFromUTF = str
	Else
		ReadFromUTF = ""
	end if
End Function

Function WriteToUTF(FilePath, Content)
Set objStream=Server.CreateObject("ADODB.Stream")
With objStream
	.Open
	.Charset="utf-8"
	.Position=objStream.Size
	.WriteText=Content
	.SaveToFile server.mappath(FilePath),2
	.Close
End With
Set objStream=Nothing
End Function
%>