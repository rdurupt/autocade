<%
Set fs = CreateObject("Scripting.FileSystemObject")
dapath = "E:\webprod\wwwflywayonline\cas-aviation\"
Set f = fs.GetFolder(dapath)
'Set fc = f.Files
Set fc = f.SubFolders 
For Each f1 in fc
	response.write (f1.name & "<br>" )
Next

%>
