<%
sRelPath = Replace(Request.QueryString("ajax_path"),"/", "\")
sRootPath = Server.MapPath(".")
sRootPath = Mid(sRootPath, 1, InStrRev(sRootPath, "\") )
sFullPath = sRootPath & "Data\" & sRelPath

' Read WHOLE template HTML file into memory
sTemplateFileName=sFullPath & "BookMark.txt"
set fso=Server.CreateObject("Scripting.FileSystemObject")
set objTemplateFile=fso.OpenTextFile(sTemplateFileName,1,false)
sFileContent=objTemplateFile.ReadAll
objTemplateFile.close
Response.Write(sFileContent)
'LW ("EQ sTemplateFileName=" & sTemplateFileName)
%>

