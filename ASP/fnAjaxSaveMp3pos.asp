<%
response.write("from AJAX ASP!")
sRelPath = Replace(Request.QueryString("ajax_path"),"/", "\")
sRootPath = Server.MapPath(".")
sRootPath = Mid(sRootPath, 1, InStrRev(sRootPath, "\") )
sFullPath = sRootPath & sRelPath

Call fnWriteTextFile (sFullPath, Request.QueryString("ajax_name") & "|" & Request.QueryString("ajax_pos"))

'======================================================================
function fnWriteTextFile (sFullPath, sMessage)
  Dim fs,f
  Const Appending=8
  Const Writing=2

  set fs=Server.CreateObject("Scripting.FileSystemObject")
  set f=fs.OpenTextFile(sFullPath & "BookMark.txt", Writing, true)
  f.WriteLine(sMessage)
  f.Close
  set f=Nothing
  set fs=Nothing

end function

%>

