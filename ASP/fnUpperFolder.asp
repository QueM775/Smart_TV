<%
function fnUpperFolder(sUpFolder)
  sHtmlTemplate = "<a href='?fld=@Fldr'><h1>Back to folder: @Fldr</h1></a></br>"
  fnUpperFolder = Replace(sHtmlTemplate, "@Fldr", sUpFolder)
end function
%>