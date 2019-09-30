<%
function fnFormatDosPath(sUnformattedPath)
  LW("1 sUnformattedPath=" & sUnformattedPath)

  ' remove leading '\' if user defined it
  if Left(sUnformattedPath,1) = "\" then
    sUnformattedPath = Mid(sUnformattedPath,2)
  end if
  LW("2 sUnformattedPath=" & sUnformattedPath)

  ' add tailing '\' to the current folder
  if Right(sUnformattedPath,1) <> "\" then
    sUnformattedPath = sUnformattedPath & "\"
    if sUnformattedPath = "\" then ' handle the exception for an empty folder
      sUnformattedPath = ""
    end if
  end if
  LW("3 sUnformattedPath=" & sUnformattedPath)

  fnFormatDosPath = sUnformattedPath
end function
%>