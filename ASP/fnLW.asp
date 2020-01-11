<%
'======================================================================
' This function (L)og (W)riter writes data into log file on server side 
'======================================================================
function LW(sMessage)
  Dim fs,f
  Const Appending=8

  set fs=Server.CreateObject("Scripting.FileSystemObject")
  set f=fs.OpenTextFile(Server.MapPath("===LogFile===.log"), Appending, true)
  f.WriteLine(fnCUrrentDateTime() & " " & fnClientIP() & " " & sMessage)
  f.Close
  set f=Nothing
  set fs=Nothing

end function

Function fnClientIP()
    Dim strIP : strIP = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
    If strIP = "" Then strIP = Request.ServerVariables("REMOTE_ADDR")
    fnClientIP = strIP
End Function


Function fnCUrrentDateTime()
  Dim dtCurrDate
  dtCurrDate = Now()
  fnCUrrentDateTime = Year(dtCurrDate) & "-" & _
                      Right("0" & Month(dtCurrDate),2) & "-" & _
                      Right("0" & Day(dtCurrDate),2) & " "  & _
                      Right("0" & Hour(dtCurrDate),2) & ":" & _
                      Right("0" & Minute(dtCurrDate),2) & ":" & _
                      Right("0" & Second(dtCurrDate),2)
End Function
%>