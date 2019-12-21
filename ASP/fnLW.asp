<%
'======================================================================
' This function (L)og (W)riter writes data into log file on server side 
'======================================================================
function LW(sMessage)
  Dim fs,f
  Dim dtCurrDate
  Const Appending=8
  dtCurrDate = Now()
  strCurrentDateTime = Hour(dtCurrDate) & ":" & Minute(dtCurrDate) & ":" & Second(dtCurrDate)

  set fs=Server.CreateObject("Scripting.FileSystemObject")
  set f=fs.OpenTextFile(Server.MapPath("===LogFile===.log"), Appending, true)
  f.WriteLine(strCurrentDateTime & " " & sMessage)
  f.Close
  set f=Nothing
  set fs=Nothing

end function
%>