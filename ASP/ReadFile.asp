<html>
<head> 
        <link rel="stylesheet" type="text/css" href="\Scripts\Smart_TV\CSS\ReadFile.css" />   
</head>
<!--This file reads text files and turns text links into clickable links-->
<body>
<% 

Const Filename = "shadowLine.txt"    ' file to read
Const ForReading = 1, ForWriting = 2, ForAppending = 3
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

' Create a filesystem object
Dim FSO
set FSO = server.createObject("Scripting.FileSystemObject")

' Map the logical path to the physical system path
Dim Filepath
Filepath = Server.MapPath("\Scripts\Smart_TV\Data\" & Filename)
 Response.Write "<h3><i><font color=blue> File " & Filepath & " exist</font></i></h3>"
 response.write Filepath
if FSO.FileExists(Filepath) Then

    ' Get a handle to the file
    Dim file    
    set file = FSO.GetFile(Filepath)
    ' Get some info about the file
    Dim FileSize
    FileSize = file.Size

    Response.Write "<p><b>File: " & Filename & " (size " & FileSize  &_
                   " bytes)</b></p><hr>"
    Response.Write "<pre>"
    ' Open the file using a textstream
    Dim TextStream
    Set TextStream = file.OpenAsTextStream(ForReading, _
                                               TristateUseDefault)

    ' Read the file line by line
    Do While Not TextStream.AtEndOfStream
        Dim Line
        Line = TextStream.readline
        ' Do something with "Line"
        Line = Line & vbCRLF
        dim strText
        strText=Line
       '  Call our function to see if this is a link
       call create_links(strText)
       response.write strText 
    Loop

    Response.Write "</pre><hr>"

    Set TextStream = nothing
    
Else

   Response.Write "<h3><i><font color=red> File " & Filename &" does not exist</font></i></h3>"

End If


function create_links(strText)
    strText = " " & strText
    strText = ereg_replace(strText, "(^|[\n ])([\w]+?://[^ ,""\s<]*)", "$1<a href=""$2"" ref=""nofollow"">$2</a>")
    strText = ereg_replace(strText, "(^|[\n ])((www|ftp)\.[^ ,""\s<]*)", "$1<a target=""_blank"" ref=""nofollow"" href=""http://$2"">$2</a>")
    strText = ereg_replace(strText, "(^|[\n ])([a-z0-9&\-_.]+?)@([\w\-]+\.([\w\-\.]+\.)*[\w]+)", "$1<a href=""mailto:$2@$3"">$2@$3</a>")
    strText = right(strText, len(strText)-1)
    strText = Replace(strText,"." & chr(34) & ">",chr(34) & ">")
    strText = Replace(strText,".)" & chr(34) & ">",chr(34) & ">")
    create_links = strText
end function

function ereg_replace(strOriginalString, strPattern, strReplacement)
    dim objRegExp : set objRegExp = new RegExp
    objRegExp.Pattern = strPattern
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    ereg_replace = objRegExp.replace(strOriginalString, strReplacement)
    set objRegExp = nothing
end function


response.write(Server.MapPath("ReadFile.asp") & "<br>")
response.write(Server.MapPath("script/ReadFile.asp") & "<br>")
response.write(Server.MapPath("/script/ReadFile.asp") & "<br>")
response.write(Server.MapPath("\script") & "<br>")
response.write(Server.MapPath("/") & "  using this backslash /<br>")

response.write(Server.MapPath("\") & "  using this backslash \<br>")

Set FSO = nothing
%>
</body>
</html>