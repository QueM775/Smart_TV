<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>My video player</title>

  <!-- CSS -->
  <link rel="stylesheet" type="text/css" href="CSS/jquery.fancybox.min.css">
  <link rel="stylesheet" type="text/css" href="CSS/movie.css">
  <link rel='stylesheet' type='text/css' href='CSS/music.css'>
</head>
<body>
<!--
Include HERE all necessary *.ASP files from the ASP\ folder 
Included files must have opening and closing  ASP brackets i.e. "< %" and "% >"
File includes in HTML code (before ASP code starts in the DEFAULT.ASP ), it is not mandatory
but we have to accept establish a rule how to do it.
-->
<!-- #include file="ASP\fnMovieBox.asp" -->
<!-- #include file="ASP\fnAudioPlayer.asp" -->
<!-- #include file="ASP\fnPdfBox.asp" -->
<!-- #include file="ASP\fnSubfolder.asp" -->
<!-- #include file="ASP\fnLW.asp" -->
<!-- #include file="ASP\fnUpperFolder.asp" -->
<!-- #include file="ASP\fnFormatDosPath.asp" -->
<%

' Root folder for all files MUST be (current-folder-where-DEFUALT.ASP-file-is)+('\Data')
' Suffix **DOS indicates that this path is in DOS notation (as IIS sees it aaa\bbb\ccc\ddd\)  
' in opposite to WEB notation /aaa/bbb/ccc/ddd/
sRootFolderDOS = Server.MapPath(".") & "\Data\"
LW("0 sRootFolderDOS=" & sRootFolderDOS)

' Current filer counts from the sRootFolderDOS
sCurrentFldrDOS = Request.QueryString("fld")
sCurrentFldrDOS = fnFormatDosPath(sCurrentFldrDOS)

' add tailing '\' to the root folder
if Right(sRootFolderDOS, 1) <> "\" then
  sRootFolderDOS = sRootFolderDOS & "\"
end if 

set fs=Server.CreateObject("Scripting.FileSystemObject")
LW("Check folder >>>  " & sRootFolderDOS & sCurrentFldrDOS)
if not (fs.FolderExists(sRootFolderDOS & sCurrentFldrDOS)) then
  LW("ERROR: current folder '" & sCurrentFldrDOS & "' does not exist.")
  LW("Reset current folder to NULL")
  sCurrentFldrDOS = ""
end if

LW("sRootFolderDOS   =" & sRootFolderDOS)
LW("sCurrentFldrDOS  =" & sCurrentFldrDOS)
'
' Create link to upper folder
'
LW("Len(sCurrentFldrDOS)=" & Len(sCurrentFldrDOS))
if Len(sCurrentFldrDOS) > 0 then
  ' we are not in the root folder, so we can go up
  sFolderNameUpper = Left(sCurrentFldrDOS,Len(sCurrentFldrDOS)-1)
  LW("sFolderNameUpper=" & sFolderNameUpper)
  LW("InStr(sFolderNameUpper,'\')=" & InStr(sFolderNameUpper,"\"))
  iLastSlashPosition = (InStrRev(sFolderNameUpper, "\"))
  if (iLastSlashPosition > 0) then
    ' we have more then one subfolders level i.e. aaa\bbb\ccc\ddd
    LW("iLastSlashPosition=" & iLastSlashPosition)
    sFolderNameUpper = Left(sFolderNameUpper, iLastSlashPosition)
    LW("sFolderNameUpper=" & sFolderNameUpper & " *********************")
  else
    sFolderNameUpper = "\"
  end if
  Response.write(fnUpperFolder(sFolderNameUpper))
end if 'if Len(sCurrentFldrDOS) > 0

'=========================================
' List of subFOLDERS in current folder
'=========================================
set objFolderCurrent=fs.GetFolder(sRootFolderDOS & sCurrentFldrDOS)
LW("objFolderCurrent.path=" & sRootFolderDOS & sCurrentFldrDOS)

LW("---- FOLDERS ----")
for each x in objFolderCurrent.SubFolders
  'Print the name of all subfolders in the test folder
  LW("x.Name = " & x.Name & " ===========================================")
  sHTML = fnSubfolder(sCurrentFldrDOS, x.Name)
  Response.write(sHTML)
  LW("sHTML = " & sHTML)
next
LW("---- FOLDERS ----")

'=====================================
' List of all *.MP4 FILES in current folder
'=====================================
LW("==== MOVIE FILES ====")
Response.write("<div id='idMovieContainer' class='clsAllMovies'>" & vbCRLF)
sHTML_code = fnMovieBox(sRootFolderDOS, sCurrentFldrDOS)
Response.write(sHTML_code)
Response.write("</div>" & vbCRLF)
LW("==== MOVIE FILES ====")

'=====================================
' List of all *.MP4 FILES in current folder
'=====================================
LW("==== AUDIO FILES ====")
Response.write(fnAudioPlayer(sRootFolderDOS, sCurrentFldrDOS))

LW("==== AUDIO FILES ====")


'=====================================
' List of all *.PDF FILES in current folder
'=====================================
LW("==== PDF FILES ====")
set objFolderCurrent=fs.GetFolder(sRootFolderDOS & sCurrentFldrDOS)
Response.write("<div id='idMovieContainer' class='clsAllMovies'>" & vbCRLF)
for each objFolderCurrent in objFolderCurrent.files
  'Print the name of all subfolders in the test folder
  if Ucase(Right(objFolderCurrent.Name, 4)) = ".PDF" then
    sHTML = fnPdfBox(sRootFolderDOS, sCurrentFldrDOS, objFolderCurrent.Name)
    LW("sHTML=" & sHTML)
    Response.write(sHTML)
  end if
next
Response.write("</div>" & vbCRLF)
LW("==== PDF FILES ====")


'============= CLOSING PAGE CODE ============='
set objFolderCurrent=nothing
set fs=nothing
%>
<script src="Javascripts/jquery-3.2.1.min.js"></script>
<script src="Javascripts/jquery.fancybox.min.js"></script>
<script src="Javascripts/AudioPlayerA.js"></script>
</body>
</html>
