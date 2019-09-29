<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>My video player</title>

  <!-- CSS -->
  <link rel="stylesheet" type="text/css" href="CSS/jquery.fancybox.min.css">
  <link rel="stylesheet" type="text/css" href="CSS/movie.css">
</head>
<body>
<%
sFolderNameRoot = Server.MapPath(".")
sFolderNameCurrent = Request.QueryString("fld")
' remove leading '\' if user defined it
if Left(sFolderNameCurrent,1) = "\" then
  sFolderNameCurrent = Mid(sFolderNameCurrent,2)
end if
' add tailing '\' to the current folder
if Right(sFolderNameCurrent,1) <> "\" then
  sFolderNameCurrent = sFolderNameCurrent & "\"
  if sFolderNameCurrent = "\" then
    sFolderNameCurrent = ""
  end if
end if
' add tailing '\' to the root folder
if Right(sFolderNameRoot, 1) <> "\" then
  sFolderNameRoot = sFolderNameRoot & "\"
end if 

set fs=Server.CreateObject("Scripting.FileSystemObject")
if not (fs.FolderExists(sFolderNameRoot & sFolderNameCurrent)) then
  LW("ERROR: current folder '" & sFolderNameCurrent & "' does not exist.")
  LW("Reset current folder to NULL")
  'sFolderNameCurrent = "\"
  sFolderNameCurrent = ""
end if

LW("sFolderNameRoot   =" & sFolderNameRoot)
LW("sFolderNameCurrent=" & sFolderNameCurrent)
'
' Create link to upper folder
'
LW("Len(sFolderNameCurrent)=" & Len(sFolderNameCurrent))
if Len(sFolderNameCurrent) > 0 then
  ' we are not in the root folder, so we can go up
  sFolderNameUpper = Left(sFolderNameCurrent,Len(sFolderNameCurrent)-1)
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
  Response.write("<a href='?fld=" & sFolderNameUpper & "'>Back to folder: <strong>" & sFolderNameUpper & "</strong></a></br>")
end if

'
' List of subFOLDERS in current folder
'
set objFolderCurrent=fs.GetFolder(sFolderNameRoot & sFolderNameCurrent)

LW("---- FOLDERS ----")
for each x in objFolderCurrent.SubFolders
  'Print the name of all subfolders in the test folder
  if (x.Name <> "Javascripts") then
    sHTML = fnSubfolder(sFolderNameCurrent, x.Name)
    Response.write(sHTML)
  end if
next
LW("---- FOLDERS ----")
'
' List of FILES in current folder
'
LW("==== FILES ====")
Response.write("<div id='idMovieContainer' class='clsAllMovies'>" & vbCRLF)
for each objFolderCurrent in objFolderCurrent.files
  'Print the name of all subfolders in the test folder
  if Ucase(Right(objFolderCurrent.Name, 4)) = ".MP4" then
    sHTML = fnMovieBox(sFolderNameRoot, sFolderNameCurrent, objFolderCurrent.Name)
    Response.write(sHTML)
  end if
next
Response.write("</div>" & vbCRLF)
LW("==== FILES ====")

'============= CLOSING PAGE CODE ============='
set objFolderCurrent=nothing
set fs=nothing
%>
<script src="Javascripts/jquery-3.2.1.min.js"></script>
<script src="Javascripts/jquery.fancybox.min.js"></script>
</body>
</html>
<%

'======================================================================
' This function creates HTML code for a single subfoler
'======================================================================
function fnSubfolder(sFolderNameCurrent, sSubfolderName)
  dim sFolderNameSubFolderLnkTemplate 
  sFolderNameSubFolderLnkTemplate = ""
  sFolderNameSubFolderLnkTemplate = sFolderNameSubFolderLnkTemplate & "<div id='idMovieContainer' class='clsFolderContainer'>" & vbCRLF
  sFolderNameSubFolderLnkTemplate = sFolderNameSubFolderLnkTemplate & "  <div id='idMoviePoster' class='clsMoviePosterBox'>" & vbCRLF
  sFolderNameSubFolderLnkTemplate = sFolderNameSubFolderLnkTemplate & "    <a href='?fld=@SubFolderLinks'>" & vbCRLF
  sFolderNameSubFolderLnkTemplate = sFolderNameSubFolderLnkTemplate & "    <img src='./FolderIcon.png' class='clsMoviePosterPic' width='90'>" & vbCRLF
  sFolderNameSubFolderLnkTemplate = sFolderNameSubFolderLnkTemplate & "    </a>" & vbCRLF
  sFolderNameSubFolderLnkTemplate = sFolderNameSubFolderLnkTemplate & "  </div>" & vbCRLF
  sFolderNameSubFolderLnkTemplate = sFolderNameSubFolderLnkTemplate & "  <p class='clsMovieTitle'>" & vbCRLF
  sFolderNameSubFolderLnkTemplate = sFolderNameSubFolderLnkTemplate & "  <a href='?fld=@SubFolderLinks'>" & vbCRLF
  sFolderNameSubFolderLnkTemplate = sFolderNameSubFolderLnkTemplate & "  @SubFolderName" & vbCRLF
  sFolderNameSubFolderLnkTemplate = sFolderNameSubFolderLnkTemplate & "  </a>" & vbCRLF
  sFolderNameSubFolderLnkTemplate = sFolderNameSubFolderLnkTemplate & "  </p>" & vbCRLF
  sFolderNameSubFolderLnkTemplate = sFolderNameSubFolderLnkTemplate & "</div>" & vbCRLF






  sFolderNameSubFolderLnkAcrual = Replace(sFolderNameSubFolderLnkTemplate, "@SubFolderLinks", sFolderNameCurrent & sSubfolderName)
  sFolderNameSubFolderLnkAcrual = Replace(sFolderNameSubFolderLnkAcrual, "@SubFolderName", sSubfolderName)
  fnSubfolder = sFolderNameSubFolderLnkAcrual
end function

'======================================================================
' This function creates HTML code for a single video file
'======================================================================
function fnMovieBox(sRootFolder, sWorkingFolder, sFileName)
  dim fs
  LW("")
  LW("function fnMovieBox - - - - - - - - - - - - - - - -")
  LW("sRootFolder    =" & sRootFolder)
  LW("sWorkingFolder =" & sWorkingFolder)
  LW("sFileName      =" & sFileName)
  '
  ' Define template HTML code
  '
  sMovieBoxTemplate = ""
  sMovieBoxTemplate = sMovieBoxTemplate & "<div id='idMovieContainer' class='clsMovieContainer'>" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "  <div id='idMoviePoster' class='clsMoviePosterBox'>" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "    <a data-fancybox data-width='1280' href='@MovieLink'>" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "    <img src='@PicURL' class='clsMoviePosterPic'>" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "    </a>" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "  </div>" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "  <span class='clsMovieTitle'>" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "  <a data-fancybox data-width='1280' href='@MovieLink'>" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "  @MovieFileName" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "  </a>" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "  </span>" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "</div>" & vbCRLF & vbCRLF
  
  '
  ' Translate template html code into actual html code 
  ' by replacing "@parameter" strings with real values
  '
  ' Make MP4 file name a little bit nicer
  sFileNameUpdt = Replace(sFileName,     "_", " ")
  sFileNameUpdt = Replace(sFileNameUpdt, "-", " ")
  sFileNameUpdt = Replace(sFileNameUpdt, ".", " ")
  sFileNameUpdt = Replace(sFileNameUpdt, "mp4", "")
  '
  ' Check, does poster exist or not
  set fs=Server.CreateObject("Scripting.FileSystemObject")
  sFileNamePoster = Left(sFileName, Len(sFileName) - 4) & ".jpg"
  LW("sFileNamePoster=" & sRootFolder & sWorkingFolder & sFileNamePoster)
  if (fs.FileExists(sRootFolder & sWorkingFolder & sFileNamePoster)) then
    sFileNamePoster = sWorkingFolder & sFileNamePoster
  else
    sFileNamePoster = "EmptyPoster.jpg"
  end if 
  set fs=Nothing
  sMovieBoxCurrent = Replace(sMovieBoxTemplate, "@PicURL",        sFileNamePoster)
  sMovieBoxCurrent = Replace(sMovieBoxCurrent,  "@MovieFileName", sFileNameUpdt)
  sMovieBoxCurrent = Replace(sMovieBoxCurrent,  "@MovieLink",     sWorkingFolder & sFileName)
  sMovieBoxCurrent = Replace(sMovieBoxCurrent,  "@MovieFileName", sFileName)
  fnMovieBox = sMovieBoxCurrent
end function

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