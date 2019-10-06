<%
'======================================================================
' This function creates HTML code for a single video file
'======================================================================
function fnPdfBox(sRootFldrDOS, sWorkingFolder, sFileName)
  dim fs
  LW("")
  LW("function fnMovieBox - - - - - - - - - - - - - - - -")
  LW("sRootFldrDOS   =" & sRootFldrDOS)
  LW("sWorkingFolder =" & sWorkingFolder)
  LW("sFileName      =" & sFileName)
  set fs=Server.CreateObject("Scripting.FileSystemObject")
  '
  ' Define template HTML code
  '
  sMovieBoxTemplate = ""
  '
  '   Read WHOLE template HTML file into memory
  sTemplateFileName=sRootFldrDOS & "..\ASP\fnPdfBox_template.html"
  set objTemplateFile=fs.OpenTextFile(sTemplateFileName,1,false)
  sMovieBoxTemplate=objTemplateFile.ReadAll
  objTemplateFile.close
  LW ("sMovieBoxTemplate=" & sMovieBoxTemplate)
  
  '
  ' Translate template html code into actual html code 
  ' by replacing "@parameter" strings with real values
  '
  ' Make MP4 file name a little bit nicer
  sFileNameUpdt = Replace(sFileName,     "_", " ")
  sFileNameUpdt = Replace(sFileNameUpdt, "-", " ")
  sFileNameUpdt = Replace(sFileNameUpdt, ".", " ")
  sFileNameUpdt = Replace(sFileNameUpdt, "pdf", "")
  '
  ' Check, does poster exist or not
  ' Poster file must have the same name as correspondent *.MP4 video file
  ' but with *.JPG extension (currently script does not support any other graphic file formats like *.GIF; *.PNG; *.BMP)
  sFileNamePoster = Left(sFileName, Len(sFileName) - 4) & ".jpg"
  LW("sFileNamePoster=" & sRootFldrDOS & sWorkingFolder & sFileNamePoster)
  if (fs.FileExists(sRootFldrDOS & sWorkingFolder & sFileNamePoster)) then
    sFileNamePoster = "Data\" & sWorkingFolder & sFileNamePoster
  else
    sFileNamePoster = "Images/pdf-logo.png"
  end if 
  set fs=Nothing
  sMovieBoxCurrent = Replace(sMovieBoxTemplate, "@PicURL",        sFileNamePoster)
  sMovieBoxCurrent = Replace(sMovieBoxCurrent,  "@MovieFileName", sFileNameUpdt)
  sMovieBoxCurrent = Replace(sMovieBoxCurrent,  "@MovieLink",     sWorkingFolder & sFileName)
  sMovieBoxCurrent = Replace(sMovieBoxCurrent,  "@MovieFileName", sFileName)
  fnPdfBox = sMovieBoxCurrent
end function
%>