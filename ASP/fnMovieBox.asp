<%
'======================================================================
' This function creates HTML code for a single video file
'======================================================================
function fnMovieBox(sRootFldrDOS, sWorkingFolder, sFileName)
  dim fs
  LW("")
  LW("function fnMovieBox - - - - - - - - - - - - - - - -")
  LW("sRootFldrDOS   =" & sRootFldrDOS)
  LW("sWorkingFolder =" & sWorkingFolder)
  LW("sFileName      =" & sFileName)
  '
  ' Define template HTML code
  '
  sMovieBoxTemplate = ""
  sMovieBoxTemplate = sMovieBoxTemplate & "<a data-fancybox data-width='1280' href='Data\@MovieLink'>" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "  <div id='idMovieContainer' class='clsMovieContainer'>" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "    <div id='idMoviePoster' class='clsMoviePosterBox'>" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "      <img src='@PicURL' class='clsMoviePosterPic'>" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "    </div>" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "    <span class='clsMovieTitle'>" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "    @MovieFileName" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "    </span>" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "  </div>" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "</a>" & vbCRLF & vbCRLF
  
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
  ' Poster file must have the same name as correspondent *.MP4 video file
  ' but with *.JPG extension (currently script does not support any other graphic file formats like *.GIF; *.PNG; *.BMP)
  set fs=Server.CreateObject("Scripting.FileSystemObject")
  sFileNamePoster = Left(sFileName, Len(sFileName) - 4) & ".jpg"
  LW("sFileNamePoster=" & sRootFldrDOS & sWorkingFolder & sFileNamePoster)
  if (fs.FileExists(sRootFldrDOS & sWorkingFolder & sFileNamePoster)) then
    sFileNamePoster = "Data\" & sWorkingFolder & sFileNamePoster
  else
    sFileNamePoster = "Images/EmptyPoster.jpg"
  end if 
  set fs=Nothing
  sMovieBoxCurrent = Replace(sMovieBoxTemplate, "@PicURL",        sFileNamePoster)
  sMovieBoxCurrent = Replace(sMovieBoxCurrent,  "@MovieFileName", sFileNameUpdt)
  sMovieBoxCurrent = Replace(sMovieBoxCurrent,  "@MovieLink",     sWorkingFolder & sFileName)
  sMovieBoxCurrent = Replace(sMovieBoxCurrent,  "@MovieFileName", sFileName)
  fnMovieBox = sMovieBoxCurrent
end function
%>