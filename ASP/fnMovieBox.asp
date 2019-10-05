<%
'======================================================================
' This function creates HTML code for a single video file
'======================================================================
function fnMovieBox(sRootFldrDOS, sWorkingFolder)
  dim fso ' file system object
  dim oFldCurrent ' 
  dim sHtmlResult

  ' 
  'LW("function fnMovieBox - - - - - - - - - - - - - - - -*")
  'LW("sRootFldrDOS   =" & sRootFldrDOS)
  'LW("sWorkingFolder =" & sWorkingFolder)
  '===================================
  '== HTML template code
  '===================================
  sMovieBoxTemplate = ""
  sMovieBoxTemplate = sMovieBoxTemplate & "<a data-fancybox data-width='1280' href='Data\@MovieLink'>" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "  <div id='idMovieContainer' class='clsMovieContainer'>" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "    <div id='idMoviePoster' class='clsMoviePosterBox'>" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "      <img src='@PicURL' class='clsMoviePosterPic'>" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "    </div>" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "    <span class='clsMovieTitle'>" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "    @MovieTitle" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "    </span>" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "  </div>" & vbCRLF
  sMovieBoxTemplate = sMovieBoxTemplate & "</a>" & vbCRLF & vbCRLF

  set fso=Server.CreateObject("Scripting.FileSystemObject")
  set oFldCurrent=fso.GetFolder(sRootFldrDOS & sCurrentFldrDOS)
  'LW("oFldCurrent.path=" & sRootFldrDOS & sCurrentFldrDOS)

  sHtmlResult = ""
  for each oFileCrnt in oFldCurrent.files
    'LW ("## Current file name = " & oFileCrnt.Name)
    'Print the name of all subfolders in the test folder
    if Ucase(Right(oFileCrnt.Name, 4)) = ".MP4" then
      ' Process **ONLY** MP4 files
      sFileNameDosFull = sRootFldrDOS & sCurrentFldrDOS & oFileCrnt.Name
      'LW ("sFileNameDosFull = " & sFileNameDosFull)
      ' Make MP4 file name a little bit nicer
      sFileNameUpdt = Replace(oFileCrnt.Name, "_", " ")
      sFileNameUpdt = Replace(sFileNameUpdt,  "-", " ")
      sFileNameUpdt = Replace(sFileNameUpdt,  ".", " ")
      sFileNameUpdt = Replace(sFileNameUpdt,  "mp4", "")
      'LW ("Easy to read current file name = " & sFileNameUpdt)
      '
      ' Check, does poster exist or not
      ' Poster file must have the same name as correspondent *.MP4 video file
      ' but with *.JPG extension (currently script does not support any other graphic file formats like *.GIF; *.PNG; *.BMP)
      sFileNamePosterDosFull = Left(sFileNameDosFull, Len(sFileNameDosFull) - 4) & ".jpg"
      'LW("sFileNamePosterDosFull=" & sFileNamePosterDosFull)
      if (fso.FileExists(sFileNamePosterDosFull)) then
        iPosTailStarts = InStr(sFileNamePosterDosFull,"Data\") ' Find where relative path starts
        sFileNamePosterWeb = Mid(sFileNamePosterDosFull, iPosTailStarts) ' Convert (left-trim) full file name into relative one
      else
        sFileNamePosterWeb = "Images/EmptyPoster.jpg"
      end if 
      '
      ' Translate template html code into actual html code 
      ' by replacing "@parameter" strings with real values
      '
      sMovieBoxCurrent = Replace(sMovieBoxTemplate, "@PicURL",        sFileNamePosterWeb)
      sMovieBoxCurrent = Replace(sMovieBoxCurrent,  "@MovieTitle",    sFileNameUpdt)
      sMovieBoxCurrent = Replace(sMovieBoxCurrent,  "@MovieLink",     sWorkingFolder & oFileCrnt.Name)
    
      'LW("sHTML=" & sMovieBoxCurrent)
      sHtmlResult = sHtmlResult & sMovieBoxCurrent
    end if ' === ".MP4"
  next
  
  set fso=Nothing
  set oFldCurrent=Nothing
  'RETURN
  fnMovieBox = sHtmlResult
end function
%>