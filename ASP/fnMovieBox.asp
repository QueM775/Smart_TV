<%
function fnMovieBox(sRootFldrDOS, sWorkingFolder)
' This function creates HTML code for a single video file
' base on fnMovieBox_template.html TEMPLATE 
' and content of a current working folder "sWorkingFolder"
'
' Dependencies:
'   fnMovieBox.asp 
'   fnMovieBox_template.html 
'   fnMovieBox_unit_test.asp
'   default.asp (CSS and JS)
'
'
  dim fso ' file system object
  dim oFldCurrent ' 
  dim sHtmlResult

  ' 
  'LW("function fnMovieBox - - - - - - - - - - - - - - - -*")
  'LW("sRootFldrDOS   =" & sRootFldrDOS)
  'LW("sWorkingFolder =" & sWorkingFolder)

  set fso=Server.CreateObject("Scripting.FileSystemObject")
  set oFldCurrent=fso.GetFolder(sRootFldrDOS & sCurrentFldrDOS)
  'LW("oFldCurrent.path=" & sRootFldrDOS & sCurrentFldrDOS)
  '
  '   Read WHOLE template HTML file into memory
  sTemplateFileName=sRootFldrDOS & "..\ASP\fnMovieBox_template.html"
  set objTemplateFile=fso.OpenTextFile(sTemplateFileName,1,false)
  sMovieBoxTemplate=objTemplateFile.ReadAll
  objTemplateFile.close
  LW ("sMovieBoxTemplate=" & sMovieBoxTemplate)

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