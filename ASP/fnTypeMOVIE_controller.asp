<%
function fnListOfMOVIES(sRootFldrDOS, sCurrentFldrDOS)
  dim fso ' file system object
  dim oFldCurrent ' 
  dim sHtmlResult

  LW("function fnPictureBox - - - - - - - - - - - - - - - -*")
  set fso=Server.CreateObject("Scripting.FileSystemObject")
  set oFldCurrent=fso.GetFolder(sRootFldrDOS & sCurrentFldrDOS)
  LW("oFldCurrent.path=" & sRootFldrDOS & sCurrentFldrDOS)
  '
  '   Read WHOLE template HTML file into memory
  sTemplateFileName = sRootFldrDOS & "..\ASP\fnTypeMOVIE_template.html"
  LW ("sTemplateFileName=" & sTemplateFileName)
  set objTemplateFile = fso.OpenTextFile(sTemplateFileName,1,false)
  sSubfolderTemplate = objTemplateFile.ReadAll
  objTemplateFile.close
  'LW ("sSubfolderTemplate=" & sSubfolderTemplate)

  ' Split whole into sections base on SPLIT-MARKERS
  arrSections = Split(sSubfolderTemplate, "<!-- SplitMarker -->")
  LW("Ubound(arrSections)=" & Ubound(arrSections))

  sDynamicBreadCrumbs = ""
  arrBreadCrumbs = Split(sCurrentFldrDOS, "\")
  LW ("UBound=" & UBound(arrBreadCrumbs))
  sHtmlLeftHeader      = arrSections(0)
  sHtmlLeftTopLoop     = arrSections(1)
  sHtmlLeftTopClose    = arrSections(2) 
  sHtmlLeftBottomOpen  = arrSections(3)
  sHtmlLeftBottomLoop  = arrSections(4) 
  sHtmlLeftBottomClose = arrSections(5)
  sHtmlPlayer          = arrSections(6) ' close left, open right side div, open/close player, open MP3 list
  sHtmlMP3listLoop     = arrSections(7) 
  sHtmlLeftFooter      = arrSections(8) ' close MP3 list, close right side div, close page
  LW("")
  '
  '
  ' Show uppaer folders list
  '
  sBreadCrumbsRelPathDOS = ""
  For iIdx=0 To UBound(arrBreadCrumbs) - 1
    sBreadCrumbsRelPathDOS = sBreadCrumbsRelPathDOS & arrBreadCrumbs(iIdx) & "\"
    sDynamicAlignment = "<br/>"

    'LW("sBreadCrumbsRelPathDOS=" & sBreadCrumbsRelPathDOS)
    sBrCrum = sHtmlLeftTopLoop
    sBrCrum = Replace(sBrCrum, "@BreadCrumbHref", sBreadCrumbsRelPathDOS)
    sBrCrum = Replace(sBrCrum, "@BreadCrumbText", left(arrBreadCrumbs(iIdx), 30))
    sBrCrum = Replace(sBrCrum, "@BreadCrumbPipUp", arrBreadCrumbs(iIdx))
    sDynamicBreadCrumbs = sDynamicBreadCrumbs & sBrCrum & sDynamicAlignment
  Next
  '
  '
  ' Scan subfolders in current folders
  '
  sDynamicFolderHtml = ""
  For Each oSubFolderCurr in oFldCurrent.SubFolders
    sFolderIcon = fnCheckFolderIcon(sRootFldrDOS, sCurrentFldrDOS, oSubFolderCurr.Name)
    'response.write("</br>" & oSubFolderCurr.Name)
    '
    ' Create current chunk of HTML code based on the template and current file name
    sHtmlLeftBottomCurrent = Replace(sHtmlLeftBottomLoop, "@FldName",  sCurrentFldrDOS & oSubFolderCurr.Name)
    sHtmlLeftBottomCurrent = Replace(sHtmlLeftBottomCurrent,   "@FldTitle", oSubFolderCurr.Name)
    sHtmlLeftBottomCurrent = Replace(sHtmlLeftBottomCurrent, "@SubFolderIcon", sFolderIcon)
    '
    ' Add dynamically created HTML code to the rest of the HTML page
    sDynamicFolderHtml = sDynamicFolderHtml & sHtmlLeftBottomCurrent
  Next
  'response.write("</br>")
  
  ' Mike 2019-12-10@21:26:38 
  ' Do not show LEFT-BOTTOM DIV with subfolders if it is empty
  If sDynamicFolderHtml = "" Then
    sHtmlLeftBottomOpen = ""
    sHtmlLeftBottomClose = ""
  End If
  '
  '
  '
  sDynaHtmlMovie = ""
  For Each oFileCurr in oFldCurrent.Files
    sFileCurrName = oFileCurr.Name
    'LW("sFileCurrName2=" & sFileCurrName)
    'Response.Write ("sFileCurrName=" & sFileCurrName & "</br>")
    sFileExt = Right(sFileCurrName, 4)
    If UCase(sFileExt) = ".MP4" Then
      ' Fix apostrophe issue
      sFileCurrNameShort = fnFixApostropheIssue(sRootFldrDOS & sCurrentFldrDOS, sFileCurrName)
      'LW ("sFileCurrNameShort = " & sRootFldrDOS & sCurrentFldrDOS & sFileCurrNameShort)
      '
      '
      ' Create current chunk of HTML code based on the template and current file name
      sCurrentHtml = Replace(sHtmlMP3listLoop, "@MovieName",  sCurrentFldrDOS & sFileCurrNameShort)
      sCurrentHtml = Replace(sCurrentHtml,   "@MovieTitle", fnMakeMovieTitleClean(sFileCurrNameShort))
      '"IMG\pdf-logo.png"
      ' Check, does poster exist or not
      ' Poster file must have the same name as correspondent *.MP4 video file
      ' but with *.JPG extension (currently script does not support any other graphic file formats like *.GIF; *.PNG; *.BMP)
      sFileNamePosterDosFull = sRootFldrDOS & sCurrentFldrDOS & sFileCurrNameShort & ".jpg"
      'LW("sFileCurrNameShort=" & sFileCurrNameShort)
      'LW("sFileNamePosterDosFull=" & sFileNamePosterDosFull)
      if (fso.FileExists(sFileNamePosterDosFull)) then
        iPosTailStarts = InStr(sFileNamePosterDosFull,"Data\") ' Find where relative path starts
        sFileNamePosterWeb = Mid(sFileNamePosterDosFull, iPosTailStarts) ' Convert (left-trim) full file name into relative one
      else
        sFileNamePosterWeb = "IMG\movie-icon.png"
      end if 
      sCurrentHtml = Replace(sCurrentHtml,   "@MoviePoster", sFileNamePosterWeb)
      '
      ' Add dynamically created HTML code to the rest of the HTML page
      sDynaHtmlMovie = sDynaHtmlMovie & sCurrentHtml
      'LW ("sCurrentHtml=" & sCurrentHtml)
    End If
  Next


  fnListOfMOVIES = sHtmlLeftHeader & _
            sDynamicBreadCrumbs & _
            sHtmlLeftTopClose & _
            sHtmlLeftBottomOpen & _
            sDynamicFolderHtml & _
            sHtmlLeftBottomClose & _
            sHtmlPlayer & _
            sDynaHtmlMovie & _
            sHtmlLeftFooter
end function

'
'
' Do your conversion here
'
function fnTranslateWhateverYouWantIntoWhateverYouNeed(sInput)
  fnTranslateWhateverYouWantIntoWhateverYouNeed = sInput
end function

function fnFixApostropheIssue(sPath, sFileCurrName)
  If (InStr(sFileCurrName,"'")) > 0 Then
    Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
    sNewName = Replace(sFileCurrName, "'", "`")
    fileSystemObject.MoveFile (sPath & sFileCurrName), (sPath & sNewName)
    fnFixApostropheIssue = sNewName
  Else
    fnFixApostropheIssue = sFileCurrName
  End If
end function

function fnMakeMovieTitleClean(sMovieTitle)
  sMovieTitle = Left(sMovieTitle, Len(sMovieTitle)-4) ' remove file name extension
  sMovieTitle = Replace(sMovieTitle,"-", " ")
  sMovieTitle = Replace(sMovieTitle,"_", " ")
  sMovieTitle = Replace(sMovieTitle,".", ". ")
  fnMakeMovieTitleClean = sMovieTitle
end function

function fnCheckFolderIcon(sRootFldrDOS, sCurrentFldrDOS, sSubFolderName)
  set fso=Server.CreateObject("Scripting.FileSystemObject")
  If (fso.FileExists(sRootFldrDOS & sCurrentFldrDOS & sSubFolderName & "\FolderIcon.png")) Then
    sResult = "./Data/" & sCurrentFldrDOS & sSubFolderName & "\FolderIcon.png"
  Else
    sResult = "./IMG/folder-down-icon.png"
  End If
  sResult = Replace(sResult, "\", "/")
  LW (sResult)
  fnCheckFolderIcon = sResult
end function

%>