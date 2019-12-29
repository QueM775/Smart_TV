<%
function fnListOfM3U(sRootFldrDOS, sCurrentFldrDOS)
  dim fso ' file system object
  dim oFldCurrent ' 
  dim sHtmlResult

  LW("function fnPictureBox - - - - - - - - - - - - - - - -*")
  set fso=Server.CreateObject("Scripting.FileSystemObject")
  set oFldCurrent=fso.GetFolder(sRootFldrDOS & sCurrentFldrDOS)
  LW("oFldCurrent.path=" & sRootFldrDOS & sCurrentFldrDOS)
  '
  '   Read WHOLE template HTML file into memory
  sTemplateFileName = sRootFldrDOS & "..\ASP\fnTypeM3U_template.html"
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
  sHtmlLeftHeader = arrSections(0)
  sHtmlLeftTopLoop = arrSections(1)
  sHtmlTopClose = arrSections(2) 
  sHtmlLeftBottomOpen = arrSections(3)
  sHtmlLeftBottomLoop = arrSections(4) 
  sHtmlLeftBottomClose = arrSections(5)
  sHtmlPlayer = arrSections(6) ' close left, open right side div, open/close player, open MP3 list
  sHtmlMP3listLoop = arrSections(7) 
  sHtmlLeftFooter = arrSections(8) ' close MP3 list, close right side div, close page
  LW("")
  '
  '
  ' Show uppaer folders list
  '
  sBreadCrumbsRelPathDOS = ""
  For iIdx=0 To UBound(arrBreadCrumbs) - 1
    sBreadCrumbsRelPathDOS = sBreadCrumbsRelPathDOS & arrBreadCrumbs(iIdx) & "\"
    sDynamicAlignment = "<br/>"

    LW("sBreadCrumbsRelPathDOS=" & sBreadCrumbsRelPathDOS)
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
  ' Show album cover
  arrFolders = Split(sRootFldrDOS & sCurrentFldrDOS, "\")
  sCurrentFldrShort = arrFolders(UBound(arrFolders)-1)
  sCurrentFldrName = Left(sCurrentFldrShort, Len(sCurrentFldrShort) - 4) 'potential error if folder extention will be more/less then 4 char long
  sCurrentAlbumCoverFileName = sRootFldrDOS & sCurrentFldrDOS & sCurrentFldrName & ".jpg" 
  set fso=Server.CreateObject("Scripting.FileSystemObject")
  If (fso.FileExists(sCurrentAlbumCoverFileName)) Then
    sResult = "./Data/" & sCurrentFldrDOS & sCurrentFldrName & ".jpg"
  Else
    sResult = "./IMG/music-general-wallpaper.jpg"
  End If
  LW("%%% sResult=" & sResult)
  set fso=Nothing

  sHtmlPlayer = Replace(sHtmlPlayer, "@AlbumCover", sResult) 
  '
  '
  ' Load list of MP3 files from current folder
  '
  sLiSum = ""
  iFlag = 0
  for each oFileCrnt in oFldCurrent.files
    If Ucase(Right(oFileCrnt.Name, 4)) = ".MP3" then
      LW ("@@@@@@@@@@ " & oFileCrnt.Name & " @@@@@@@@@@@@@@")
      If (iFlag = 0) Then
        iFlag = 1
        sHtmlMP3listLoop = Replace(sHtmlMP3listLoop, "@FirstTrack", "Data\" & sCurrentFldrDOS & oFileCrnt.Name)
      End If
      'sLi = sLi & "<li><a href='Data\" & sCurrentFldrDOS & oFileCrnt.Name & "'>" & oFileCrnt.Name & "</a></li>" & vbCRLF
      sLi = Replace(sHtmlMP3listLoop, "@MP3link", "Data\" & sCurrentFldrDOS & oFileCrnt.Name)
      sLi = Replace(sLi, "@MP3Name", oFileCrnt.Name)
      sLiSum = sLiSum & sLi
    End If
  next ' for each



  fnListOfM3U = sHtmlLeftHeader & _
            sDynamicBreadCrumbs & _
            sHtmlTopClose & _
            sHtmlLeftBottomOpen & _
            sDynamicFolderHtml & _
            sHtmlLeftBottomClose & _
            sHtmlPlayer & _
            sLiSum & _
            sHtmlLeftFooter
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
  set fso = Nothing
  fnCheckFolderIcon = sResult
end function

%>