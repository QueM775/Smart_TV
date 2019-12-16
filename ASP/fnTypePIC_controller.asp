<%
function fnListOfPIC(sRootFldrDOS, sCurrentFldrDOS)
  dim fso ' file system object
  dim oFldCurrent ' 
  dim sHtmlResult

  'LW("function fnPictureBox - - - - - - - - - - - - - - - -*")
  set fso=Server.CreateObject("Scripting.FileSystemObject")
  set oFldCurrent=fso.GetFolder(sRootFldrDOS & sCurrentFldrDOS)
  'LW("oFldCurrent.path=" & sRootFldrDOS & sCurrentFldrDOS)
  '
  ' Read WHOLE template HTML file into memory
  sTemplateFileName=sRootFldrDOS & "..\ASP\fnTypePIC_template.html"
  'LW ("14+sTemplateFileName=" & sTemplateFileName)
  set objTemplateFile=fso.OpenTextFile(sTemplateFileName,1,false)
  sSubfolderTemplate=objTemplateFile.ReadAll
  objTemplateFile.close
  'LW ("sSubfolderTemplate=" & sSubfolderTemplate)

  ' Split whole into sections base on SPLIT-MARKERS
  arrSections = Split(sSubfolderTemplate, "<!-- SplitMarker -->")
  sHtmlLeftHeader     = arrSections(0)
  sHtmlLeftTopLoop    = arrSections(1)
  sHtmlTopClose       = arrSections(2)
  sHtmlBottomOpen     = arrSections(3)
  sHtmlCloseSubFldOpenPlayer = arrSections(4)
  sHtmlBottomClose    = arrSections(5)
  sHtmlLeftFooter3    = arrSections(6)
  sHtmlPictureLoop    = arrSections(7)
  sHtmlLeftFooter1    = arrSections(8)
  sHtmlLeftFooter     = arrSections(9)
  
  'LW("Ubound(arrSections)=" & Ubound(arrSections))
  '
  ' ////////////////////////////////////////////////////////////////////////
  ' Debug section begin
  For iii=0 to Ubound(arrSections)
  '  LW(iii & " ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
  '  LW(arrSections(iii))
  Next
  'LW("ZZ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
  ' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  ' Debug section end

  sDynamicBreadCrumbs = ""
  arrBreadCrumbs = Split(sCurrentFldrDOS, "\")
  'LW ("UBound=" & UBound(arrBreadCrumbs))
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
    '
    ' Create current chunk of HTML code based on the template and current file name
    sHtmlLeftBottomCurrent = Replace(sHtmlBottomOpen, "@FldName", sCurrentFldrDOS & oSubFolderCurr.Name)
    sHtmlLeftBottomCurrent = Replace(sHtmlLeftBottomCurrent,   "@FldTitle", oSubFolderCurr.Name)
    sHtmlLeftBottomCurrent = Replace(sHtmlLeftBottomCurrent, "@SubFolderIcon", sFolderIcon)
    '
    ' Add dynamically created HTML code to the rest of the HTML page
    sDynamicFolderHtml = sDynamicFolderHtml & sHtmlLeftBottomCurrent
  Next
  '
  '


  sLiSum = ""
  iFlag = 0
  for each oFileCrnt in oFldCurrent.files
    If Ucase(Right(oFileCrnt.Name, 4)) = ".MP3" then
      If (iFlag = 0) Then
        iFlag = 1
        sHtmlCloseSubFldOpenPlayer = Replace(sHtmlCloseSubFldOpenPlayer, "@MP3link", "Data\" & sCurrentFldrDOS & oFileCrnt.Name )
      End If
      sLi = sLi & "<li><a href='Data\" & sCurrentFldrDOS & oFileCrnt.Name & "'>" & oFileCrnt.Name & "</a></li>" & vbCRLF
      sLi = Replace(arrSections(5), "@MP3link", "Data\" & sCurrentFldrDOS & oFileCrnt.Name)
      sLi = Replace(sLi, "@MP3Name", oFileCrnt.Name)
      sLiSum = sLiSum & sLi
    End If
  next ' for each

  sPicSum = ""
  set oFldCurrent=fso.GetFolder(sRootFldrDOS & sCurrentFldrDOS)
  for each oFileCrnt in oFldCurrent.files
    If (InStr("~.JPG.JPEG.PNG.GIF.TIF.TIFF.BMP", Ucase(Right(oFileCrnt.Name, 4)) ) > 0) _
        AND (Ucase(oFileCrnt.Name) <> Ucase("FolderIcon.png")) Then
      sCurrentImageFileName = oFileCrnt.Name
      LW(sRootFldrDOS)
      LW(sCurrentFldrDOS)
      LW(sCurrentImageFileName)
      LW("~~~~~~~~~~~~~~~~~~~")
      sCurrentImageFileName = fnFixApostropheIssue(sRootFldrDOS & sCurrentFldrDOS, sCurrentImageFileName)
      sPicCurr = Replace(sHtmlPictureLoop, "@PicSrc", sCurrentFldrDOS & sCurrentImageFileName)
      sPicSum = sPicSum & sPicCurr
    End If
  next
  'response.write("</br>")
  '              http-header         BreadCrumbs           buffer          Sub-Folders     PLayerItself     M3U-list    Suffix
  fnListOfPIC = sHtmlLeftHeader & _
                sDynamicBreadCrumbs & _
                sHtmlTopClose & _
                sDynamicFolderHtml & _
                sHtmlCloseSubFldOpenPlayer & _
                sLiSum & _
                arrSections(6) & _
                sPicSum & _
                arrSections(UBound(arrSections))
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