<%
function fnOTHER(sRootFldrDOS, sCurrentFldrDOS)
  dim fso ' file system object
  dim oFldCurrent ' 
  dim sHtmlResult

  LW("function fnPictureBox - - - - - - - - - - - - - - - -*")
  set fso=Server.CreateObject("Scripting.FileSystemObject")
  set oFldCurrent=fso.GetFolder(sRootFldrDOS & sCurrentFldrDOS)
  LW("oFldCurrent.path=" & sRootFldrDOS & sCurrentFldrDOS)
  '
  '   Read WHOLE template HTML file into memory
  sTemplateFileName = sRootFldrDOS & "..\ASP\fnTypeOTHER_template.html"
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
  sHtmlBottomOpen = arrSections(3)
  sHtmlLeftBottomLoop = arrSections(4) 
  sHtmlBottomClose = arrSections(5)
  sHtmlLeftFooter = arrSections(6)
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
    sHtmlBottomOpen = ""
    sHtmlBottomClose = ""
  End If

  fnOTHER = sHtmlLeftHeader & _
            sDynamicBreadCrumbs & _
            sHtmlTopClose & _
            sHtmlBottomOpen & _
            sDynamicFolderHtml & _
            sHtmlBottomClose & _
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
  fnCheckFolderIcon = sResult
end function

%>