<%
function fnListOfPDF(sRootFldrDOS, sCurrentFldrDOS)
  dim fso ' file system object
  dim oFldCurrent ' 
  dim sHtmlResult

  LW("function fnPictureBox - - - - - - - - - - - - - - - -*")
  set fso=Server.CreateObject("Scripting.FileSystemObject")
  set oFldCurrent=fso.GetFolder(sRootFldrDOS & sCurrentFldrDOS)
  LW("oFldCurrent.path=" & sRootFldrDOS & sCurrentFldrDOS)
  '
  ' Read WHOLE template HTML file into memory
  sTemplateFileName=sRootFldrDOS & "..\ASP\fnTypePDF_template.html"
  LW ("14+sTemplateFileName=" & sTemplateFileName)
  set objTemplateFile=fso.OpenTextFile(sTemplateFileName,1,false)
  sSubfolderTemplate=objTemplateFile.ReadAll
  objTemplateFile.close
  'LW ("sSubfolderTemplate=" & sSubfolderTemplate)

  ' Split whole into sections base on SPLIT-MARKERS
  arrSections = Split(sSubfolderTemplate, "<!-- SplitMarker -->")
  LW("Ubound(arrSections)=" & Ubound(arrSections))
  '
  ' ////////////////////////////////////////////////////////////////////////
  ' Debug section begin
  For iii=0 to Ubound(arrSections)
    LW(iii & " ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
    LW(arrSections(iii))
  Next
  LW("ZZ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
  ' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  ' Debug section end

  sDynamicBreadCrumbs = ""
  arrBreadCrumbs = Split(sCurrentFldrDOS, "\")
  LW (" EQ UBound=" & UBound(arrBreadCrumbs))
  sHtmlLeftHeader = arrSections(0)
  sHtmlLeftTopLoop = arrSections(1)
  sHtmlLeftTopClose = arrSections(2) 
  sHtmlLeftBottomOpen = arrSections(3)
  sHtmlLeftBottomLoop = arrSections(4)
  sHtmlLeftBottomClose = arrSections(5) 
  sHtmlLeftClose = arrSections(6)
  sHtmlRightLoop = arrSections(7)
  sHtmlRightClose = arrSections(8)
  sHtmlPageClose = arrSections(9)
  ' 
  '
  ' Show uppaer folders list
  '
  sBreadCrumbsRelPathDOS = ""
  For iIdx=0 To UBound(arrBreadCrumbs) - 1
    sBreadCrumbsRelPathDOS = sBreadCrumbsRelPathDOS & arrBreadCrumbs(iIdx) & "\"
    sDynamicAlignment = "<br/>"
    ' I added the line above Dec 19. Erich 

    LW("sBreadCrumbsRelPathDOS=" & sBreadCrumbsRelPathDOS)
    sBrCru = sHtmlLeftTopLoop
    sBrCru = Replace(sBrCru, "@BreadCrumbHref", sBreadCrumbsRelPathDOS)
   'sBrCru = Replace(sBrCru, "@BreadCrumbTitle", arrBreadCrumbs(iIdx))
    sBrCru = Replace(sBrCru, "@BreadCrumbText", left(arrBreadCrumbs(iIdx), 30))
    sBrCru = Replace(sBrCru, "@BreadCrumbPipUp", arrBreadCrumbs(iIdx))
    ' I changed the 2 lines above Dec 19. Erich
    sDynamicBreadCrumbs = sDynamicBreadCrumbs & sBrCru & sDynamicAlignment
  Next


  sDynaHtmlFolder = ""
  For Each oSubFolderCurr in oFldCurrent.SubFolders
  sFolderIcon = fnCheckFolderIcon(sRootFldrDOS, sCurrentFldrDOS, oSubFolderCurr.Name)
    'Response.Write(oSubFolderCurr.Name & oSubFolderCurr.Name & "</br>")
    '
    ' Create current chunk of HTML code based on the template and current file name
    'sCurrentHtml = Replace(sHtmlBottomOpen, "@FldName", sCurrentFldrDOS & oSubFolderCurr.Name)
    'sCurrentHtml = Replace(sCurrentHtml,   "@FldTitle", oSubFolderCurr.Name)
    'sCurrentHtml = Replace(sCurrentHtml, "@SubFolderIcon", sFolderIcon)
    sHtmlLeftBottomCurrent = Replace(sHtmlLeftBottomLoop, "@FldName",  sCurrentFldrDOS & oSubFolderCurr.Name)
    sHtmlLeftBottomCurrent = Replace(sHtmlLeftBottomCurrent,   "@FldTitle", oSubFolderCurr.Name)
    sHtmlLeftBottomCurrent = Replace(sHtmlLeftBottomCurrent, "@SubFolderIcon", sFolderIcon)
    '
    ' I added the line above Dec 18. Erich
    '
    ' Add dynamically created HTML code to the rest of the HTML page
    'Response.Write("<b>AAAA</b></br>")
    'sDynaHtmlFolder = sDynaHtmlFolder & sCurrentHtml EQ
    sDynaHtmlFolder = sDynaHtmlFolder & sHtmlLeftBottomCurrent
  Next
  ' Erich  Dec. 19 
  ' Do not show LEFT-BOTTOM DIV with subfolders if it is empty
    If sDynaHtmlFolder = "" Then
       sHtmlLeftBottomOpen = ""
       sHtmlLeftBottomClose = ""
    End If

  sDynaHtmlPDF = ""
  For Each oFileCurr in oFldCurrent.Files
    'Response.Write ("oFileCurr.Name=" & oFileCurr.Name & "</br>")
    sFileExt = Right(oFileCurr.Name, 4)
    If UCase(sFileExt) = ".PDF" Then
      '
      ' Create current chunk of HTML code based on the template and current file name
      sCurrentHtml = Replace(sHtmlRightLoop, "@PdfName",  sCurrentFldrDOS & oFileCurr.Name)
      sCurrentHtml = Replace(sCurrentHtml,   "@PdfTitle", oFileCurr.Name)
      '
      ' Add dynamically created HTML code to the rest of the HTML page
      sDynaHtmlPDF = sDynaHtmlPDF & sCurrentHtml
      LW ("sCurrentHtml=" & sCurrentHtml)
    End If
  Next
  'response.write("</br>")
  
  
  fnListOfPDF = sHtmlLeftHeader & _
                sDynamicBreadCrumbs & _ 
                sHtmlLeftTopClose & _ 
                sHtmlLeftBottomOpen & _
                sDynaHtmlFolder & _ 
                sHtmlLeftBottomClose & _ 
                sHtmlLeftClose & _
                sDynaHtmlPDF & _ 
                sHtmlRightClose & _
                sHtmlPageClose
end function


%>