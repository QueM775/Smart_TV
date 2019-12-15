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
  LW ("UBound=" & UBound(arrBreadCrumbs))
  sBreadCrumbsRelPathDOS = ""
  For iIdx=0 To UBound(arrBreadCrumbs) - 1
    sBreadCrumbsRelPathDOS = sBreadCrumbsRelPathDOS & arrBreadCrumbs(iIdx) & "\"
    LW("sBreadCrumbsRelPathDOS=" & sBreadCrumbsRelPathDOS)
    sBrCru = arrSections(1)
    sBrCru = Replace(sBrCru, "@BreadCrumbHref", sBreadCrumbsRelPathDOS)
    sBrCru = Replace(sBrCru, "@BreadCrumbTitle", arrBreadCrumbs(iIdx))
    sDynamicBreadCrumbs = sDynamicBreadCrumbs & sBrCru
  Next


  sDynaHtmlFolder = ""
  For Each oSubFolderCurr in oFldCurrent.SubFolders
    'Response.Write(oSubFolderCurr.Name & oSubFolderCurr.Name & "</br>")
    '
    ' Create current chunk of HTML code based on the template and current file name
    sCurrentHtml = Replace(arrSections(3), "@FldName", sCurrentFldrDOS & oSubFolderCurr.Name)
    sCurrentHtml = Replace(sCurrentHtml,   "@FldTitle", oSubFolderCurr.Name)
    '
    ' Add dynamically created HTML code to the rest of the HTML page
    'Response.Write("<b>AAAA</b></br>")
    sDynaHtmlFolder = sDynaHtmlFolder & sCurrentHtml
  Next
  '
  '
  sDynaHtmlPDF = ""
  For Each oFileCurr in oFldCurrent.Files
    'Response.Write ("oFileCurr.Name=" & oFileCurr.Name & "</br>")
    sFileExt = Right(oFileCurr.Name, 4)
    If UCase(sFileExt) = ".PDF" Then
      '
      ' Create current chunk of HTML code based on the template and current file name
      sCurrentHtml = Replace(arrSections(5), "@PdfName",  sCurrentFldrDOS & oFileCurr.Name)
      sCurrentHtml = Replace(sCurrentHtml,   "@PdfTitle", oFileCurr.Name)
      '
      ' Add dynamically created HTML code to the rest of the HTML page
      sDynaHtmlPDF = sDynaHtmlPDF & sCurrentHtml
      LW ("sCurrentHtml=" & sCurrentHtml)
    End If
  Next
  'response.write("</br>")
  '              http-header         BreadCrumbs           buffer          Sub-Folders       buffer           PDF-icons        Suffix
  fnListOfPDF = arrSections(0) &  sDynamicBreadCrumbs & arrSections(2) & sDynaHtmlFolder & arrSections(4) & sDynaHtmlPDF & arrSections(6) & arrSections(UBound(arrSections))
end function
%>