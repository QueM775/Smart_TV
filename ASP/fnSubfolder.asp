<%
'======================================================================
' This function creates HTML code for a single subfoler
'======================================================================
function fnSubfolder(sCUrrentFldrDOS, sSubfolderName)
  dim sTemplate 
  sTemplate = ""
  sTemplate = sTemplate & "<a href='?fld=@SubFolderLinks'>" & vbCRLF
  sTemplate = sTemplate & "  <div id='idMovieContainer' class='clsFolderContainer'>" & vbCRLF
  sTemplate = sTemplate & "    <div id='idMoviePoster' class='clsMoviePosterBox'>" & vbCRLF
  sTemplate = sTemplate & "      <img src='./Images/FolderIcon.png' class='clsMoviePosterPic' width='90'>" & vbCRLF
  sTemplate = sTemplate & "    </div>" & vbCRLF
  sTemplate = sTemplate & "    <p class='clsMovieTitle'>" & vbCRLF
  sTemplate = sTemplate & "    @SubFolderName" & vbCRLF
  sTemplate = sTemplate & "    </p>" & vbCRLF
  sTemplate = sTemplate & "  </div>" & vbCRLF
  sTemplate = sTemplate & "</a>" & vbCRLF

  sFolderNameSubFolderLnkAcrual = Replace(sTemplate, "@SubFolderLinks", sCUrrentFldrDOS & sSubfolderName)
  sFolderNameSubFolderLnkAcrual = Replace(sFolderNameSubFolderLnkAcrual, "@SubFolderName", sSubfolderName)
  fnSubfolder = sFolderNameSubFolderLnkAcrual
end function
%>