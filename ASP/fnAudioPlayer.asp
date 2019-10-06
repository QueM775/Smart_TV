<%
function fnAudioPlayer(sRootFolderDosFull, sWorkFolderRelative)
  ' This function creates AudioPlayer HTML code based on the Player's template "ASP\fnAudioPlayer_template.html"
  ' and content of a current working folder "sWorkFolderRelative"
  '
  ' Dependencies:
  ' This file must be modified in conjunction with
  '   defualt.asp (CSS and JS)
  '   fnAudioPlayer.asp
  '   fnAudioPlayer_template.html
  '   AudioPlayerA.js
  '
  '
  dim sHtmlTemplate
  set fso=Server.CreateObject("Scripting.FileSystemObject")
  set oFldCurrent=fso.GetFolder(sRootFolderDosFull & sWorkFolderRelative)
  '
  '   Read WHOLE template HTML file into memory
  sTemplateFileName=sRootFolderDosFull & "..\ASP\fnAudioPlayer_template.html"
  set objTemplateFile=fso.OpenTextFile(sTemplateFileName,1,false)
  sHtmlTemplate=objTemplateFile.ReadAll
  objTemplateFile.close
  'LW ("sHtmlTemplate=" & sHtmlTemplate)
  '
  ' Split whole into sections base on SPLIT-MARKERS
  ' It MUST be 3 section in the template HTML file
  ' 1-st section is a HEADER
  ' 2-nd section is dynamically created list
  ' 2-rd section is a FOOTER
  arrSections = Split(sHtmlTemplate, "<!-- SplitMarker -->")
  '
  ' Loop through files in current folder and update+duplicate 2-nd section with actual files names
  sLiSum = ""
  iFlag = 0
  for each oFileCrnt in oFldCurrent.files
    If Ucase(Right(oFileCrnt.Name, 4)) = ".MP3" then
      'LW ("@@@@@@@@@@ " & oFileCrnt.Name & " @@@@@@@@@@@@@@")
      If (iFlag = 0) Then
        iFlag = 1
        arrSections(0) = Replace(arrSections(0), "@FirstTrack", "Data\" & sWorkFolderRelative & oFileCrnt.Name)
      End If
      'sLi = sLi & "<li><a href='Data\" & sWorkFolderRelative & oFileCrnt.Name & "'>" & oFileCrnt.Name & "</a></li>" & vbCRLF
      sLi = Replace(arrSections(1), "@MP3link", "Data\" & sWorkFolderRelative & oFileCrnt.Name)
      sLi = Replace(sLi, "@MP3Name", oFileCrnt.Name)
      sLiSum = sLiSum & sLi
    End If
  next ' for each
  'LW("sLiSum=" & sLiSum)

  set fso=Nothing
  set oFldCurrent=Nothing

if (Len(sLiSum) > 0) Then
  fnAudioPlayer = arrSections(0) & vbCRLF & sLiSum & arrSections(2)
else
  fnAudioPlayer = ""
End If
end function
%>
