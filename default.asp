<!DOCTYPE html>

<!-- #include file="ASP\fnFormatDosPath.asp" -->
<!-- #include file="ASP\fnLW.asp" -->
<!-- #include file="ASP\fnSubfolder.asp" -->
<!-- #include file="ASP\fnTypeOTHER_controller.asp" -->
<!-- #include file="ASP\fnTypeMOVIE_controller.asp" -->
<!-- #include file="ASP\fnTypePDF_controller.asp" -->
<!-- #include file="ASP\fnTypeM3U_controller.asp" -->
<!-- #include file="ASP\fnTypePIC_controller.asp" -->
<!-- #include file="ASP\fnTypeLNK_controller.asp" -->
<!-- #include file="ASP\fnUpperFolder.asp" -->
<%
' Root folder for all DATA-files MUST be (current-folder-where-DEFUALT.ASP-file-is)+('\Data\')
' Suffix **DOS indicates that this path is in DOS notation (as IIS sees it C:\inetpub\wwwroot\html\ddd\)  
' in opposite to WEB notation /html/ddd/
sRootFolderDOS = Server.MapPath(".") & "\Data\"
LW("0 sRootFolderDOS=" & sRootFolderDOS)

' Current filer counts from the sRootFolderDOS
sCurrentFldrDOS = Request.QueryString("fld")
sCurrentFldrDOS = fnFormatDosPath(sCurrentFldrDOS) ' removing leading and adding tailing backslashes ('\')

' Make sure that we have correct (existing) current folder name
' in case of error just set it to `home` i.e. /DATA/ folder name
set fs=Server.CreateObject("Scripting.FileSystemObject")
LW("Check folder >>>  " & sRootFolderDOS & sCurrentFldrDOS)
if not (fs.FolderExists(sRootFolderDOS & sCurrentFldrDOS)) then
  LW("ERROR: current folder '" & sCurrentFldrDOS & "' does not exist.")
  LW("Reset current folder to NULL")
  sCurrentFldrDOS = ""
end if

LW("sRootFolderDOS   =" & sRootFolderDOS)
LW("sCurrentFldrDOS  =" & sCurrentFldrDOS)

'=====================================
' Now we are done with all COMMON settings 
' and configuration.
' The rest will depend on folder' specific 
' format (purpose)
' As of now we have 4 types of folder-specific-formats
' PIC - set of image files (JPG/PNG/GIF/TIFF) in the folder with a possible bunch of MP3 files as background music
' MOV - set of MP4 files with possible JPG file as a poster
' M3U - set of MP3 files
' PDF - set of PDF files
' LNK - loads chunk of HTML code
' ??? - OTHER folder (without extension) - set of sub-folders only (content of the folder should be ignoted)
'
'
' Folder's type defined through it's folder-name-extension, like "something-something.PIC"
' if current folder does not have any of these 4 types
' script will teat if as a folder without any content (with sub-folders only)
'
'=====================================
'
' Determinate what type of folder do we have as a current folder
sFolderType = Ucase(Left(Right(sCurrentFldrDOS,5),4))
LW("sFolderType=" & sFolderType)

Select Case sFolderType
  Case ".MOV"
    'Response.write("MOV type")
    Response.write(fnListOfMOVIES(sRootFolderDOS, sCurrentFldrDOS))
  Case ".PDF"
    'Response.write("PDF type")
    Response.write(fnListOfPDF(sRootFolderDOS, sCurrentFldrDOS))
  Case ".M3U"
    'Response.write("M3U type")
    Response.write(fnListOfM3U(sRootFolderDOS, sCurrentFldrDOS))
  Case ".PIC"
    'Response.write("PIC type")
    Response.write(fnListOfPIC(sRootFolderDOS, sCurrentFldrDOS))
  Case ".LNK"
    'Response.write("LNK type")
    Response.write(fnLoadHTML(sRootFolderDOS, sCurrentFldrDOS))
  Case Else
    'Response.write("OTHER type")
    Response.write(fnOTHER(sRootFolderDOS, sCurrentFldrDOS))
End Select
%>