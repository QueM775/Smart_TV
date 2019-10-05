<!-- 
 This test file should not be in production environment

 This files is for testing purposes only. 
 It tests the fnMovieBox.asp module.
 This file must be run from the same folder where the main DEFAULT.ASP file is located
 because all paths to CSS/JS/ASP files are defined as it is in the DEFAULT.ASP

 Before run (open in browser) this file you must define:
 sRootFolderDosFull and
 sWorkingFolder
 variables, see below
-->
<%
' Next two parameters must be updated for each test case
dim sRootFolderDosFull
dim sWorkingFolder

sRootFolderDosFull = "I:\WEB\HTTP\erich\Smart_TV\Data\"
sWorkingFolder = ""
%>
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>My video player</title>

  <!-- CSS -->
  <link rel="stylesheet" type="text/css" href="CSS/jquery.fancybox.min.css">
  <link rel="stylesheet" type="text/css" href="CSS/movie.css">
</head>
<body>
<!-- #include file="asp/fnMovieBox.asp" -->
<!-- #include file="asp/fnLW.asp" -->

<%

sHTML_code = fnMovieBox(sRootFolderDosFull, sWorkingFolder)
LW("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
LW("sHTML_code = " & vbNewLine & sHTML_code)
LW("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")

Response.write(sHTML_code)

%>
<script src="Javascripts/jquery-3.2.1.min.js"></script>
<script src="Javascripts/jquery.fancybox.min.js"></script>

</body>