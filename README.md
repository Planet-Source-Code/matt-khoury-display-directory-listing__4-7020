<div align="center">

## Display Directory Listing


</div>

### Description

First off let me say that this code is a modified version of Kaustav Acharya&#8217;s post on Planet Source Code. I simplified the code to better suit my needs. You should just be able to copy the code posted below, and in order for this code to work you MUST update the SELECT CASE statement and configure it to work on your server&#8217;s structure. I think the other code is a bit harder (not user friendly) to use/fix for inexperienced programmers. Once configured properly, this code will display the contents of a directory wherever this file is placed. I named the file to be inserted into the directory to display index.asp, just because my default web site is set up to read an index.asp file as the default web page. You can also get cute with this script as well, write another select case statement checking for the file types, and display the pertinent icon according to the file name. Not what I need this code for, but just a suggestion.

Happy Coding!
 
### More Info
 
Edit the SELECT CASE Statement to fit YOUR needs. YOU MUST EDIT THIS!

Here's how I use this file. Create an ASP named, "index.asp", or "default.asp" whichever you use. Then place the newly created file in the location where you wish the contents to be displayed. Now, you might have to edit your NT permissions in order to have at least "READ" permissions.

Returns the contents of a directory.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matt Khoury](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matt-khoury.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Server Side](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/server-side__4-31.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matt-khoury-display-directory-listing__4-7020/archive/master.zip)





### Source Code

```
<%@ Language=VBScript %>
<% Option Explicit %>
<!doctype html public "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>Directory Listing</title>
<style type="text/css">
TABLE
{
 BORDER-BOTTOM: 0px;
 BORDER-LEFT: 0px;
 BORDER-RIGHT: 0px;
 BORDER-TOP: 0px;
 FONT-FAMILY: tahoma,sans-serif;
 FONT-SIZE: 12px
}
</style>
</head>
<body bgcolor="#FFFFFF" text="#000000" link="blue" vlink="blue">
<%
	Dim strDirName, strMyPath, objFSO, strFolder, strFiles, intFileCount, strHost
		'---------------------------------------------------------------------------------------------
		strHost = Request.ServerVariables("HTTP_HOST")
		'-- REMEMBER TO EDIT THE FIRST CASE STATEMENT, OF THE CODE WILL NOT WORK
		'-- I ADD A CASE STATEMENT FOR EACH ENVIRONMENT THE SCRIPT WILL RESIDE SO I DON'T HAVE TO EDIT
		'-- THE TWO VARIABLES BELOW WHEN MOVED FROM DEVELOPEMNT TO PRODCUTION.
		Select Case strHost
			Case "acctdev.int.westgroup.com"
				strDirName = Server.MapPath("/global_planning/docs/")
				strMyPath = Server.MapPath("/global_planning/docs/")
			Case Else
				strDirName = Server.MapPath("/gp/docs/")
				strMyPath = Server.MapPath("/gp/docs/")
		End Select
		'---------------------------------------------------------------------------------------------
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set strFolder = objFSO.GetFolder(strDirName)
	Set strFiles = strFolder.Files
		intFileCount = strFolder.Files.Count
	Dim strPath		'-- PATH OF DIRECORY TO DISPLAY
	Dim objFolder	'-- FOLDER VARIABLE
	Dim objItem		'-- VARIABLE USED TO LOOP THROUGH AND DISPLAY THE CONTENTS
		strPath = strMyPath
	'--	CREATE THE FILE SYSTEM OBJECT
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	'--	SET THE OBJECT AND PASS IN THE PATH AS AN ARGUMENT
	Set objFolder = objFSO.GetFolder(strPath)
%>
<table align="center" border="2" bordercolor="#103052" cellspacing="0" cellpadding="2" width="100%">
<%
If intFileCount = "0" Then
%>
	<tr>
		<td colspan="5" nowrap>Sorry, there are no files located in this directory.</td>
	</tr>
</table>
<%
Else
%>
	<tr bgcolor="#103052">
		<td align="CENTER"><font color="#FFFFFF"><b>File Name:</b></font></td>
		<td align="CENTER"><font color="#FFFFFF"><b>Date Created:</b></font></td>
		<td align="CENTER"><font color="#FFFFFF"><b>File Type:</b></font></td>
	</tr>
<%
'--	FIRST OFF DEAL WITH THE SUBDIRECTORIES
	For Each objItem In objFolder.SubFolders
	'-- DEAL WITH THE VTI'S THAT GIVE USERS 404 ERRORS
	If InStr(1, objItem, "_vti", 1) = 0 Then
		If Not objItem.Name = "Index.asp" Then
%>
	<tr bgcolor="#eff7ff">
		<td align="left" nowrap><a href="<%=objItem.Name %>" target="_TOP"><%=objItem.Name%></a></td>
		<td align="left" nowrap><%=objItem.DateCreated%></td>
		<td align="left" nowrap><%=objItem.Type%> </td>
	</tr>
<%
		End If
	End If
	Next
'-- NEXT OBJITEM IN THE COLLECTION
'--	NOW THAT THE SUBFOLDERS ARE CREATED, CREATE THE FILES
	For Each objItem In objFolder.Files
		If Not objItem.Name = "Index.asp" Then
		'-- CHANGE THE COLOR OF THE ROW IF DISPLAYING A FILE FOLDER
%>
	<tr>
		<td align="left" nowrap><a href="<%=objItem.Name %>" target="_TOP"><%=objItem.Name%></a></td>
		<td align="left" nowrap><%=objItem.DateCreated %></td>
		<td align="left" nowrap><%=objItem.Type %></td>
	</tr>
<%
		End If
	Next
'-- NEXT OBJITEM IN THE COLLECTION
End If
'--	KILL THE OBJECT VARIABLES
Set objItem = Nothing
Set objFolder = Nothing
Set objFSO = Nothing
%>
</table>
</body>
</html>
```

