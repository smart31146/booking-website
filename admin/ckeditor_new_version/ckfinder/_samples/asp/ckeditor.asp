<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<%
' CKFinder
' ========
' http://ckfinder.com
' Copyright (C) 2007-2010, CKSource - Frederico Knabben. All rights reserved.
'
' The software, this file and its contents are subject to the CKFinder
' License. Please read the license.txt file before using, installing, copying,
' modifying or distribute this file or part of its contents. The contents of
' this file is part of the Source Code of CKFinder.
%>
<% ' You must set "Enable Parent Paths" on your web site in order this relative include to work. %>
<!-- #INCLUDE file="../../ckfinder.asp" -->
<!-- #INCLUDE VIRTUAL="/ckeditor/ckeditor.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<title>CKFinder - Sample - CKEditor</title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta name="robots" content="noindex, nofollow" />
	<link href="../sample.css" rel="stylesheet" type="text/css" />
</head>
<body>
	<h1>
		CKFinder - Sample - CKEditor Integration
	</h1>
	<hr />
	<p>
		CKFinder can be easily integrated with <a href="http://ckeditor.com">CKEditor</a>. Try it now, by clicking
		the "Image" or "Link" icons and then the "<strong>Browse Server</strong>" button.</p>
<%
	Dim oCKEditor, initialValue
	Set oCKEditor = New CKEditor

	oCKEditor.BasePath	= "/ckeditor/"

	initialValue = "<p>Just click the <b>Image</b> or <b>Link</b> button, and then <b>&quot;Browse Server&quot;</b>.</p>"

	' Just call CKFinder::SetupCKEditor before calling editor(), replace() or replaceAll()
	' in CKEditor. The second parameter (optional), is the path for the
	' CKFinder installation (default = "/ckfinder/").
	CKFinder_SetupCKEditor oCKEditor, "../../", empty, empty

	oCKEditor.editor "CKEditor1", initialValue
%>
</body>
</html>
