<?php

class globals
{
	function header($title)
	{
		GLOBAL $data;

		return <<<EOF
<html>
<head>
<title>$title</title>
<style type="text/css">
BODY {
scrollbar-base-color: #c0d0e0;
scrollbar-arrow-color: #505050;
scrollbar-track-color: #ffffff;
font-family: verdana,arial, sans-serif;
font-size: 11px;
 cursor: default;
 margin: 0px;
 }
INPUT, TEXTAREA, SELECT {
 background: #c0d0e0;
 font-family: arial;
 font-size: 11px;
 color: #505050;
}
A:LINK {
 color: #202020;
 font-weight: bold;
 text-decoration: none;
 cursor: default;
 }
A:VISITED {
 color: #202020;
 font-weight: bold;
 text-decoration: none;
 cursor: default;
 }
A:ACTIVE {
 color: #2070f0;
 font-weight: bold;
 text-decoration: none;
 cursor: default;
 }
A:HOVER {
 color: #2070f0;
 font-weight: bold;
 font-style: normal;
 cursor: default;
 text-decoration: underline;
 }
.outertable {
 border: 1px solid #505050;
 }
</style>
</head>
<body bgcolor="#ffffff" text="#202020" link="#202020" alink="#2070f0" vlink="#505050">
<center>
<table width="75%" border="0" cellpadding="2" cellspacing="6"><tr><td>
{$data->vars['normalfont']} Jump: {navbits} {$data->vars['cn']}
</td><td align="right">
{$data->vars['normalfont']} <a href="http://www.comeplaydying.com">Home</a> {$data->vars['cn']}
</td></tr></table>
</center>
<br>
EOF;
	}

	function footer()
	{
		GLOBAL $data;
		return <<<EOF
<br><br><br>
<center>
Powered by
<!-- Removing this copyright notice violates the license agreement. That is unchristian. -->
<a href="http://www.anphin.com/?page=Image+Gallery">Anphin Image Gallery</a> v{$data->vars['version']}<br>
<!-- The program is free, so just leave my name in it. It won't kill you. The same cannot be said for me. -->
{$data->vars['smallfont']} Page built in: {page_build_time} seconds {$data->vars['cs']}
</center>
</body>
</html>
EOF;
	}

	function makenavbit($gallery)
	{
		GLOBAL $data;
		$galleryname = str_replace('_', ' ', $gallery);
		return "<option value=\"?gallery=$gallery\">$galleryname</option>";

	}
}

?>