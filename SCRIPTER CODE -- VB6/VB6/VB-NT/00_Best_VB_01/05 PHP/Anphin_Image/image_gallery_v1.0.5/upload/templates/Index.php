<?php

class index
{
	var $gallerynav = '';

	function scandir($dir)
	{
		GLOBAL $data;

		$contents = array();

		if (is_dir($dir)) {
			$dirhandle = opendir($dir);

			while ($file = readdir($dirhandle)) {
				if (is_dir("$dir/$file") && !preg_match("/\.\.?$/", $file)) {
					$contents[] = $file;
				}
			}

			closedir($dirhandle);

			foreach ($contents as $val) {
				$newdir = "$dir/$val";
				if (is_dir($newdir)) {
					$galleryname = strrchr($newdir, '/');
					$galleryname = str_replace('/', '', $galleryname);
					$gallerylink = str_replace($data->vars['default_path'], '', $newdir);
					$this->gallerynav .= "<li><a href=\"gallery.php?gallery=$gallerylink\">$galleryname</a></li>";
					$this->gallerynav .= '<ul>';
					$this->scandir($newdir);
					$this->gallerynav .= '</ul>';
				}
			}
		}
	}

	function front_page()
	{
		GLOBAL $data;

		$this->gallerynav .= '<ul>';
		$this->scandir($data->vars['default_path']);
		$this->gallerynav .= '</ul>';

		return <<<EOF
<center>
<table width="75%" border="0" cellspacing="0" cellpadding="0" class="outertable" bgcolor="{$data->vars['tablebordercolor']}"><tr><td>
<table width="100%" border="0" cellspacing="1" cellpadding="3">
<tr><td bgcolor="{$data->vars['tableheadbgcolor']}">
{$data->vars['largefont']} Anphin Image Gallery {$data->vars['cl']}
</td></tr>
<tr><td bgcolor="{$data->vars['tableheadbgcolor']}" align="center">
{$data->vars['normalfont']}
These images are presented to you on a standalone php script
called Image Gallery.<br>
It is free to download, use and modify.
<br><br>
Use the links below to view the images by category.
{$data->vars['cn']}
<table border="0" width="250" align="center"><tr><td>
{$data->vars['normalfont']} {$this->gallerynav} {$data->vars['cn']}
</td></tr></table>
<br>
</td></tr></table>
</td></tr></table>
</center>
EOF;

	}
}

?>