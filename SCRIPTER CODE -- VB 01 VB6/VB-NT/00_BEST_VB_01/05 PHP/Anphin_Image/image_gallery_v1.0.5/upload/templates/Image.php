<?php

class image
{
	function showimage($imagepath, $imagename, $gallery, $caption='', $previd, $nextid)
	{
		GLOBAL $data;

		if (is_dir($imagepath)) {
			$caption = "$imagename is a sub gallery. <a href=\"gallery.php?gallery=$gallery/$imagename\">Click here</a> to open it.";
			$imagepath = './folder.gif';
		}

		if ($data->vars['parse_captions']) {
			$caption = $this->parsephpcode($caption);
		}

		$caption = nl2br($caption);

		$galleryname = str_replace('_', ' ', $gallery);
		return <<<EOF
<center>
<table width="90%" border="0" cellspacing="0" cellpadding="0" class="outertable" bgcolor="{$data->vars['tablebordercolor']}"><tr><td>
<table width="100%" border="0" cellspacing="1" cellpadding="3">
<tr><td bgcolor="{$data->vars['tableheadbgcolor']}">
{$data->vars['largefont']} <a href="gallery.php?gallery=$gallery">$galleryname</a> &raquo $imagename {$data->vars['cl']}
</td></tr>
<tr><td bgcolor="{$data->vars['tableheadbgcolor']}" align="center">
<table width="500" border="0" cellpadding="0" cellspacing="0"><tr>
<td>
{$data->vars['normalfont']} <a href="gallery.php?gallery=$gallery&showimage=$previd">&laquo Previous</a> {$data->vars['cn']}
</td>
<td align="right">
{$data->vars['normalfont']} <a href="gallery.php?gallery=$gallery&showimage=$nextid">Next &raquo </a>  {$data->vars['cn']}
</td></tr></table>
</td></tr>
<tr><td colspan="2" bgcolor="{$data->vars['tableheadbgcolor']}" align="center">
{$data->vars['normalfont']} $caption {$data->vars['cn']}
<br>
<img src="$imagepath" border="0" />
</td></tr></table>
</td></tr></table>
</center>
EOF;
	}

	function parsephpcode($text)
	{
		$arr1 = split('<phpcode>', $text);
		$text = $arr1[0];

		foreach($arr1 as $key=>$value) {
			if(eregi('</phpcode>', $value)) {
				$arr2 = split('</phpcode>', $value);
				$arr2[0] = eval($arr2[0]);

				foreach($arr2 as $key2=>$value2) {
					if($key2 == 0) {
						$value2 = "<phpcode>$value2</phpcode>";
					}
					$text .= $value2;
	           }
	       }
	   }
	   return $text;
	}
}


?>