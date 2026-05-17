<?php

class photo_gallery
{
	function gallerytop($galleryname)
	{
		GLOBAL $data;
		$galleryname = str_replace('_', ' ', $galleryname);
		return <<<EOF
<center>
<table width="90%" border="0" cellspacing="0" cellpadding="0" class="outertable" bgcolor="{$data->vars['tablebordercolor']}"><tr><td>
<table width="100%" border="0" cellspacing="1" cellpadding="3"><tr>
<td bgcolor="{$data->vars['tableheadbgcolor']}" colspan="{$data->vars['thumbs_per_row']}" align="center">
{$data->vars['largefont']} $galleryname {$data->vars['cl']}
</td></tr>
EOF;
	}

	function thumbcell($imagepath, $imagename, $gallery, $imageid)
	{
		GLOBAL $data;
		$thumbcaption = '';
		if ($data->vars['override_thumb_size']==1) {
			$dimensions = 'height="'.$data->vars['force_thumb_height'].'" width="'.$data->vars['force_thumb_width'].'"';
		} else {
			$dimensions = '';
		}
		$imgname = str_replace('_', ' ', $imagename);
		$url = "gallery.php?gallery=$gallery&showimage=$imageid";
		if ($data->vars['open_photos_in_new_window']==1)
			$target = '_blank';
		else
			$target = '_self';
		if ($data->vars['use_thumbnails']==1) {
			$thumbdir = 'thumbnails';
		} else {
			$thumbdir = $data->vars['default_path'] . $gallery;
		}
		if (is_dir($data->vars['default_path'] . "$gallery/$imagename")) {
			$url = "gallery.php?gallery=$gallery/$imagename";
			$thumbdir = '.';
			$imagename = $data->vars['default_path'] . 'folder.gif';
			$thumbcaption = '<br>'.$data->vars['smallfont'].$imgname.' '.$data->vars['cs'];
		}
		return <<<EOF

<td bgcolor="{$data->vars['tablecellbgcolor']}" width="{$data->vars['photocellwidth']}%" align="center">
<a href="$url" target="$target"><img src="$thumbdir/$imagename" $dimensions border="0" title="$imgname" /></a>
$thumbcaption
</td>
EOF;
	}

	function emptycell()
	{
		GLOBAL $data;
		return "<td width=\"{$data->vars['photocellwidth']}%\" bgcolor=\"{$data->vars['tablecellbgcolor']}\">&nbsp;</td>";
	}

	function emptygallery()
	{
		GLOBAL $data;
		return "<tr><td align=\"center\" colspan=\"{$data->vars['thumbs_per_row']}\" bgcolor=\"{$data->vars['tablecellbgcolor']}\">{$data->vars['normalfont']} There are no images in this gallery. {$data->vars['cn']}</td></tr>";
	}

	function gallerybottom()
	{
		GLOBAL $data;
		return <<<EOF
</table></td></tr></table>
{$data->vars['normalfont']} {$data->vars['pagelinks']} {$data->vars['cn']}
</center>
EOF;
	}
}

?>