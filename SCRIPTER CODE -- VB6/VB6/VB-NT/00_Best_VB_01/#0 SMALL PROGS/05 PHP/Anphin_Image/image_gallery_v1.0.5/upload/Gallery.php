<?php

/* +----------------------------------------------+
   |               Image Gallery                  |
   |  Written by Darrel Grant October 31, 2003    |
   |   You are free to use or distribute this     |
   |   script however you wish as long as this    |
   |   copyright notice remains intact and says.  |
   |               "Darrel Grant."                |
   +----------------------------------------------+ */

// Define the path to the config file, the templates directory and the text directory.
$webroot = './';

// load config settings
require $webroot.'config.php';

// give me all the gory details
error_reporting(E_ALL);

// general purpose data wrapper
class config
{
	var $vars = '';
	var $starttime = 0;

	function config()
	{
		GLOBAL $cfg;
		$this->vars = &$cfg;
	}

	function timer_start()
	{
		$this->starttime = gettimeofday();
	}

	function timer_stop()
	{
		$endtime = gettimeofday();
		return (float)($endtime['sec'] - $this->starttime['sec']) + ((float)($endtime['usec'] - $this->starttime['usec'])/1000000);
	}
}

// loads templates, sets data vars to be used by display
class dataloader
{
	var $page = 1;
	var $thumbs = array();
	var $photopath = '';
	var $template = '';
	var $templates = '';
	var $header = '';
	var $footer = '';
	var $title = "tier lieder's Image Gallery";
	var $globaltemplates = '';
	var $gallery = '';
	var $action = 'index';
	var $image = '';
	var $imageid = '';
	var $caption = '';
	var $totalthumbs = 0;

	function dataloader()
	{
		GLOBAL $HTTP_GET_VARS,$data;

		$this->globaltemplates = $this->gettemplate('globals');

		$this->page = isset($HTTP_GET_VARS['page']) ? intval($HTTP_GET_VARS['page']) : '1';
		$this->gallery = isset($HTTP_GET_VARS['gallery']) ? $this->validate($HTTP_GET_VARS['gallery']) : 'None';
		$this->imageid = isset($HTTP_GET_VARS['showimage']) ? intval($HTTP_GET_VARS['showimage']) : 'None';

		// show single image
		if (is_numeric($this->imageid)) {
			$this->action = 'showimage';
			$this->page = ceil($this->imageid / $data->vars['thumbs_per_page']);
			if (($this->imageid % $data->vars['thumbs_per_page']) == 0) {
				$this->page++;
			}
			$this->getthumbnails();
			$this->image = $this->thumbs[$this->imageid];
			$this->title = 'Viewing '.$this->image.' in '.str_replace('_', ' ', $this->gallery);
			$this->template = $this->gettemplate('image');
			$this->setcaption();
		}

		// show thumbnails
		if ($this->gallery != 'None' && !is_numeric($this->imageid)) {
			$this->page = isset($HTTP_GET_VARS['page']) ? intval($HTTP_GET_VARS['page']) : $this->page;
			$data->vars['photocellwidth'] = round(100 / $data->vars['thumbs_per_page']);
			$this->template = $this->gettemplate('photo_gallery');
			$this->getthumbnails();
			$this->action = 'gallery';
			$this->title = 'Viewing '.str_replace('_', ' ', $this->gallery);
		}

		// show index
		if ($this->gallery == 'None') {
			$this->action = 'index';
			$this->template = $this->gettemplate('index');
		}

		$this->header = $this->globaltemplates->header( $this->title );
		$this->footer = $this->globaltemplates->footer();
	}

	// get template class
	function gettemplate($templatename)
	{
		GLOBAL $webroot;
		require $webroot."templates/$templatename.php";
		return new $templatename;
	}

	// build array of thumbnails
	function getthumbnails()
	{
		GLOBAL $data;

		$dirhandle = @opendir($data->vars['default_path'].$this->gallery);
		if (!$dirhandle) return FALSE;

		$count = -2;
		$pageend = $this->page * $data->vars['thumbs_per_page'];
		$pagestart = $pageend - $data->vars['thumbs_per_page'];

		$thumbsinpage = 0;
		if ($this->page != 1) {
			$key = ($this->page - 1) * $data->vars['thumbs_per_page'];
		} else {
			$key = 0;
		}
		while (false !== ($file = readdir($dirhandle))) {
			$count++;
			if ($file != '..' && $count >= $pagestart && $count <= $pageend && $thumbsinpage < $data->vars['thumbs_per_page']) {
				$this->thumbs[$key] = $file;
				$key++;
				$thumbsinpage++;
			}
		}
		$this->totalthumbs = $count;
		closedir($dirhandle);
	}

	function setcaption()
	{
		GLOBAL $webroot;
		$imagename = substr($this->image, 0, strlen($this->image) - 4);
		$filename = $webroot.'text/'.$imagename.'.txt';
		$filehandle = @fopen($filename, 'rb');
		if ($filehandle != FALSE) {
			$this->caption = fread($filehandle, filesize($filename));
			fclose($filehandle);
		} else {
			$this->caption = '';
		}
	}

	// makes {tag} replacents, takes array or var
	function parse_tags($var, $template)
	{
		GLOBAL $data;
		if ( count($var) == 1) {
			return str_replace('{'.$var.'}', $data->vars[$var], $template);
		} else {
			foreach ($var as $tag) {
				$template = str_replace('{'.$tag.'}', $data->vars[$tag], $template);
			}
			return $template;
		}
	}

	// trim potentially dangerous characters from GET inputs
	function validate($var)
	{
		$var = str_replace('..', '', $var);
		$var = str_replace('&', '', $var);
		$var = stripslashes($var);
		return $var;
	}
}

// render pages with data retrieved by dataloader
class display
{
	var $output = '';

	// builds list of photos
	function display()
	{
		GLOBAL $dataloader;
		$this->makenav();

		if ($dataloader->action == 'gallery') {
			$this->build_gallery();
		} elseif ($dataloader->action == 'showimage') {
			$this->build_image();
		} elseif ($dataloader->action == 'index') {
			$this->output = $dataloader->template->front_page();
		}
	}

	// builds photo gallery (because you couldnt tell that by reading the name)
	function build_gallery()
	{
		GLOBAL $dataloader,$data;

		$this->output = $dataloader->template->gallerytop($dataloader->gallery);

		$numphotos = count($dataloader->thumbs);
		if ($numphotos > 0) {
			$count = 0;
			if ($dataloader->page != 1) {
				$key = ($dataloader->page - 1) * $data->vars['thumbs_per_page'];
			} else {
				$key = 0;
			}
			foreach ($dataloader->thumbs as $thumbname) {
				$cell = $dataloader->template->thumbcell($data->vars['default_path'].$dataloader->gallery, $thumbname, $dataloader->gallery, $key);
				if ($count==0) {
					$this->output .= '<tr>';
					$this->output .= $cell;
					$count++;
				} else {
					$this->output .= $cell;
					$count++;
				}
				if ($count == $data->vars['thumbs_per_row']) {
					$this->output .= '</tr>';
					$count = 0;
				}
				$key++;
			}
			if ($count != 0) {
				while ($count < $data->vars['thumbs_per_row']) {
				$this->output .= $dataloader->template->emptycell();
					$count++;
				}
				$this->output .= '</tr>';
			}
		} else {
			$this->output .= $dataloader->template->emptygallery();
		}
		$data->vars['pagelinks'] = $this->pagelinks($data->vars['thumbs_per_page'], $dataloader->totalthumbs, $dataloader->page, 'Image');
		$this->output .= $dataloader->template->gallerybottom();
	}

	// paginate thumbnails
	function pagelinks($itemsperpage, $totalitems, $currentpage, $name)
	{
		GLOBAL $PHP_SELF,$HTTP_GET_VARS,$dataloader;

		$temp='';
		$html='';

		foreach($HTTP_GET_VARS as $key=>$value) {
			if($key!='page') {
				$temp.="&$key=$value";
			}
		}

		$offset = $currentpage * $itemsperpage;
		$limit = $itemsperpage;

		if ($currentpage > 1) {
			$prevpage = $currentpage - 1;
			$html .= "<a href=\"$PHP_SELF?$temp&page=$prevpage\" title=\"Previous $limit {$name}s\">&laquo;</a> \n";
		}
		if ($limit > 0) {
			$pages = floor( $totalitems / $limit );
			$curpage = round( $offset / $limit );
			$curpage = $curpage + 1;
		} else {
			$curpage = 1;
			$pages = 1;
		}

		if ($totalitems % $limit)
			$pages++;

		for ($i=1; $i<=$pages; $i++) {
			if ($pages > 1) {
				if ($i == $currentpage) {
					$html .= "<font size=\"+1\"><a href=\"$PHP_SELF?$temp&page=$i\">[$i]</a></font> \n";
				} else {
					$html .= "<a href=\"$PHP_SELF?$temp&page=$i\">$i</a> \n";
				}
			}
		}

		if ($pages != 1 && $currentpage != $pages) {
			$nextpage = $currentpage + 1;
			$html .= "<a href=\"$PHP_SELF?$temp&page=$nextpage\" title=\"Next $limit {$name}s\">&raquo;</a>\n";
		}
		return $html;
	}

	// get links to all galleries and put them in the header
	function makenav()
	{
		GLOBAL $data,$dataloader;

		$dirhandle = opendir('./' . $data->vars['default_path']);

		while ($file = readdir($dirhandle)) {
			if (is_dir('./' . $data->vars['default_path'] . $file) && !preg_match("/\.\.?$/", $file))
				$gallery[] = $file;
		}

		closedir($dirhandle);

		if (empty($gallery)) $gallery = 0;

		if (count($gallery) > 0 && is_array($gallery)) {
			$data->vars['navbits'] = "<form style=\"display: inline;\"><select onChange=\"window.location=('gallery.php'+this.options[this.selectedIndex].value);\"><option value=\"\">Gallery Index</option>";
			foreach ($gallery as $val) {
				$data->vars['navbits'] .= $dataloader->globaltemplates->makenavbit($val);
			}
			$data->vars['navbits'] .= '</select><input type="submit" value="Go" /></form>';
		} else {
			$data->vars['navbits'] = '[ <b>No galleries found</b> ]';
		}

		$dataloader->header = $dataloader->parse_tags('navbits', $dataloader->header);
	}

	function build_image()
	{
		GLOBAL $dataloader,$data;
		$numphotos = count($dataloader->thumbs);
		if ($dataloader->imageid == 0) {
			$prevkey = $dataloader->totalthumbs - 1;
		} else {
			$prevkey = $dataloader->imageid - 1;
		}
		if ($dataloader->imageid == ($dataloader->totalthumbs-1)) {
			$nextkey = 0;
		} else {
			$nextkey = $dataloader->imageid + 1;
		}
		$this->output = $dataloader->template->showimage($data->vars['default_path'].$dataloader->gallery.'/'.$dataloader->image, $dataloader->image, $dataloader->gallery, $dataloader->caption, $prevkey, $nextkey);
	}

	// do footer stuff
	function makefooter()
	{
		GLOBAL $dataloader,$data;

		$dataloader->footer = $dataloader->parse_tags('page_build_time', $dataloader->footer);
	}
}

// create global wrapper
$data = new config();

// start clock
$data->timer_start();

$action = isset($HTTP_GET_VARS['action']) ? $HTTP_GET_VARS['action'] : 'index';

// load data access class
$dataloader = new dataloader;

// load page renderer
$display = new display;

// add http headers
@header('HTTP/1.0 200 OK');
@header('HTTP/1.1 200 OK');
@header('Content-type: text/html');

if ($data->vars['addnocacheheaders']) {
	@header('Expires: Mon, 26 Jul 1997 05:00:00 GMT');
	@header('Cache-Control: no-cache, must-revalidate');
	@header('Pragma: no-cache');
}

// compress pages for spicy flavour
if ($data->vars['gzcompress'])
	ob_start('ob_gzhandler');

// stop clock
$data->vars['page_build_time'] = $data->timer_stop();

// build page
$display->makefooter();

echo $dataloader->header . $display->output . $dataloader->footer;

if ($data->vars['gzcompress'])
	ob_end_flush();

?>