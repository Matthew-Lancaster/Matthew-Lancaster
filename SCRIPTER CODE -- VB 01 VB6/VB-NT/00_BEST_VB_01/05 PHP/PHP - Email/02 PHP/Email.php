<?php

/***********************************************/
/* Copy and paste this code into the code for
/* your website and enter the email address you
/* wish your statistics to be mailed to in the
/* $email variable.
/***********************************************/

///////////////////////////////////////////////////////////
// Begin tracking code, copy this into your php code
///////////////////////////////////////////////////////////

$email = "email@email.com"; // ENTER YOUR EMAIL ADDRESS HERE
$from = "MatthewLan.com PHP Email";

if ($_SERVER['REMOTE_ADDR'] != "128.223.188.97") {
	if(array_key_exists("HTTP_REFERER",$_SERVER)) {
		$body = $_SERVER['SERVER_NAME'].$_SERVER['REQUEST_URI']."\n".
						date("Y-m-d H:i:s")."\n".
						$_SERVER['REMOTE_ADDR']."\n".
						gethostbyaddr($_SERVER['REMOTE_ADDR'])."\n".
						$_SERVER['HTTP_REFERER'];
	} else {
		$body = $_SERVER['SERVER_NAME'].$_SERVER['REQUEST_URI']."\n".
						date("Y-m-d H:i:s")."\n".
						$_SERVER['REMOTE_ADDR']."\n".
						gethostbyaddr($_SERVER['REMOTE_ADDR']);
	}
	mail($email, $_SERVER['REQUEST_URI'], $body, "From: ".$_SERVER['REMOTE_ADDR']." <".$email.">");
}

///////////////////////////////////////////////////////////
// End tracking code
///////////////////////////////////////////////////////////

?>