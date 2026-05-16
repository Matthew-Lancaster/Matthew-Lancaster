<?

// EDITING STARTS
$dbhost = "localhost";	    // usually this is localhost
$dbname = "sawyou_naj";		// name of the databse
$dbuser = "sawyou_naj";	    // your mySQL username (must have access to database above)
$dbpass = "reloadae231";	// your mySQL password
// EDITING ENDS

// ---------------------------------------------------------------------------------------
//
// SawYouOnline 1.0
// Copyrighted to Najwa @ http://sawyoucry.net
// All rights reserved.
// Created September 10th, 2006
// Published September 12th, 2006
// For any question, please contact najwa@sawyoucry.net
//
// ---------------------------------------------------------------------------------------

$db = @mysql_connect("$dbhost","$dbuser","$dbpass");
mysql_select_db("$dbname") or die( "Unable to select database.");

$ip = $REMOTE_ADDR;
$where = $PHP_SELF;
$timestamp = time();
$timeout = 300;
$noofrows = 0;

$result1 = mysql_query("SELECT * FROM usersonline");
while ($result2 = mysql_fetch_array($result1)) {
	if ($result2[1] == $ip) {
		$noofrows = 1;
	}
}

if ($noofrows == 1) {
	mysql_query("UPDATE usersonline SET timestamp = '$timestamp', url = '$where' WHERE ip = '$ip'");
}

if ($noofrows == 0) {
	mysql_query("INSERT INTO usersonline (ip, timestamp, url) VALUES ('$ip', '$timestamp', '$where')");
}

$alt = $timestamp-$timeout;
mysql_query("DELETE FROM usersonline WHERE timestamp < '$alt'");

$result3 = mysql_query("SELECT DISTINCT ip FROM usersonline");
$online = mysql_numrows($result3);

if ($online == 1) {
	echo "$online user online";
} else {
	echo "$online users online";
}

?>
