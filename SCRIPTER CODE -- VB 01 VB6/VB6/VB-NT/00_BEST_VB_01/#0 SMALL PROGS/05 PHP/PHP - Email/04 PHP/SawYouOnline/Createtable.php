<?

// EDITING STARTS
$dbhost = "localhost";	    // usually this is localhost
$dbname = "sawyou_naj";		// name of the databse
$dbuser = "sawyou_naj";	    // your mySQL username (must have access to database above)
$dbpass = "reloadae231";	// your mySQL password
// EDITING ENDS

// ------------------------------------------------------------------------------------

$db = @mysql_connect("$dbhost","$dbuser","$dbpass");
mysql_select_db("$dbname") or die( "Unable to select database.");

mysql_query("CREATE TABLE usersonline (id int(10) NOT NULL auto_increment, ip varchar(255) NOT NULL default '', timestamp varchar(255) NOT NULL default '', url varchar(100) NOT NULL, PRIMARY KEY  (id))") or die(mysql_error());

print ("Congratulations!! The table 'usersonline' has been created sucessfully. Now please delete this file for security purpose.");

?>