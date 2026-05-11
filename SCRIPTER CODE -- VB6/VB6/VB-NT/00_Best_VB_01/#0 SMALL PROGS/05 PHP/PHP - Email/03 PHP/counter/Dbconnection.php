<?php
<?
/*
hi friends
i am back again but this time something most required and wanted by every web site
that is a counter which is easy to use and most informative , plus with every new user 
hitting ur web site u get a confirmation email , just to know that someone is seeing it now
ofcource u can upgrade this program to get visitor ipaddress and country .
to configure please read the "readme instructions" , advantage with this code
is u can plug it to any web site on the net ( javaservlet, jsp, asp , php , perl etc..) , 
if u have access to it , i wrote it in php but u can do the same in any langauge of ur chose
u are comfortable. 

01) Step 1 : I assume that u have access to a webserver running php and mysql database
[please note: refer http://www.php.net to learn about php and http://www.mysql.com both FREE , 
FAST , POWERFULL and BEST in open source environment over 3 million sites on net i read somewhere , 
u can download it and install it easily ]
02) Step 2 : login to MYSQL Database 
03) Step 3 : create a database named "counterdb"
14) Step 4 : Edit dbconnection.php file and Edit these lines
var $IPADD = "localhost";       
//if running on ur pc use 'localhost' else 'ipaddress' 
//ex (194.215.24.223)	
var $USER = "databaseusername"; 
//database username
var $PASS = "*********";        
//database password
var $DBASE = "counterdb";       
//database name
04) Step 4 : create table named "mycount"
05) Step 5 : structure is as follows ::
CREATE TABLE mycount (
   mycntin int(4) NOT NULL auto_increment,
   mycnt int(4) DEFAULT '0' NOT NULL,
   mytime varchar(20),
   PRIMARY KEY (mycntin)
);
06) Step 6 : now make the index.html (home page or start page) file of ur site into frames ,
one of the frames will call this file 
"http://www.yourphpmysqlservername.com/directoryname/top.php"
and other frame will call ur actual site
07) Step 7 : whenever user visits ur site the program "top.php" gets executed and 
   a)) it registers a hit
   b)) it sends Emails to u (Administrator)
08) Step 8 : u can check ur email in ur inbox
09) Step 9 : view counter at  "http://www.yourphpmysqlservername.com/directoryname/counter.php""
10) Step 10 : download source code from this zip file here or from
http://www.geocities.com/james_smith73/counter.zip 
11) Step 11 : view its working on mysite
http://www.geocities.com/james_smith73/
by clicking on the link BEST COUNTER 
*/
class classname
{
	// Connection variables ======================================================
	var $CONN = "";
	var $IPADD = "localhost";       //if running on ur pc use 'localhost' else 'ipaddress'	
	var $USER = "databaseusername"; //database username
	var $PASS = "*********";        //database password
	var $DBASE = "counterdb";       //database name
	//End Connection Variables ====================================================		
	function error($text)
	{
		$no = mysql_errno ();
		$msg = mysql_error ();		
		echo "[$text] ( $no : $msg) <br> \n";
		exit ;
	}

	function init()
	{
		$ipadd = $this->IPADD;
		$user = $this->USER;
		$pass = $this->PASS;
		$dbase = $this->DBASE;
		$conn = mysql_pconnect ($ipadd,$user,$pass);		
		if (!$conn)
		{
			$this->error ("Connection could not be established with server");
		}
		if (!mysql_select_db($dbase))
		{
			$this->error ("Connention with database failed");
		}
		$this->CONN = $conn;
		return true;		
	}
	
}
