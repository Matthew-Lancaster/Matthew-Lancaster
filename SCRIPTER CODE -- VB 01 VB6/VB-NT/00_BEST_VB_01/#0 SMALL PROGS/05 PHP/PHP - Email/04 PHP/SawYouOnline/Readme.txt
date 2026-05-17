-------------------------------------------------------
*
* SawYouOnline 1.0
* Copyrighted to Najwa @ http://sawyoucry.net
* All rights reserved.
* Created September 10th, 2006
* Published September 12th, 2006
* For any question, please contact najwa@sawyoucry.net
*
-------------------------------------------------------

>> SAWYOUONLINE
~~~~~~~~~~~~~~~~~
A user online counter that will keep track of how many users are online and browsing your site at that current time.


>> REQUIREMENTS
~~~~~~~~~~~~~~~~~
1. PHP 3 or higher
2. A mySQL database


>> INSTALLATION
~~~~~~~~~~~~~~~~~
1. Extract the files in this zip file. There are only 3 files.

2. Open createtable.php and sawyouonline.php to edit some variables. You basically only need to configure your 
	- mysql host, name, user and password

3. Upload those 2 files to your server's main directory
	- createtable.php
	- sawyouonline.php

5. Run createtable.php on your browser to create the tables needed.

6. Walla! That's it! And it would be best if you can delete the createtable.php file for security purpose ;)


>> USAGE
~~~~~~~~~~
- You can display the counter by including only ths line of code;

<?php include("/your-site's-absolute-path/sawyouonline.php"); ?>

* Change the "your-site's-absolute-path" part to your site's absolute path.