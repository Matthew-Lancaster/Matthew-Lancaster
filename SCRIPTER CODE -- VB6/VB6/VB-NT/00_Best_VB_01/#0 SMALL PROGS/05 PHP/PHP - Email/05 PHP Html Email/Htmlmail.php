<html> 
<title>My Send Mail</title> 
<body> 
<? 
 $username = "James Smith"; 
 $compname = "La La Saloon"; 
 $compphone = "091-080-2275926"; 
 $compid = 11; 
 $jobtitle = "Hair Stylist"; 
 $email = "james_smith73@yahoo.com"; 
 $Fromname = "La La Salon"; 
 $Fromaddress = "info"; 
 $mailsubject = "You are Selected ! by $compname"; 
 $imgurl = "http://www.geocities.com/james_smith73/lala_pic.jpg"; 
 $template = "http://house1.gomark.com/search/dummy.php"; 
 $image = "<a href='$template'><img src='$imgurl' border=0></a>"; 
 $logo = "<img src='http://www.geocities.com/james_smith73/band3.jpg'>"; 
 $updcv = "http://www.yourphpsite.com/yourfolder/dummy.php"; 

 $msg = "<font face='OCR A Extended' size='2'>Hi $username \n<br>"; 
 $msg = $msg.''."\t We have Received Your CV , and have processed \n<br> "; 
 $msg = $msg.''."the same for the post of $jobtitle in $compname \n<br>"; 
 $msg = $msg.''."<br>"; 
 $msg = $msg.''."\t We are pleased to call you for an Interview kindly \n<br>"; 
 $msg = $msg.''."fix an appointment by calling us on $compphone <br>"; 
 $msg = $msg.''."<br>"; 
 $msg = $msg.''."To Know more about $compname Click <a href='$template'>Here</a><BR>"; 
 $msg = $msg.''."<br>"; 
 $msg = $msg.''."Best Regards<br>"; 
 $msg = $msg.''."<b>Management</b></font>"; 
 $msg = "<table border=0><tr><td>$logo</td></tr></table><table border=0><tr><td>$msg</td><td>$image</td></tr></table>"; 
 $msg = $msg.''."<br><font face='OCR A Extended' size='2'>To Update Your Resume Click  <a href='$updcv'>Here</a></font>"; 

 print ("<PRE>"); 
 if (mail($username." <".$email.">", $mailsubject, $msg, "From: ".$Fromname." <".$Fromaddress.">\nContent-Type: text/html; charset=iso-8859-1")) 
 { 
 		print ("Mail Sent To $username<br><br>"); 
  		print ($msg); 
 } 
 else 
 { 
  		print ("Mail Dead"); 
 } 
 print ("</PRE>"); 
?> 

