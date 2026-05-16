    <?php
    $fileLocation = $HTTP_SERVER_VARS["PHP_SELF"];
    $fileName = basename ($fileLocation);

    $notcomplete = false;
    $nodata = false;

    if ( !isset($_POST["mailto"]) ) {
       $nodata = true;
    }
    if ( !isset($_POST["subject"]) ){
       $nodata = true;
    }
    if ( !isset($_POST["email"]) ){
       $nodata = true;
    }
    if ( !isset($_POST["name"]) ){
       $nodata = true;
    }
    if ( !isset($_POST["message"]) ){
       $nodata = true;
    }

    if ($nodata){
    ?>
    <form action="<?=$fileName?>" method="post">
    <table border="0" cellpadding="0" cellspacing="0"><tr><td>
    Name:<br><input type="text" value="name" name="name" size="20"><br>
    Email:<br><input type="text" value="email" name="email" size="20"><br>
    Subject:<br><input type="text" value="subject" name="subject" size="20"><br>
    Send to:<br><input type="text" value="who@what.com" name="mailto" size="20"><br>
    Message:<br><textarea name="message" cols="30" rows="6"></textarea><br>
    <input type="submit" name="submit" value=" Send "></td></tr></table></form>
    <?php
      exit("");
    }//end if nodata

    if ( trim($_POST["mailto"] == "") ){
       $notcomplete = true;
    }
    if ( trim($_POST["subject"] == "") ){
       $notcomplete = true;
    }
    if ( trim($_POST["email"] == "") ){
       $notcomplete = true;
    }
    if ( trim($_POST["name"] == "") ){
       $notcomplete = true;
    }
    if ( trim($_POST["message"] == "") ){
       $notcomplete = true;
    }

    if ($notcomplete){
         echo "missing required parameter";
         //you could change the a
         exit ('a href="javascript:history.back(1)"> click here a');
    }
    $mailto = $_POST["mailto"];
    $email = $_POST["email"];
    $subject = $_POST["subject"];
    $message = $_POST["message"];
    $name = $_POST["name"];

    //$mailto = "$mailto";
    //$subject = "$subject";
    $msg = ereg_replace("\\\'", "'", $message);
    $msg1 = ereg_replace('\\\"', "\"", $msg);
    $message1 = "From: $name\nemail: $email\nmessage:\n$msg1";
    $sent = mail($mailto, $subject, $msg, "From: $email\r\nReply-to: $email\r\n");
    if ($sent){
         echo ("Your Email has been sent to $mailto!");
    }else{
         echo "Message sending Failed";
    }

    ?>