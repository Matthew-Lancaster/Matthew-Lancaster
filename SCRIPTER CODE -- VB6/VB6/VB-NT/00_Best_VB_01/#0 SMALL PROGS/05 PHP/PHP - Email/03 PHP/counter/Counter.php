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
	//Start Database Connection Parameter
	if (empty($connect) OR $connect!='1')
	{
	 	 	$connect = include ("dbconnection.php");
	}
	$db = new classname;
	$db->init();
	if (!$db->init())
	{
		echo "Connection Failed";
	}
	// End Database Connection	
	if($fromcounter)
	{
		$todaytime= mktime (0,0,0,$month,$day,$year);
		$dtodaytime =date("d",$todaytime);		
		$mtodaytime =date("m",$todaytime);		
		$Ytodaytime =date("Y",$todaytime);	
	}
	else
	{
		$todaytime=time();
		$dtodaytime =date("d",$todaytime);		
		$mtodaytime =date("m",$todaytime);		
		$Ytodaytime =date("Y",$todaytime);	
	}
	//End Connection

class Calendar
{
    function Calendar()
    {
	
    }
    function getDayNames()
    {
        return $this->dayNames;
    }
    function setDayNames($names)
    {
        $this->dayNames = $names;
    }
    function getMonthNames()
    {
        return $this->monthNames;
    }
    function setMonthNames($names)
    {
        $this->monthNames = $names;
    }
    function getStartDay()
    {
        return $this->startDay;
    }
    function setStartDay($day)
    {
        $this->startDay = $day;
    }
    function getStartMonth()
    {
        return $this->startMonth;
    }
    function setStartMonth($month)
    {
        $this->startMonth = $month;
    }
    function getCalendarLink($month, $year)
    {
        return "";
    }
    function getDateLink($day, $month, $year)
    {
		$link = "counter.php?day=$day&month=$month&year=$year&fromcounter=1";
        return "$link";
    }
    function getCurrentMonthView()
    {
        $d = getdate(time());
        return $this->getMonthView($d["mon"], $d["year"]);
    }
    function getCurrentYearView()
    {
        $d = getdate(time());
        return $this->getYearView($d["year"]);
    }
    function getMonthView($month, $year)
    {
        return $this->getMonthHTML($month, $year);
    }
    function getYearView($year)
    {
        return $this->getYearHTML($year);
    }
    function getDaysInMonth($month, $year)
    {
        if ($month < 1 || $month > 12)
        {
            return 0;
        }   
        $d = $this->daysInMonth[$month - 1];   
        if ($month == 2)
        {
            if ($year%4 == 0)
            {
                if ($year%100 == 0)
                {
                    if ($year%400 == 0)
                    {
                        $d = 29;
                    }
                }
                else
                {
                    $d = 29;
                }
            }
        }    
        return $d;
    }
    function getMonthHTML($m, $y, $showYear = 1)
    {
        $s = "";        
        $a = $this->adjustDate($m, $y);
        $month = $a[0];
        $year = $a[1];               
    	$daysInMonth = $this->getDaysInMonth($month, $year);
    	$date = getdate(mktime(12, 0, 0, $month, 1, $year));    	
    	$first = $date["wday"];
    	$monthName = $this->monthNames[$month - 1];    	
    	$prev = $this->adjustDate($month - 1, $year);
    	$next = $this->adjustDate($month + 1, $year);
    	
    	if ($showYear == 1)
    	{
    	    $prevMonth = $this->getCalendarLink($prev[0], $prev[1]);
    	    $nextMonth = $this->getCalendarLink($next[0], $next[1]);
    	}
    	else
    	{
    	    $prevMonth = "";
    	    $nextMonth = "";
    	}
    	
    	$header = $monthName . (($showYear > 0) ? " <font color='orange'><b>" . $year : "</b></font>");
    	
    	$s .= "<table class=\"calendar\" border=4 cellspacing=4 cellpadding=4 align=center bordercolor=blue>\n";
    	$s .= "<tr>\n";
    	$s .= "<td align=\"center\" valign=\"top\">" . (($prevMonth == "") ? "&nbsp;" : "<a href=\"$prevMonth\">&lt;&lt;</a>")  . "</td>\n";
    	$s .= "<td align=\"center\" valign=\"top\" class=\"calendarHeader\" colspan=\"5\">$header</td>\n"; 
    	$s .= "<td align=\"center\" valign=\"top\">" . (($nextMonth == "") ? "&nbsp;" : "<a href=\"$nextMonth\">&gt;&gt;</a>")  . "</td>\n";
    	$s .= "</tr>\n";
    	
    	$s .= "<tr>\n";
    	$s .= "<td align=\"center\"  width=70 class=\"calendarHeader\">" . $this->dayNames[($this->startDay)%7] . "</td>\n";
    	$s .= "<td align=\"center\"  width=70  class=\"calendarHeader\">" . $this->dayNames[($this->startDay+1)%7] . "</td>\n";
    	$s .= "<td align=\"center\"  width=70  class=\"calendarHeader\">" . $this->dayNames[($this->startDay+2)%7] . "</td>\n";
    	$s .= "<td align=\"center\"  width=70  class=\"calendarHeader\">" . $this->dayNames[($this->startDay+3)%7] . "</td>\n";
    	$s .= "<td align=\"center\"  width=70  class=\"calendarHeader\">" . $this->dayNames[($this->startDay+4)%7] . "</td>\n";
    	$s .= "<td align=\"center\"  width=70  class=\"calendarHeader\">" . $this->dayNames[($this->startDay+5)%7] . "</td>\n";
    	$s .= "<td align=\"center\"  width=70  class=\"calendarHeader\">" . $this->dayNames[($this->startDay+6)%7] . "</td>\n";
    	$s .= "</tr>\n";
     	$d = $this->startDay + 1 - $first;
    	while ($d > 1)
    	{
    	    $d -= 7;
    	}
        $today = getdate(time());
    	while ($d <= $daysInMonth)
    	{
    	    $s .= "<tr>\n";       
    	    
    	    for ($i = 0; $i < 7; $i++)
    	    {
        	    $class = ($year == $today["year"] && $month == $today["mon"] && $d == $today["mday"]) ? "calendarToday" : "calendar";
    	        $s .= "<td class=\"$class\" align=\"center\"  >";       
    	        if ($d > 0 && $d <= $daysInMonth)
    	        {
    	            $link = $this->getDateLink($d, $month, $year);
					$mtime=mktime();
					$dnow=date("d",$mtime);
					$mnow=date("m",$mtime);
					if(($dnow==$d) AND ($mnow==$month))
					{
						$s .= (($link == "") ? $d : "<b><a href=\"$link\"><font face='courier new' color='red'>$d</font></a></b>");
					}
					else
					{
						$s .= (($link == "") ? $d : "<a href=\"$link\"><font face='courier new' color='blue'>$d</font></a>");
					}
    	        }
    	        else
    	        {
    	            $s .= "&nbsp;";
    	        }
      	        $s .= "</td>\n";       
        	    $d++;
    	    }
    	    $s .= "</tr>\n";    
    	}
    	
    	$s .= "</table>\n";    	
    	return $s;  	
    }
    
    
    /*
        Generate the HTML for a given year
    */
    function getYearHTML($year)
    {
        $s = "";
    	$prev = $this->getCalendarLink(0, $year - 1);
    	$next = $this->getCalendarLink(0, $year + 1);
        $s .= "<table class=\"calendar\" border=\"1\">\n"; 
        $s .= "<tr>";
    	$s .= "<td align=\"center\" valign=\"top\" align=\"left\">" . (($prev == "") ? "&nbsp;" : "<a href=\"$prev\">&lt;&lt;</a>")  . "</td>\n";
        $s .= "<td class=\"calendarHeader\" valign=\"top\" align=\"center\">" . (($this->startMonth > 1) ? $year . " - " . ($year + 1) : $year) ."</td>\n";
    	$s .= "<td align=\"center\" valign=\"top\" align=\"right\">" . (($next == "") ? "&nbsp;" : "<a href=\"$next\">&gt;&gt;</a>")  . "</td>\n";
        $s .= "</tr>\n";
        $s .= "<tr>";
        $s .= "<td class=\"calendar\" valign=\"top\">" . $this->getMonthHTML(0 + $this->startMonth, $year, 0) ."</td>\n";
        $s .= "<td class=\"calendar\" valign=\"top\">" . $this->getMonthHTML(1 + $this->startMonth, $year, 0) ."</td>\n";
        $s .= "<td class=\"calendar\" valign=\"top\">" . $this->getMonthHTML(2 + $this->startMonth, $year, 0) ."</td>\n";
        $s .= "</tr>\n";
        $s .= "<tr>\n";
        $s .= "<td class=\"calendar\" valign=\"top\">" . $this->getMonthHTML(3 + $this->startMonth, $year, 0) ."</td>\n";
        $s .= "<td class=\"calendar\" valign=\"top\">" . $this->getMonthHTML(4 + $this->startMonth, $year, 0) ."</td>\n";
        $s .= "<td class=\"calendar\" valign=\"top\">" . $this->getMonthHTML(5 + $this->startMonth, $year, 0) ."</td>\n";
        $s .= "</tr>\n";
        $s .= "<tr>\n";
        $s .= "<td class=\"calendar\" valign=\"top\">" . $this->getMonthHTML(6 + $this->startMonth, $year, 0) ."</td>\n";
        $s .= "<td class=\"calendar\" valign=\"top\">" . $this->getMonthHTML(7 + $this->startMonth, $year, 0) ."</td>\n";
        $s .= "<td class=\"calendar\" valign=\"top\">" . $this->getMonthHTML(8 + $this->startMonth, $year, 0) ."</td>\n";
        $s .= "</tr>\n";
        $s .= "<tr>\n";
        $s .= "<td class=\"calendar\" valign=\"top\">" . $this->getMonthHTML(9 + $this->startMonth, $year, 0) ."</td>\n";
        $s .= "<td class=\"calendar\" valign=\"top\">" . $this->getMonthHTML(10 + $this->startMonth, $year, 0) ."</td>\n";
        $s .= "<td class=\"calendar\" valign=\"top\">" . $this->getMonthHTML(11 + $this->startMonth, $year, 0) ."</td>\n";
        $s .= "</tr>\n";
        $s .= "</table>\n";        
        return $s;
    }
    function adjustDate($month, $year)
    {
        $a = array();  
        $a[0] = $month;
        $a[1] = $year;
        
        while ($a[0] > 12)
        {
            $a[0] -= 12;
            $a[1]++;
        }
        
        while ($a[0] <= 0)
        {
            $a[0] += 12;
            $a[1]--;
        }
        
        return $a;
    }
    var $startDay = 0;
    var $startMonth = 1;
    var $dayNames = array("<font color='red' face='courier new'>Sunday</font>", "<font color='green' face='courier new'>Monday", "<font color='green' face='courier new'>Tuesday", "<font color='green' face='courier new'>Wednesday", "<font color='green' face='courier new'>Thursday", "<font color='green' face='courier new'>Friday", "<font color='red' face='courier new'>Saturday");
    var $monthNames = array("<font color='maroon' size=4 face='courier new'><b>January</b></font>", 
	                        "<font color='maroon' size=4 face='courier new'><b>February</b></font>", 
							"<font color='maroon' size=4 face='courier new'><b>March</b></font>", 
							"<font color='maroon' size=4 face='courier new'><b>April</b></font>", 
							"<font color='maroon' size=4 face='courier new'><b>May</b></font>", 
							"<font color='maroon' size=4 face='courier new'><b>June</b></font>",
                            "<font color='maroon' size=4 face='courier new'><b>July</b></font>", 
							"<font color='maroon' size=4 face='courier new'><b>August</b></font>", 
							"<font color='maroon' size=4 face='courier new'><b>September</b></font>", 
							"<font color='maroon' size=4 face='courier new'><b>October</b></font>", 
							"<font color='maroon' size=4 face='courier new'><b>November</b></font>", 
							"<font color='maroon' size=4 face='courier new'><b>December</b></font>");
    var $daysInMonth = array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);
    
}
class MyCalendar extends Calendar
{
    function getCalendarLink($month, $year)
    {
        $s = getenv('SCRIPT_NAME');
        return "$s?month=$month&year=$year";
    }
}
$d = getdate(time());
if ($month == "")
{
    $month = $d["mon"];
}

if ($year == "")
{
    $year = $d["year"];
}
$cal = new MyCalendar;
echo $cal->getMonthView($month, $year);
print("<br>");
$qry="SELECT * FROM mycount";
$result=mysql_query($qry);
$num=mysql_num_rows($result);
$count=0;
echo "
	<table border=4 cellspacing=4 cellpadding=4 align=center bordercolor=maroon>		
		<tr>
			<td><font face='courier new' color='red'>Sl No</td><td><font face='courier new' color='red'>Date</td><td><font face='courier new' color='red'>Time</td>
		</tr> ";
if($row=mysql_fetch_array($result))
{
	do
	{
		$mycnt=$row['mycnt'];
		$mytime=$row['mytime'];
		$dshowtime = date("d",$mytime);		
		$mshowtime = date("m",$mytime);		
		$Yshowtime = date("Y",$mytime);		
		$showtimenow1 = date("D dS, F  Y ",$mytime);		
		$showtimenow2 = date("g:i a",$mytime);		
		if(($dshowtime==$dtodaytime) AND ($mshowtime==$mtodaytime) AND ($Yshowtime==$Ytodaytime))
		{
			echo "
			<tr>
				<td><font face='courier new' color='blue'>Hit Number $mycnt</td>
				<td><font face='courier new' color='blue'>$showtimenow1</td>
				<td><font face='courier new' color='blue'>$showtimenow2</td>
			</tr>
			";
			$count++;
		}
		
	}
	while($row=mysql_fetch_array($result));
}
echo "	
		<tr>
			<td></td><td><font face='courier new' color='green'>Number of hits Today : <b>$count</b></td>
		</tr>
		<tr>
			<td></td><td><font face='courier new' color='green'>Number of hits Totally : <b>$num</b></td>
		</tr>
	</table>
";
?> 


