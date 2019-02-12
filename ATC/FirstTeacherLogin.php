<?php
//This ought to be hell for haxers :D
if(FI == md5("F@G!HWP*"))
{
	$result1 = mysql_query("SELECT * FROM `teachers`");
	while($row1 = mysql_fetch_array($result1))
	{
		//echo C5($row1['tname'])."-".$_POST['specificid'].'====';
		//echo C5($row1['teacherpass']). "-".$_POST['routerid'].'<br />';
		
		if(C5($row1['tname']) == $_POST['specificid'])
		if(C5($row1['teacherpass']) == $_POST['routerid'])
		{
			mysql_query("UPDATE `teachers` SET `logged_in`=1, `teacherpass`='".substr(hash("sha256",md5(md5(RandID()."!!t.rGR"."EE"."D").")2r"."K9z".'tez'."5U#%^*&T%"."12K^91L")),0,26)."' WHERE `id`=".$row1['id']." ");
			$R1 = mysql_query("SELECT * FROM `account`");
			$I='';
			
			while($Rr = mysql_fetch_array($R1)) $I=$Rr['id'];
			echo "INSERT INTO `account` (`name`, `rlname`, `pass`, `email`, `admin_level`, `about_me`, `img`, `banned`, `allow_profile_views`, `allow_bot_views`) VALUES ('".anti_injection($_POST['uname'])."', '".$row['teachername']."', '".L5($_POST['tpass'])."', '".$row1['teacheremail']."', '"._GLEVEL($_POST['uname'],'t',md5(']%^&*()OLK!JH\'GTY&*(OPL<MD#'),false)."', '', 'img/user/no_user.png', '0', '1');";
			mysql_query("INSERT INTO `account` (`name`, `rlname`, `pass`, `email`, `admin_level`, `about_me`, `img`, `banned`, `allow_profile_views`) VALUES ('".anti_injection($_POST['uname'])."', '".$row['teachername']."', '".L5($_POST['tpass'])."', '".$row1['teacheremail']."', '"._GLEVEL($_POST['uname'],'t',md5(']%^&*()OLK!JH\'GTY&*(OPL<MD#'),false)."', '', 'img/user/no_user.png', '0', '1');");
			mysql_query("UPDATE `teachers` SET `accountroute`='".A5($I)."' WHERE `id` = ".$row1['id']);
			echo 'Your account '.$_REQUEST['uname'].' has been created.<br/><center><a href="./?page='.ILink("login").'"Login</center>';
			
		}
	}
}
?>