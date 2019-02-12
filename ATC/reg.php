<?php 
//footer page... Unique.
if(FI == md5("F@G!HWP*"))
{



if(isset($_POST['username23']))
{
	$user1 = anti_injection($_POST['username23']);
	$pass1 = anti_injection($_POST['pass23']);
	$pass2 = anti_injection($_POST['pass223']);
	$email1 = anti_injection($_POST['email23']);
	$contains_user = false;
	
	
	$result = mysql_query("SELECT * FROM `account` WHERE name = '".$user1."'");
	if($row = mysql_fetch_array($result))
	{
		$contains_user=true;
	}
	$data="";
	if($contains_user) $data="User Exists!<br /><a href=\"./?page=".ILink("reg")."\">Back</a>";
	else{
		if($pass1 != $pass2) $data="Passwords do not match!<br /><a href=\"./?page=".ILink("reg")."\">Back</a>";
		else
		{
			//Admin level is impossible to edit...without all hell breaking loose.
			mysql_query("INSERT INTO `account` (`id`,`name`,`pass`,`admin_level`,`email`,`about_me`,`img`,`banned`) VALUES (NULL,'$user1','".md5(md5($pass1."&&^.^GFHD")."%UK95U2K2ADDT12K5IL")."','".md5($user1."F8FY1Y2AT5HAG2F"."user")."','$email1','','ATCDataSITE/img/user/nouser.png',0);");
			$data="User $user1 Registered!";
			//echo "INSERT INTO `account` (`id`,`name`,`pass`,`email`,`admin_level`,`banned`) VALUES (NULL,'$user1','".md5(md5($pass1."&&^.^GFHD")."%UK95U2K2ADDT12K5IL")."','".md5($user1."F8FY1Y2AT5HAG2F"."user")."','$email1',0);";
		}
		//insert into database... use anti-sqlinject...
	}
	OpenTable("Register","250px");
	echo $data;
	CloseTable();
}
else{
	OpenTable("Register","250px");
								?><form style="display: inline;" method="post" action="./?page=<?php echo ILink("reg");?>">
								Username: <input type="text" name="username23" size="20" /><br />
								Password: <input type="password" name="pass23" size="20" /><br />
								Repeat: <input type="password" name="pass223" size="23" /><br />
								Email: <input type="text" name="email23" size="25" /><br />
								<input type="submit" value="Register" />
								</form><?php 
	CloseTable();
}
}
?>