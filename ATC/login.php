<?php 
//footer page... Unique.
if(FI == md5("F@G!HWP*"))
{


if(isset($_POST['username23']))
{
	
	$user1 = anti_injection($_POST['username23']);
	$pass1 = anti_injection($_POST['pass23']);
	$result = mysql_query("SELECT * FROM `account` WHERE name = '".$user1."'");
	
	if($row = mysql_fetch_array($result))
	{
		//echo L5($pass1).'-'.$row['pass'];
		if(str_replace(' ','',$row['pass']) == L5($pass1))
		{
			$G=$row['admin_level'];
			$Xuser = GetUserLevel($user1,str_replace(' ','',$row['admin_level']));
			//echo $Xuser.' '.$row['admin_level'];
			if($Xuser == GAdmin(md6(FERG.SHLOPPPPPPPING.SHIDDLEFIERRR.ERFDS,'md6-run')) | $Xuser == 'user' | $Xuser == 'teacher' | $Xuser == 'tadmin') //checks if the user has edited his/her user level...
			{
				$_SESSION['username'] = $_POST['username23'];
				//echo '<script>alert("'.$Xuser.'<br />'.GAdmin(md6(FERG.SHLOPPPPPPPING.SHIDDLEFIERRR.ERFDS,'md6-run')).'");</script>';
				//echo $Xuser;
				if($Xuser == 'user')
				{
					$_SESSION['level'] = _GLEVEL($user1,'u',md5("]%^&*()OLK!JH'GTY&*(OPL<MD#"),false);
				}
				else if($Xuser == 'teacher')
				{
					$_SESSION['level'] = _GLEVEL($user1,'t',md5("]%^&*()OLK!JH'GTY&*(OPL<MD#"),false);
				}
				else if($Xuser == 'tadmin')
				{
					$_SESSION['level'] = _GLEVEL($user1,'p',md5("]%^&*()OLK!JH'GTY&*(OPL<MD#"),false);
				}
				else if ($Xuser == GAdmin(md6(FERG.SHLOPPPPPPPING.SHIDDLEFIERRR.ERFDS,'md6-run')))
				{
					$_SESSION['level'] = _GLEVEL($user1,'a',md5("]%^&*()OLK!JH'GTY&*(OPL<MD#"),false);
				}
				$_SESSION['pass'] = md5($row['pass'].md5("f8fgG8a9d==")); //store session cookie, but give them something they don't expect..lolz
				//echo '<script>alert("'.$_SESSION['level'].'");</script>';
				$_SESSION['id'] = $row['id'];
				?><script>location.href="./";</script><?php 
				OpenTable("Login","250px");
?>					Logged in Successfully!
					<script>
						wait(<?php echo $wait_time;?>);
						location.href="./?page=<?php echo ILink("home"); ?>";
					</script><?php 
				CloseTable();
		
			}
			else
			{
				OpenTable("Login","250px");
?>									Unable to Login: Your account might be under maintenance. Please try later.
									<script>
									wait(<?php echo $wait_time;?>);
									location.href="./?page=<?php echo ILink("home"); ?>";
									</script><?php 
				CloseTable();
			}
		}
		else
		{
		
				OpenTable("Login","500px");?>
									Wrong Username or password!<br />If you don't remember your email address, Click <a href="./?page=<?php echo ILink("preset"); ?>">here</a> for a password reset email.<br /><center><?php 
									mkpageILink("login","Back");?></center><?php
				CloseTable();
		}
	}
	else
	{
		$result1 = mysql_query("SELECT * FROM `teachers` WHERE tname = '".$user1."'");
	
		if($row1 = mysql_fetch_array($result1))
		{
			echo $row1['teacherpass'].'--'.L4($pass1);
			if($row1['logged_in'] == 0){ // The teacher has not logged on yet...
			if(str_replace(' ','',$row1['teacherpass']) == L4($pass1))
			{ 	//Anti SQL Injection, Multi-layer password for teacher password security
				//todo: auto-make teacher username... Set Password Reset once the teacher logs on with this user...
				
				//Teacher accepted... now set teacher data...
				//Change password (Security reasons, no subtle use...), and accept teacher...
				
				
				
				OpenTable("First Time","500px");?>
				This is your first time logging in. Please specify your new username, and password:
				<form style="display: inline;" method="post" action="./?page=<?php echo ILink("teacherfirst"); ?>">
				<input type="hidden" name="specificid" value="<?php echo C5($user1);?>" />
				<input type="hidden" name="routerid" value="<?php echo C5($row1['teacherpass']);?>" />
				Username: <input type="text" style="width:120px;" name="uname" value="<?php echo $row1['tname'];?>" /><br />
				Password: <input type="password" id='w1' style="width:120px;" name="tpass" /><br />
				Password (Again): <input type="password" id='w2' style="width:120px;" /><br />
				<input type="submit" value="Create" onclick="if(document.getElementById('w1').value!=document.getElementById('w2').value) {alert('Passwords do not match!');return false;}" />
				</form>
				<?php CloseTable();
				
			}
			else
			{
				OpenTable("Login","250px");
?>									Wrong Username/Password!<br /><?php 
									mkpageILink("login","Back");
				CloseTable();
			}
			}
			else
			{
				OpenTable("Login","250px");
?>									Teacher account has already been created!<br />Log in with your teacher account name.<br /><?php 
									mkpageILink("login","Back");
				CloseTable();
			}
			
		}
		else
		{
				
				OpenTable("Login","250px");
?>									Wrong Username!<br /><?php 
									mkpageILink("login","Back");
				CloseTable();
		}
	}
}
else{
	OpenTable("Login","250px");
?>								<form style="display: inline;" method="post" action="./?page=<?php echo ILink("login");?>">
								Username: <input type="text" name="username23" size="25" /><br />
								Password: <input type="password" name="pass23" size="25" /><br />
								<input type="submit" value="Login" />
								</form><?php 
	CloseTable();
	OpenTable("Notice","250px");
	echo 'If you have not set your account yet, search your email, or email dreadslicer@gmail.com. (You must use your LISD email address)';
	CloseTable();
}
}
?>