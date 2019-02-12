<?php 
if(FI == md5("F@G!HWP*"))
{

if(isset($_SESSION['username']))
{

if(isset($_POST['newad']))
{

}
else if(isset($_POST['pass']))
{
	
}
else if(isset($_POST['abme']))
{

}
else if(isset($_POST['picture']))
{

}
else if(isset($_POST['security']))
{

}

//get data... (this is executed after set data, so we have already delivered the old)
$user1=$_SESSION['username'];
$result = mysql_query("SELECT * FROM `account` WHERE name = '".$user1."'");
if($row = mysql_fetch_array($result))
{
							
							//OpenTable("Edit Profile", "80%");
								OpenTable2("Security","90%");
								?>
										<form style="display: inline;" name="pass" method="post" action="./?page=<?php echo ILink("profile_edit");?>">
										Old Password: <input type="password" name="oldpass34" /><br />
										New Password: <input type="password" name="newpass34" /><br />
										New Password Again: <input type="password" name="newpass234" /><br />
										<div align=right><input type="submit" value="Submit" /></div>
										</form>
								<?php 
								CloseTable();
								OpenTable2("About Me","90%");
								?>
										<form style="display: inline;" name="abme" method="post" action="./?page=<?php echo ILink("profile_edit");?>">
										<input type="text" value="<?php echo $row['about_me'];?>" name="aboutme2" style="width:100%;height:150px;" />
										<div align=right><input type="submit" value="Submit" /></div>
										</form>
								<?php 
								CloseTable();
								OpenTable2("Picture","90%");
								?>
										 <form style="display: inline;" name="picture" method="post" enctype="multipart/form-data"  action="./?page=<?php echo ILink("profile_edit");?>">
										 <table style="width:100%;">
											<tr>
											<td><input type="file" name="image"><br /><input name="Submit" type="submit" value="Upload image"></td>
											<td align=right><img height=50 width=50 src="./?img=<?php echo ILink($_SESSION['username']);?>" />
											</td></tr>
										 </table>	
										 </form>
								<?php 
								CloseTable();
								OpenTable2("Security","90%");
								?>
										<form style="display: inline;" name="security" method="post" action="./?page=<?php echo ILink("profile_edit");?>">
											<input name="r_view1" <?php if($row['allow_profile_views'])echo'checked=true';?> type="checkbox" /> Allow profile views<br />
											<input name="r_bot1" <?php if($row['allow_bot_views'])echo'checked=true';?> type="checkbox" /> Allow AIBot views<br />
											<div><input name="Submit" type="submit" value="Submit"></div>
										</form>
								<?php CloseTable();?>
								<br />
								<?php mkpageILink("profile_view","View Profile");
							//CloseTable();
								

}
}
}
?>