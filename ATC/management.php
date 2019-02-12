<?php 
//footer page... Unique.
if(FI == md5("F@G!HWP*"))
{
//error_reporting(E_ALL);
?>
<table style="width:100%;">
	<tr>
		<td style="width:160px;"  valign=top><?php 
		OpenTable("Management","98%");?>
									<UL>
										<LI><?php mksubpageILink("management","stats","Statistics");?>
										
										<?php 
										$A = GetUserELevel($_SESSION['username'],$_SESSION['level']);
										//echo $_SESSION['username'].'<br />'.$_SESSION['level'].'<br />'._GLEVEL($_SESSION['username'],'a',md5("]%^&*()OLK!JH'GTY&*(OPL<MD#"),false);
										if($_SESSION['level'] == _GLEVEL($_SESSION['username'],'a',md5("]%^&*()OLK!JH'GTY&*(OPL<MD#"),false))
										{
										?>
										<br /><br />Administration:
										<LI><?php mksubpageILink("management","users","Users");?>
										<LI><?php mksubpageILink("management","articles","Articles");?>
										<?php 
										}
										if($_SESSION['level'] == _GLEVEL($_SESSION['username'],'a',md5("]%^&*()OLK!JH'GTY&*(OPL<MD#"),false) || $_SESSION['level'] == _GLEVEL($_SESSION['username'],'p',md5("]%^&*()OLK!JH'GTY&*(OPL<MD#"),false))
										{
										?>
										<br /><br />Admin:
										<LI><?php mksubpageILink("management","rooms","Edit Rooms");
										
										}
										if($_SESSION['level'] == _GLEVEL($_SESSION['username'],'t',md5("]%^&*()OLK!JH'GTY&*(OPL<MD#"),false))
										{
										?>
										<br /><br />Admin:
										<LI><?php mksubpageILink("management","rooms","Edit Your Rooms");
										
										}?>
									</UL>
								<?php echo SmallSpacer();
	CloseTable();
	?>
		</td>
		<td valign=top>
	<?php 




if(isset($_REQUEST['subpage']))
{
	$data = "";
	$subaction="";
	$prename = "Management - ";
	$subname="";
	
	//echo $_SESSION['level'];
	//echo _GLEVEL($_SESSION['username'],'p',md5("]%^&*()OLK!JH'GTY&*(OPL<MD#"),false);
	
	if($_REQUEST['subpage'] == ILink("bots"))
	{
		$subname = "My Bots";
		$result = mysql_query("SELECT * FROM `bot` WHERE `owner` = '".$user1."'");
		$data.='<div style="overflow:auto;"><table style="width:95%;" cellspacing=0><tr style="background-color:gray;color:white;"><td>Name</td><td style="width:50px;">Mood</td><td style="width:70px;">Actions</td></tr>';
		while($row = mysql_fetch_array($result))
		{
			$data.= '<tr style="background-color:'.$col.';" onmouseover=\'this.style.backgroundColor="'.$color3.'";\' onmouseout=\'this.style.backgroundColor="'.$col.'";\'><td>'.$row['botname'].'</td><td style="width:50px;">'.$row['botmood'].'</td><td style="width:50px;"></td></tr>';
		}
		$data.="</table></div>";
		$subaction='Create AI Bot,97%|<br /><form style="display: inline;" method="post" action="./?page='.ILink("management").'&subpage='.ILink("bots").'">Name: <input type="text" /><br />Seed Memory:<br /><textarea style="width:100%; height:60px;"></textarea><br /> The Seed Memory is use to root the base memory of the AI bot. This will define it\'s characteristics, and persona. If you do not define it\'s age and gender, it will ask you for it... or you have to tell it.<br /><br /><input type="submit" value="Submit" /></form>';
	}
	else if($_REQUEST['subpage'] == ILink("stats")){$subname = "Statistics";}
	//$A = ;
	
	else if($_REQUEST['subpage'] == ILink("rooms") && ($_SESSION['level'] == _GLEVEL($_SESSION['username'],'a',md5("]%^&*()OLK!JH'GTY&*(OPL<MD#"),false) || $_SESSION['level'] == _GLEVEL($_SESSION['username'],'p',md5("]%^&*()OLK!JH'GTY&*(OPL<MD#"),false))){
		if(isset($_REQUEST['AID']))
		{
			$ASSET_POSTED="";
			$DATA_POSTED="";
			$VISUAL_POSTED="";
			$REVISION_POSTED="";
			if(isset($_POST['asset']))
			{
				if($_POST['asset'] == 'attributes')
				{
					$I1 = '0';$I2 = '0';
					if($_POST['inventory'] == 'true') $I1 = '1';if($_POST['dataview'] == 'true') $I2 = '1';
					//Set the revision change data...
					//Figure out what is changed...
					$CHGD = ""; $result2 = mysql_query("SELECT * FROM `rooms` WHERE `room` = '".anti_injection($_REQUEST['AID'])."'");if($row2 = mysql_fetch_array($result2)){
					if($row2['teachers']!=$_POST['teachers'])$CHGD = "`teachers`=".anti_injection($_POST['teachers'])."\n";
					if($row2['roomtype']!=$_POST['roomtype'])$CHGD .= "`roomtype`=".anti_injection($_POST['roomtype'])."\n";
					if($row2['Inventory']!=$I1)$CHGD .= "`Inventory`=".$I1."\n";
					if($row2['DataView']!=$I2)$CHGD .= "`DataView`=".$I2."\n";
					if($row2['roomnotes']!=$_POST['comment'])$CHGD .= "`roomnotes`=".anti_injection($_POST['comment']);
					}
					if($CHGD != "") //there is change! >:|
						mysql_query("INSERT INTO `atc_data`.`rooms_revisions` (`room` ,`changetype` ,`changedata`,`changer`)VALUES ('".anti_injection($_REQUEST['AID'])."', 'room attributes', '$CHGD','".$_SESSION['username']."')");
					
					//Send the changes
					mysql_query("UPDATE `rooms` SET `teachers`='".anti_injection($_POST['teachers'])."', `roomtype`='".anti_injection($_POST['roomtype'])."', `Inventory`=".$I1.", `DataView`=".$I2.", `roomnotes`='".anti_injection($_POST['comment'])."' WHERE `room` = '".anti_injection($_REQUEST['AID'])."'");
					$ASSET_POSTED=" - Changed Successfully!";
				}
				else if($_POST['asset'] == 'dedit')
				{
				
				}
			}
			//else{
				//room ID = token
				$RID = $_REQUEST['AID'];
				$subname = "<a style='color:darkred;' href='?page=".ILink("management")."&subpage=".ILink("rooms")."'>Back</a> - Editing Room ".$RID;
				$prename = "";
				$data.="<table style='width:100%;'><tr><td style='width:450px;'>";
				$result2 = mysql_query("SELECT * FROM `rooms` WHERE `room` = '$RID'");
				if($row2 = mysql_fetch_array($result2))
				{
					//Attributes Box
					$data.='<table style="width:450px;"cellspacing=1><tr><td style="height:0px;">'.OpenTable("Attributes [".$RID."]".$ASSET_POSTED,"400px",true);
					$TID="";if($Ro2 = mysql_fetch_array(mysql_query("SELECT * FROM `teachers` WHERE `id` = '".$row2['teachers']."'")))$TID = $Ro2['teachername']; //Gets teacher's name
					$data.='<form style="display: inline;" style="display: inline;" action="?page='.ILink("management").'&subpage='.ILink("rooms").'&AID='.$RID.'" method="post">Teacher: '.SelBox('teachers','t',$TID,$row2['teachers']).'<br /><input type="hidden" name="asset" value="attributes" />';
					$data.='Room Type: <input type="text" name="roomtype" style="width:200px;" value="'.$row2['roomtype'].'" /><br />'; $TRRY = ""; if($row2['Inventory'] == true) $TRRY="checked";
					$data.='Inventory Complete: <input type="checkbox" value="true" name="inventory" '.$TRRY.' /><br />';$TRRY = ""; if($row2['DataView'] == true) $TRRY="checked";
					$data.='Data View Complete: <input type="checkbox" value="true" name="dataview" '.$TRRY.' /><br />Comment:<br /> <textarea name="comment" style="width:320px">'.$row2['roomnotes'].'</textarea> ';
					$data.='<input type="submit" value="Change" /></form>';
					$data.=CloseTable(true).'</td></tr><tr><td style="height: 100%;" valign=top>';
					
					//Revisions
					$data.=OpenTable("Revision Viewer [".$RID."]".$REVISION_POSTED,"400px",true);
					
					$data.=CloseTable(true);
					
					$data.='</td></tr></table></td><td><table style="width:100%;height:100%;"><tr><td>';
					
					//DataEdit Box (allows the user to edit the devices, and inventory (data system is revisioned (and LOGGED), and
					$data.=OpenTable("Data Edit [".$RID."]".$DATA_POSTED,"90%",true);
					$data.='<form style="display: inline;" style="display: inline;" action="?page='.ILink("management").'&subpage='.ILink("rooms").'&AID='.$RID.'" method="post">';
					$data.='<input type="hidden" name="asset" value="dedit" />';
					$data.='<input type="checkbox" value="true" />Use Room Header';
					$data.='</form>'.CloseTable(true).'</td></tr><tr><td>';
					
					$data.=OpenTable("Visual Edit [".$RID."]".$VISUAL_POSTED,"90%",true);
					
					$data.=''.CloseTable(true).'</td></tr></table></td></tr>';
				}
			//}
			$data.="</table>";
		}
		else{$subname = "Select Room to edit";
		//$data.='List of rooms to choose from:';
			$trList[0]='';$RST=0;
			
			//print_r($trList);
			$YY = 'ORDER BY `room` ASC LIMIT 0 , 200';
			if(substr($Z,0,5) == 'WHERE') $YY = "";
			$result = mysql_query("SELECT * FROM `rooms` $YY $Z");
			$data.= '<div style="height:600px;width:100%;overflow:scroll;background-color:silver;"><table cellspacing=0 style="width:100%;border-width=1px;border-style:solid;"><tr style="background-color:gray;color:white;"><td width=27 style="border-right:1px solid lightblue;">#</td><td width=67 style="border-right:1px solid lightblue;">Room Type</td><td width=200 style="border-right:1px solid lightblue;">Teachers</td><td width=30 style="border-right:1px solid lightblue;">PCs</td><td width=30 style="border-right:1px solid lightblue;">MACs</td><td width=30 style="border-right:1px solid lightblue;">Printers</td><td width=10 style="width:10px;border-right:1px solid lightblue;">Projectors</td><td width=30 style="border-right:1px solid lightblue;">TVs</td></tr>';

			//Stack 3-4 per line
			$I=0;
			while($row = mysql_fetch_array($result))
			{
				$I++;$ins = false;
				
				//$post_id=$row['id'];
				$data .= "<tr style='border:1px solid white;'>".BLK();
				$room_number=$row['room'];
				$teacher=$row['teachers'];
				$type=$row['roomtype'];
				$U1 = "($type)"; if($U1 == "(Class)"){$U1 = "";}
				
				$data.= '<td style="border-right:1px solid white;" width=20>'.'<a href="./?page='.ILink('management').'&subpage='.ILink('rooms').'&AID='.$room_number.'">'.$room_number.'</a>'.'</td><td style="border-right:1px solid white;" width=100>'.$type.'</td>';
				//OpenTable2(subpageILink("ATCDATA","room","Room: $room_number $U"),"230px;");
				if(strstr($teacher,';'))
				{
					$str = explode($teacher,';');
					$data .= "<td style='border-right:1px solid black;'>".BLK();
					foreach($str as $t)
					{
						//Get teacher data... We only need desc, lisd click, and name.
						
						$result3 = mysql_query("SELECT * FROM `teachers` WHERE `id` = '".$t."'");
						if($row3 = mysql_fetch_array($result3))
						{
							$data.= $row3['teachername'];
						}
						
					}
					$data .= '</td>';
				}
				else{
					$data .= "<td style='border-right:1px solid black;'>".BLK();
		//			error_reporting(E_ERROR); 
					try{
					$result2 = mysql_query("SELECT * FROM `teachers` WHERE `id` = $teacher");
					if($row2 = mysql_fetch_array($result2))
					{
						$data .= "<a target='_new' alt='".$row2['teacherdesc']."' class='tlink' href='".$row2['lubbockisd_profile']."'>".$row2['teachername']."</a>";
					}
					}
					catch(Exception $e) {}
					$data .= '</td>';
				}
				
				//Discect the data sting...
				$setm = explode("\n",$row['data']);
				//print_r($setm);
				//echo count($setm);
				//starts with:
				//(x,x): - is a device/unit coord
				//RFormat_Size - Gives us the non-visible dimensions (# of device units)
				//room_format - basic shape of the room
				//est_size - Estimated size of the room...
				$Projectors = 0; $PCs=0; $MACs=0; $Printers=0;$TVs=0;
				foreach($setm as $A)
				{
					//echo substr($A,0,1);
					if(substr($A,0,1) == '(')
					{
						$set2 = explode(':',$A);
						if($set2[1] == 'PC') $PCs++;
						if($set2[1] == 'Projector') $Projectors++;
						if($set2[1] == 'Printer') $Printers++;
						if($set2[1] == 'MAC') $MACs++;
						if($set2[1] == 'TV') $TVs++;
					}
				}
				$data .= "<td width=30 style='border-right:1px solid black;'>$PCs</td><td width=30 style='border-right:1px solid black;'>$MACs</td><td width=30 style='border-right:1px solid black;'>$Printers</td><td width=30 style='border-right:1px solid black;'>$Projectors</td><td width=30>$TVs</td>";
				
				//CloseTable();
				$data.='</tr>';
			
			}
			$data.='<tr style="background-color:gray;color:white;"><td width=27 style="border-right:1px solid lightblue;">#</td><td width=67 style="border-right:1px solid lightblue;">Room Type</td><td width=200 style="border-right:1px solid lightblue;">Teachers</td><td width=30 style="border-right:1px solid lightblue;">PCs</td><td width=30 style="border-right:1px solid lightblue;">MACs</td><td width=30 style="border-right:1px solid lightblue;">Printers</td><td width=10 style="width:10px;border-right:1px solid lightblue;">Projectors</td><td width=30 style="border-right:1px solid lightblue;">TVs</td></tr></table></div>';
		}
	}
	
	else if($_SESSION['level'] == _GLEVEL($_SESSION['username'],'a',md5("]%^&*()OLK!JH'GTY&*(OPL<MD#"),false))
	{
		
		if($_REQUEST['subpage'] == ILink("users")){$subname = "Users";
			$result = mysql_query("SELECT * FROM `account`");
			$data.='<div style="overflow:auto;"><table style="width:100%;" cellspacing=0><tr style="background-color:gray;color:white;"><td>Name</td></tr>';
			$color="lightblue";
			$color2="#CCCCFF";
			$color3="#DDDDFF";
			$is=false;
			while($row = mysql_fetch_array($result))
			{
				$col = "";
				if($is) $col = $color;
				else $col=$color2;
				$data.= '<tr style="background-color:'.$col.';" onmouseover=\'this.style.backgroundColor="'.$color3.'";\' onmouseout=\'this.style.backgroundColor="'.$col.'";\'><td style="border-bottom:1px solid darkblue;">'.$row['name'].'</td></tr>';
			}
			$data.="</table></div>";
		}
		else if($_REQUEST['subpage'] == ILink("rooms")){
			$subname = "Rooms";
		
		}
		else if($_REQUEST['subpage'] == ILink("articles")){
			$subname = "Articles";
			if(isset($_REQUEST['AID']))
			{
			$result = mysql_query("SELECT * FROM `news` WHERE `id` = ".anti_injection($_REQUEST['AID']));
			//Print a select box...
			
			//tinymce("textdata");
			//$data.=' Pick the article: <select name="articlesel" onchange="location.href=\'./?page='.ILink("article").'&subpage='.ILink('edit').'&AID=\'+this.value;"><option value=""></option>';
			//require_once("./header.php");
			if($row = mysql_fetch_array($result))
			{
				if(isset($_POST['textdata']))
				{
					mysql_query("UPDATE `news` SET `data`='".str_replace('\n','<br />',$_POST['textdata'])."' WHERE `id` = ".anti_injection($_REQUEST['AID']));
					header("Location: ./?page=".ILink("management")."&subpage=".ILink('articles'));
					//$data = "Saved.<br />".subpageILink("management","articles","Go toArticles");
				}
				else{
				
				$post_id=$row['id'];
				$post_name=$row['article'];
				$post_author=$row['user'];
				$post_data=$row['data'];
				
							//OpenTable(mksubpageILink_("management","edit","[Back]")." - Edit - ".$file,"97%");
							//$file1=fopen($file,"r");
							$dat1="";
							//TinyMCE("textdata"); //we cannot use it because it will not work...
							//EditArea("textdata");
							
							//$dat1=str_replace("<","&#60;",$dat1);
							//$dat1=str_replace(">","&#62;",$dat1);
							
							//
							$data.= "<form style='display: inline;' method='post' action='./?page=".ILink("management")."&subpage=".ILink('articles')."&AID=".anti_injection($_REQUEST['AID'])."'><textarea name='textdata' id='textdatad' style='width:100%;height:500px;'>".$post_data."</textarea><input type=submit value='X-Mit' />&nbsp;<input type='button' onclick='location.href=\"?/page=".ILink("delfile")."&subpage=".ILink($file)."\";' value='Delete Me' /></form>";
							EditArea("textdatad");
							//CloseTable();
							//TinyMCE("textdatad");
				}
							
			}
			//require_once("./footer.php");
			
			}
			else{
			$result = mysql_query("SELECT * FROM `news` ORDER BY `id` DESC LIMIT 0 ,20");
			//Print a select box...
			
			//tinymce("textdata");
			//$data.='<div style="overflow:auto;"><table style="width:100%;" cellspacing=0><tr style="background-color:gray;color:white;"><td>Name</td></tr>';
			$color="lightblue";
			$color2="#CCCCFF";
			$color3="#DDDDFF";
			$is=false;
			$data.=' Pick the article: <select name="articlesel" onchange="location.href=\'./?page='.ILink("management").'&subpage='.ILink('articles').'&AID=\'+this.value;"><option value=""></option>';
			
			while($row = mysql_fetch_array($result))
			{
				$post_id=$row['id'];
				$post_name=$row['article'];
				$post_author=$row['user'];
				$post_data=$row['data'];
				$col = "";
				if($is) $col = $color;
				else $col=$color2;
				$data.= '<option value="'.$post_id.'">'.$post_name.'</option>';
			}
			$data.="</select>";}
		}
		else if($_REQUEST['subpage'] == ILink("edit")){
			$subname = "Edit";
			$data.="Edit page:<br />";
			//get files in a dir
			if($handle = opendir("."))
			{
				while(false !== ($file = readdir($handle)))
				{
					if(!is_dir($file)) $data.=mksubpageILink_("edit",$file,$file)."<br/>"; //Goes to the specific file in the edit system...
				}
				closedir($handle);
			}
		}
		else{}
		
	}
	else{}
	
	OpenTable2($prename.$subname,"97%");
	echo $data.SmallSpacer();
	CloseTable();
	
	//Subaction - add another table?
	if(strlen($subaction) >= 1)
	{
		$sub=explode('|',$subaction);
		$sub2=explode(',',$sub[0]);
		//echo $sub2;
		$name35435= $sub2[0];
		$size1= $sub2[1];
		OpenTable2($name35435,$size1);
		echo $sub[1]; //all data retained must be echoed.
		CloseTable();
	}
}?>
		</td>
	</tr>
</table>
<?php }?>
