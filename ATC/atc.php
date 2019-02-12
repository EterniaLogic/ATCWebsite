<?php  
//footer page... Unique.
function DRWButtons($layout)
{
	?><table><tr><?php 
	$R1 = false;$R2 = false;$R3 = false;
	if($layout == 0) echo '<td style="border-style:solid;border-width:2px;font-size:12px;">Compact View</td>'; else echo '<td><input type=button onclick="location.href=\'./?page='.ILink('ATCDATA').'&subpage='.ILink('chview0').'\'" value="Compact View" /></td>';
	if($layout == 1) echo '<td style="border-style:solid;border-width:2px;font-size:12px;">List View</td>'; else echo '<td><input type=button onclick="location.href=\'./?page='.ILink('ATCDATA').'&subpage='.ILink('chview1').'\'" value="List View" /></td>';
	if($layout == 2) echo '<td style="border-style:solid;border-width:2px;font-size:12px;">Map View</td>'; else echo '<td><input type=button onclick="location.href=\'./?page='.ILink('ATCDATA').'&subpage='.ILink('chview2').'\'" value="Map View" /></td>';
	
	?></tr></table><br /><?php 
}
function DRWSearch()
{
	echo '<form style="display: inline;" method="post" action="./?page='.ILink('ATCDATA').'">';?>
	<input type="text" name="searchy" style="width:300px;" value="<?php if(isset($_POST['searchy'])) echo $_POST['searchy'];?>" />
	<input type="submit" value="Search" /><br/>
	<input type="radio" <?php if(isset($_POST['searchby'])) {if($_POST['searchby'] == 'room') {echo 'checked=true';}} else echo 'checked=true';?> name="searchby" value="room" />By Room 
	<input type="radio"<?php if(isset($_POST['searchby'])) {if($_POST['searchby'] == 'teacher') {echo 'checked=true';}}?> name="searchby" value="teacher" />By Teacher 
	<input type="radio" name="searchby" <?php if(isset($_POST['searchby'])) {if($_POST['searchby'] == 'type') {echo 'checked=true';}}?> value="type" />By Service Tag
	<input type="hidden" name="searchit" />	
	</form>
	<?php 
}
function DRWList($Z="",$searchtype=-1)
{
?>List of rooms to choose from:<?php 
	$trList[0]='';$RST=0;
	if($searchtype == 1)
	{//$Z is out output... even if it was our input.
		$result2 = mysql_query("SELECT * FROM `teachers` WHERE `teachername` LIKE '%".$_POST['searchy']."%'");
		while($row2 = mysql_fetch_array($result2))
		{
			$Z = 'WHERE `teachers` = '.$row2['id'];
		}
	}
	else if($searchtype == 2) // Awesomeness :D
	{
		$result2 = mysql_query("SELECT * FROM `rooms`");
		while($row2 = mysql_fetch_array($result2))
		{
			//find the rooms that contain these service-tags...
			$setm = explode("\n",$row2['data']);
			foreach($setm as $A)
			{
				if(substr($A,0,1) == '(')
				{
					$s = explode(':',$A);
					//$s[0] == coords
					//$s[1] == type
					//$s[2] == model
					//$s[3] == Serial
					//$s[4] == comment
					
					if(strstr($s[3],$_POST['searchy']))
					{
						$trList[$RST] = $row2['room']; $RST++;
					}
				}
			}
			//$Z = 'WHERE `teachers` = '.$row2['id'];
		}
	}
	
	//print_r($trList);
	$YY = 'ORDER BY `room` ASC LIMIT 0 , 200';
	if(substr($Z,0,5) == 'WHERE') $YY = "";
	$result = mysql_query("SELECT * FROM `rooms` $YY $Z");
	?><div style="height:600px;width:100%;overflow:scroll;background-color:silver;"><table cellspacing=0 style="width:100%;border-width=1px;border-style:solid;"><tr style="background-color:gray;color:white;"><td width=27 style="border-right:1px solid lightblue;">#</td><td width=67 style="border-right:1px solid lightblue;">Room Type</td><td width=200 style="border-right:1px solid lightblue;">Teachers</td><td width=30 style="border-right:1px solid lightblue;">PCs</td><td width=30 style="border-right:1px solid lightblue;">MACs</td><td width=30 style="border-right:1px solid lightblue;">Printers</td><td width=10 style="width:10px;border-right:1px solid lightblue;">Projectors</td><td width=30 style="border-right:1px solid lightblue;">TVs</td></tr><?php 

	//Stack 3-4 per line
	$I=0;
	while($row = mysql_fetch_array($result))
	{
		$I++;$ins = false;
		if($searchtype == 2) 
		{
			
			foreach($trList as $T)
			{
				if($T == $row['room'])
				{
					$ins=true;
				}
			}
			//if($ins==false) {break;}
		}
		if($searchtype == 2 && $ins == false) continue;
		
		//$post_id=$row['id'];
		?><tr style="border:1px solid white;"><?php 
		$room_number=$row['room'];
		$teacher=$row['teachers'];
		$type=$row['roomtype'];
		$U = "($type)"; if($U == "(Class)"){$U = "";}
		
		echo '<td style="border-right:1px solid white;" width=20>'.'<a href="./?page='.ILink('ATCDATA').'&class='.$room_number.'">'.$room_number.'</a>'.'</td><td style="border-right:1px solid white;" width=100>'.$type.'</td>';
		//OpenTable2(subpageILink("ATCDATA","room","Room: $room_number $U"),"230px;");
		if(strstr($teacher,';'))
		{
			$str = explode($teacher,';');
			?><td style="border-right:1px solid black;"><?php 
			foreach($str as $t)
			{
				//Get teacher data... We only need desc, lisd click, and name.
				
				$result2 = mysql_query("SELECT * FROM `teachers` WHERE `id` = $t");
				if($row2 = mysql_fetch_array($result2))
				{
					echo "<a target='_new' alt='".$row2['teacherdesc']."' class='tlink' href='".$row2['lubbockisd_profile']."'>".$row2['teachername']."</a> ";
				}
				
			}
			?></td><?php 
		}
		else{
			?><td style='border-right:1px solid black;'><?php echo BLK();
//			error_reporting(E_ERROR); 
			try{
			$result2 = mysql_query("SELECT * FROM `teachers` WHERE `id` = $teacher");
			if($row2 = mysql_fetch_array($result2))
			{
				echo "<a target='_new' alt='".$row2['teacherdesc']."' class='tlink' href='".$row2['lubbockisd_profile']."'>".$row2['teachername']."</a>";
			}
			}
			catch(Exception $e) {}
			?></td><?php 
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
		echo "<td width=30 style='border-right:1px solid black;'>$PCs</td><td width=30 style='border-right:1px solid black;'>$MACs</td><td width=30 style='border-right:1px solid black;'>$Printers</td><td width=30 style='border-right:1px solid black;'>$Projectors</td><td width=30>$TVs</td>";
		
		//CloseTable();
		?></tr><?php 
	
	}
	?><tr style="background-color:gray;color:white;"><td width=27 style="border-right:1px solid lightblue;">#</td><td width=67 style="border-right:1px solid lightblue;">Room Type</td><td width=200 style="border-right:1px solid lightblue;">Teachers</td><td width=30 style="border-right:1px solid lightblue;">PCs</td><td width=30 style="border-right:1px solid lightblue;">MACs</td><td width=30 style="border-right:1px solid lightblue;">Printers</td><td width=10 style="width:10px;border-right:1px solid lightblue;">Projectors</td><td width=30 style="border-right:1px solid lightblue;">TVs</td></tr></table></div><?php 
}
function DRWMap()
{
	?>
	<img src="facilitymap.gif" width="550" height="340" border="0" usemap="#map" />

<map name="map">
<!-- #$-:Image map file created by GIMP Image Map plug-in -->
<!-- #$-:GIMP Image Map plug-in by Maurits Rijk -->
<!-- #$-:Please do not edit lines starting with "#$" -->
<!-- #$VERSION:2.3 -->
<!-- #$AUTHOR:dread -->
<area shape="poly" coords="39,32,103,30,103,106,59,106,60,75,38,74" alt="124 Construction LAB" href="./?page=<?php echo ILink('ATCDATA');?>&class=124" />
<area shape="poly" coords="143,85,150,63,147,32,220,33,218,50,237,71,235,105,142,106" alt="138 Metals LAB" href="./?page=<?php echo ILink('ATCDATA');?>&class=138" />
<area shape="poly" coords="267,106,266,29,387,30,385,105" alt="140 Automotive" href="./?page=<?php echo ILink('ATCDATA');?>&class=140" />
<area shape="poly" coords="428,33,426,60,401,65,399,90,451,93,452,106,535,105,537,33" alt="146 Automotive (East)" href="./?page=<?php echo ILink('ATCDATA');?>&class=146" />
<area shape="poly" coords="413,128,408,243,447,242,449,259,530,259,535,106,452,107,453,118,413,119" alt="152 Automotive (West)" href="./?page=<?php echo ILink('ATCDATA');?>&class=152" />
<area shape="rect" coords="398,106,427,119" alt="152a Automotive office" href="./?page=<?php echo ILink('ATCDATA');?>&class=152a" />
<area shape="rect" coords="436,106,451,119" alt="152c Automotive Office" href="./?page=<?php echo ILink('ATCDATA');?>&class=152c" />
<area shape="poly" coords="267,207,267,233,281,233,283,245,297,245,297,233,350,232,350,244,367,244,368,225,383,223,385,207" alt="158 Computer Access" href="./?page=<?php echo ILink('ATCDATA');?>&class=158" />
<area shape="poly" coords="145,207,143,259,190,261,195,235,191,208" alt="101 Confrence" href="./?page=<?php echo ILink('ATCDATA');?>&class=101" />
<area shape="rect" coords="354,158,385,193" href="./?page=<?php echo ILink('ATCDATA');?>&class=159" />
<area shape="rect" coords="309,156,340,193" href="./?page=<?php echo ILink('ATCDATA');?>&class=161" />
<area shape="rect" coords="263,156,296,193" href="./?page=<?php echo ILink('ATCDATA');?>&class=162" />
<area shape="rect" coords="231,156,263,193" href="./?page=<?php echo ILink('ATCDATA');?>&class=163" />
<area shape="rect" coords="201,156,230,193" href="./?page=<?php echo ILink('ATCDATA');?>&class=164" />
<area shape="rect" coords="242,120,267,157" href="./?page=<?php echo ILink('ATCDATA');?>&class=141" />
<area shape="rect" coords="269,121,294,157" href="./?page=<?php echo ILink('ATCDATA');?>&class=142" />
<area shape="rect" coords="293,120,320,156" href="./?page=<?php echo ILink('ATCDATA');?>&class=143" />
<area shape="rect" coords="321,120,347,156" href="./?page=<?php echo ILink('ATCDATA');?>&class=144" />
<area shape="rect" coords="346,120,372,157" href="./?page=<?php echo ILink('ATCDATA');?>&class=145" />
<area shape="rect" coords="372,120,387,131" href="./?page=<?php echo ILink('ATCDATA');?>&class=149" />
<area shape="rect" coords="373,132,387,143" href="./?page=<?php echo ILink('ATCDATA');?>&class=150" />
<area shape="rect" coords="372,144,386,157" href="./?page=<?php echo ILink('ATCDATA');?>&class=151" />
<area shape="rect" coords="397,29,425,62" href="./?page=<?php echo ILink('ATCDATA');?>&class=147" />
<area shape="rect" coords="249,83,265,105" href="./?page=<?php echo ILink('ATCDATA');?>&class=139" />
<area shape="rect" coords="116,84,141,105" href="./?page=<?php echo ILink('ATCDATA');?>&class=125" />
<area shape="rect" coords="35,74,60,106" href="./?page=<?php echo ILink('ATCDATA');?>&class=123" />
<area shape="rect" coords="35,123,61,158" href="./?page=<?php echo ILink('ATCDATA');?>&class=122" />
<area shape="rect" coords="72,122,98,159" href="./?page=<?php echo ILink('ATCDATA');?>&class=126" />
<area shape="rect" coords="99,122,125,158" href="./?page=<?php echo ILink('ATCDATA');?>&class=127" />
<area shape="rect" coords="125,121,152,158" href="./?page=<?php echo ILink('ATCDATA');?>&class=128" />
<area shape="rect" coords="153,121,179,159" href="./?page=<?php echo ILink('ATCDATA');?>&class=129" />
<area shape="rect" coords="82,159,113,193" href="./?page=<?php echo ILink('ATCDATA');?>&class=104" />
<area shape="rect" coords="115,158,143,193" href="./?page=<?php echo ILink('ATCDATA');?>&class=103" />
<area shape="rect" coords="143,158,174,193" href="./?page=<?php echo ILink('ATCDATA');?>&class=102" />
<area shape="rect" coords="34,169,60,208" href="./?page=<?php echo ILink('ATCDATA');?>&class=117" />
<area shape="rect" coords="70,206,84,220" href="./?page=<?php echo ILink('ATCDATA');?>&class=110" />
<area shape="rect" coords="70,220,83,233" href="./?page=<?php echo ILink('ATCDATA');?>&class=111" />
<area shape="rect" coords="69,234,83,245" href="./?page=<?php echo ILink('ATCDATA');?>&class=112" />
<area shape="rect" coords="69,245,84,260" href="./?page=<?php echo ILink('ATCDATA');?>&class=113" />
<area shape="rect" coords="32,234,59,247" href="./?page=<?php echo ILink('ATCDATA');?>&class=111" />
<area shape="rect" coords="86,219,118,260" href="./?page=<?php echo ILink('ATCDATA');?>&class=107" />
<area shape="rect" coords="119,220,143,260" href="./?page=<?php echo ILink('ATCDATA');?>&class=106" />
<area shape="rect" coords="129,209,144,220" href="./?page=<?php echo ILink('ATCDATA');?>&class=105" />
<area shape="rect" coords="84,207,96,220" href="./?page=<?php echo ILink('ATCDATA');?>&class=109" />
<area shape="rect" coords="98,207,110,221" href="./?page=<?php echo ILink('ATCDATA');?>&class=108" />
<area shape="rect" coords="369,244,384,261" href="./?page=<?php echo ILink('ATCDATA');?>&class=157" />
<area shape="rect" coords="261,244,273,260" href="./?page=<?php echo ILink('ATCDATA');?>&class=165" />
<area shape="rect" coords="238,245,249,260" href="./?page=<?php echo ILink('ATCDATA');?>&class=169" />
<area shape="rect" coords="249,244,260,260" href="./?page=<?php echo ILink('ATCDATA');?>&class=167" />
<area shape="rect" coords="397,242,423,261" href="./?page=<?php echo ILink('ATCDATA');?>&class=152d" />
<area shape="rect" coords="424,243,434,259" href="./?page=<?php echo ILink('ATCDATA');?>&class=152e" />
<area shape="rect" coords="435,243,447,260" href="./?page=<?php echo ILink('ATCDATA');?>&class=152f" />
<area shape="rect" coords="396,202,409,234" href="./?page=<?php echo ILink('ATCDATA');?>&class=156" />
<area shape="rect" coords="427,105,436,117" href="./?page=<?php echo ILink('ATCDATA');?>&class=152b" />
<area shape="rect" coords="428,92,437,106" href="./?page=<?php echo ILink('ATCDATA');?>&class=146b" />
<area shape="rect" coords="398,92,427,105" href="./?page=<?php echo ILink('ATCDATA');?>&class=146a" />
<area shape="rect" coords="437,93,451,104" href="./?page=<?php echo ILink('ATCDATA');?>&class=146c" />
<area shape="rect" coords="357,244,369,260" href="./?page=<?php echo ILink('ATCDATA');?>&class=158f" />
<area shape="rect" coords="282,245,357,260" alt="158e - Teachers Lounge" href="./?page=<?php echo ILink('ATCDATA');?>&class=158e" />
<area shape="rect" coords="298,233,310,244" href="./?page=<?php echo ILink('ATCDATA');?>&class=158a" />
<area shape="rect" coords="310,233,323,244" href="./?page=<?php echo ILink('ATCDATA');?>&class=158b" />
<area shape="rect" coords="323,232,337,244" href="./?page=<?php echo ILink('ATCDATA');?>&class=158c" />
<area shape="rect" coords="336,232,350,244" href="./?page=<?php echo ILink('ATCDATA');?>&class=158d" />
<area shape="rect" coords="275,244,281,260" href="./?page=<?php echo ILink('ATCDATA');?>&class=158g" />
<area shape="rect" coords="31,245,47,260" href="./?page=<?php echo ILink('ATCDATA');?>&class=111b" />
<area shape="rect" coords="47,245,60,261" href="./?page=<?php echo ILink('ATCDATA');?>&class=111a" />
<area shape="rect" coords="105,96,116,106" href="./?page=<?php echo ILink('ATCDATA');?>&class=124a" />
<area shape="rect" coords="104,78,116,95" href="./?page=<?php echo ILink('ATCDATA');?>&class=124b" />
<area shape="rect" coords="104,67,129,77" href="./?page=<?php echo ILink('ATCDATA');?>&class=124c,138c" />
<area shape="rect" coords="104,57,129,67" href="./?page=<?php echo ILink('ATCDATA');?>&class=124d" />
<area shape="rect" coords="129,56,147,64" href="./?page=<?php echo ILink('ATCDATA');?>&class=138f" />
<area shape="rect" coords="130,45,147,56" href="./?page=<?php echo ILink('ATCDATA');?>&class=138e" />
<area shape="rect" coords="129,31,147,45" href="./?page=<?php echo ILink('ATCDATA');?>&class=138d" />
<area shape="rect" coords="105,30,129,46" href="./?page=<?php echo ILink('ATCDATA');?>&class=124f" />
<area shape="rect" coords="105,45,129,57" href="./?page=<?php echo ILink('ATCDATA');?>&class=124e" />
<area shape="rect" coords="236,70,248,91" href="./?page=<?php echo ILink('ATCDATA');?>&class=138b" />
<area shape="rect" coords="236,91,248,105" href="./?page=<?php echo ILink('ATCDATA');?>&class=138a" />
</map>
	<?php 
}
function DRWCompact()
{
	?>List of rooms to choose from:<?php 
	
	$result = mysql_query("SELECT * FROM `rooms` ORDER BY `room` ASC LIMIT 0 , 200");
	?><div style="height:600px;width:100%;overflow:scroll;background-color:silver;"><table cell-spacing=0 style="width:100%;"><tr><td></td><?php 

	//Stack 3-4 per line
	$I=0;$R=0;$RMAX=3;
	while($row = mysql_fetch_array($result))
	{
		$I++;
		if($RMAX == $R) {$R = 0;?><td></td></tr><tr><td></td><?php }$R++;
		//$post_id=$row['id'];
		?><td><?php 
		$room_number=$row['room'];
		$teacher=$row['teachers'];
		$data=$row['data'];
		$type=$row['roomtype'];
		$U = "($type)"; if($U == "(Class)"){$U = "";}
		OpenTable2('<a href="./?page='.ILink('ATCDATA').'&class='.$room_number.'" >Room: '.$room_number.' '.$U.'</a>',"230px;");
		if(strstr($teacher,';'))
		{
			$str = explode($teacher,';');
			foreach($str as $t)
			{
				//Get teacher data... We only need desc, lisd click, and name.
				$result2 = mysql_query("SELECT * FROM `teachers` WHERE `id` = $t");
				if($row2 = mysql_fetch_array($result2))
				{
					echo "<a target='_new' class='tlink' href='".$row2['lubbockisd_profile']."'>".$row2['teachername']."</a>";
				}
			}
		}
		else{
			//error_reporting(E_ERROR);
			$result2 = mysql_query("SELECT * FROM `teachers` WHERE `id` = $teacher");
			if($row2 = mysql_fetch_array($result2))
			{
				echo "<a target='_new' class='tlink' href='".$row2['lubbockisd_profile']."'>".$row2['teachername']."</a>";
			}
		}
		CloseTable();
		?></td><?php 
	
	}
	?></table></div><?php 
}
function hex2bin($str) {
    $bin = "";
    $i = 0;
    do {
        $bin .= chr(hexdec($str{$i}.$str{($i + 1)}));
        $i += 2;
    } while ($i < strlen($str));
    return $bin;
}
$GR="";
if(FI == md5("F@G!HWP*"))
{
	if(isset($_REQUEST['subpage'])){
	if($_REQUEST['subpage'] == ILink('chview0')) {setcookie("ELAYOUT", 0);}
	if($_REQUEST['subpage'] == ILink('chview1')) {setcookie("ELAYOUT", 1);}
	if($_REQUEST['subpage'] == ILink('chview2')) {setcookie("ELAYOUT", 2);}
	//header("Location: ./?page=".ILink('ATCDATA'));
	?><script>location.href="./?page=<?php  echo ILink("ATCDATA"); ?>";</script><?php 
	}
	
	if(isset($_POST['searchy']))
	{
		//searchy,searchby
		//This is going to be list mode...
		OpenTable2("Search","90%");
		DRWSearch();
		CloseTable();
		OpenTable("Room Layout","90%");
		if($_POST['searchby'] == 'room') {DRWList('WHERE `room` = "'.$_POST['searchy'].'"'); echo $_POST['searchy'];}
		if($_POST['searchby'] == 'teacher') {//Get a second query...
			DRWList('',1);}
		if($_POST['searchby'] == 'type') {DRWList('',2);}
		CloseTable();
	}
	else{
		if(isset($_REQUEST['class']))
		{
			OpenTable2("Search","90%");
			DRWSearch();
			CloseTable();
			//Starts a plotter... Results will be in an image format...
			$result = mysql_query("SELECT * FROM `rooms` WHERE `room` = ".anti_injection($_REQUEST['class']));
			//$GR=$GR+'<map name="gmap">';
			if($row = mysql_fetch_array($result))
			{
				$setm = explode("\n",$row['data']);
				$RFormat_Size=""; $my_img=""; $IO=-1;
				$Projectors = 0; $PCs=0; $MACs=0; $Printers=0;$TVs=0;$CTables=0;$SIZE="";$PROSE="";$IMP = "prose";$devimg30="";
				?><table style="width:100%"><tr><td style="width:10%;"></td><td><?php 
				OpenTable("Room ".$_REQUEST['class'], '518px');
				
				?><table style="width:100%;height:200px;"><tr><td><div style="width:514px;height:200px;overflow:scroll;background-color:gainsboro;"><table style="wdith:100%;" cellspacing=0><tr style="background-color:darkgray;"><td style="border-right:1px solid black;">Device Type</td><td style="border-right:1px solid black;">Model</td><td style="border-right:1px solid black;">Serial</td><td style="border-right:1px solid black;">Column(x)</td><td style="border-right:1px solid black;">Row(y)</td><td>Rooms (LTR)</td></tr><?php 
				foreach($setm as $A)
				{
					//echo substr($A,0,1);
					if(substr($A,0,8) == 'est_size'){
						$s = explode(': ',$A);
						$SIZE = $s[1];
					}
					if(substr($A,0,5) == 'Prose')
					{
						$s = explode(': ',$A);
						$PROSE = $s[1];
						if($PROSE == "90") $IMP = "prose_90";
						if($PROSE == "180") $IMP = "prose_180";
						if($PROSE == "270") $IMP = "prose_270";
						//echo '<img src="./icons/'.$IMP.'.png" style="border:0;" usemap="#gmap" >';
						$img1 = './icons/'.$IMP.'.png';
						$devimg30 = imagecreatefrompng($img1);
						list($width1,$height1) = getimagesize($img1);
						$e = explode('x', $RFormat_Size);
						//echo $img1;
						//$my_img = imagecreatetruecolor( ((int)$e[0])*100, ((int)$e[1])*100 );
						imagecopy($my_img[$IO],$devimg30,((int)$e[0])*100-60,((int)$e[1])*100-55,0,0,$width1,$height1);
					}
					if(substr($A,0,12) == 'RFormat_Size'){
						$IO++;
						$s = explode(': ',$A);
						$RFormat_Size = $s[1];
						$e = explode('x', $RFormat_Size);
						$my_img[$IO] = imagecreatetruecolor( ((int)$e[0])*100, ((int)$e[1])*100 );
						
						$ioff[$IO] = '';

					}
					if(substr($A,0,6) == 'Offset'){
						$s = explode(': ',$A);
						$ioff[$IO] = $s[1];
					}
					if(substr($A,0,7) == 'CTables'){
						$s = explode(': ',$A);
						$CTables = $s[1];
					}
					else if(substr($A,0,1) == '+')
					{
						$A1 = substr($A,1,strlen($A)-1);
						$s = explode(':',$A1);
						//Desc
						//Serial
						//Bar
						//Location
						//Date
						$RR = $s[1].$s[2];
						if($s[1] != '' && $s[2] != '') $RR = $s[1].' / '.$s[2];
						$TYPE1 = $s[0];
						$date = $row['Date'];
						$TYPE=GetTyper($s[0]);
						echo '<tr style="font-size:10pt;"><td style="border-right:1px solid black;">'.BLK().$TYPE.'</td><td style="border-right:1px solid black;">'.BLK().$s[0].'</td><td style="border-right:1px solid black;">'.BLK().$RR.'</td><td style="border-right:1px solid black;"align="center">'.BLK().'</td><td style="border-right:1px solid black;"align="center">'.BLK().'</td><td>'.$IO.'</td></tr>';
						$I++;
					}
					if(substr($A,0,1) == '[')
					{
						$s = explode(':',$A);
						//$s[0] == coords
						//$s[1] == type
						//$s[2] == model
						//$s[3] == Serial
						//$s[4] == comment
						$R22 = str_replace(')','',$s[0]);
						$coords = explode(',',str_replace("[",'',$R22));
						list($width,$height) = getimagesize($img);
						$RS = $s[1];
						if($RS == 'Printer') $RS = 'Print';
						if($RS == 'Projector') $RS = 'Proj';
						$st = $RS;
						if($s[2] == '') $st = $RS;
						else $st = $RS.'-'.$s[2];
						$st = str_replace('DesignJet','JT',str_replace('LaserJet','LJ',$s[2]));
						$st = str_replace('ScanJet','SJ',str_replace('Laser Printer','LP',$st));
						$TRS = str_replace(hex2bin('0d'),'',$s[3]);
						if($TRS == 'SERVICETAGHERE') $TRS = '';
					//echo $s[3];
						
						if($s[1] == 'PC') $PCs++;
						if($s[1] == 'Projector') $Projectors++;
						if($s[1] == 'Printer') $Printers++;
						if($s[1] == 'MAC') $MACs++;
						if($s[1] == 'TV') $TVs++;
						if($s[1] == 'CTable') $CTables++;
						if($s[1] != 'Door')
							echo '<tr style="font-size:10pt;"><td style="border-right:1px solid black;">'.BLK().$s[1].'</td><td style="border-right:1px solid black;">'.BLK().$s[2].'</td><td style="border-right:1px solid black;">'.BLK().$s[3].'</td><td style="border-right:1px solid black;"align="center">'.BLK().$coords[0].'</td><td style="border-right:1px solid black;"align="center">'.BLK().$coords[1].'</td><td>'.$IO.'</td></tr>';
						//echo $s[1].'- Model: '.$s[2].', Serial Tag: '.$s[3].'<br />';
						//$GR=$GR+"<area shape='rect' alt='hi!' coords='".($coords[0]*100).",".($coords[1]*100).",".(($coords[0]*100)+50).",".(($coords[1]*100)+50)."' href='?page=".ILink('ATCDATA')."&class<form style="display: inline;"class']."&device="."' />";
					}
					if(substr($A,0,1) == '(')
					{
						$s = explode(':',$A);
						//$s[0] == coords
						//$s[1] == type
						//$s[2] == model
						//$s[3] == Serial
						//$s[4] == comment
						$background = imagecolorallocate( $my_img, 255, 255, 255 );
						$text_colour = imagecolorallocate( $my_img, 255, 255, 0 );
						$line_colour = imagecolorallocate( $my_img, 128, 255, 0 );
						$img = './icons/'.$s[1].'.png';
						$devimg = imagecreatefrompng($img);
						$R22 = str_replace(')','',$s[0]);
						$coords = explode(',',str_replace("(",'',$R22));
						$devingh=90;$devingw=90;
						//imagealphablending($devimg, true);
						//imagesavealpha($devimg, false);
						list($width,$height) = getimagesize($img);
						imagecopyresized($my_img[$IO],$devimg,$coords[0]*100,$coords[1]*100,0,0,50,50,$width,$height);
						$textcolor = imagecolorallocate($my_img[$IO], 255, 0, 1);
						//echo "|".$textcolor."|";
						/*$array = preg_split('//', $s[3], -1, PREG_SPLIT_NO_EMPTY);
						foreach($array as $r)
						{
							echo bin2hex($r).'<br />';
						}*/
						$RS = $s[1];
						if($RS == 'Printer') $RS = 'Print';
						if($RS == 'Projector') $RS = 'Proj';
						$st = $RS;
						if($s[2] == '') $st = $RS;
						else $st = $RS.'-'.$s[2];
						$st = str_replace('DesignJet','JT',str_replace('LaserJet','LJ',$s[2]));
						$st = str_replace('ScanJet','SJ',str_replace('Laser Printer','LP',$st));
						$TRS = str_replace(hex2bin('0d'),'',$s[3]);
						if($TRS == 'SERVICETAGHERE') $TRS = '';
						imagestring($my_img[$IO],2,$coords[0]*100+3,($coords[1]*100)+50+3,$TRS,$textcolor);
						imagestring($my_img[$IO],3,$coords[0]*100,($coords[1]*100)-3,$st,$textcolor);
					//echo $s[3];
						
						if($s[1] == 'PC') $PCs++;
						if($s[1] == 'Projector') $Projectors++;
						if($s[1] == 'Printer') $Printers++;
						if($s[1] == 'MAC') $MACs++;
						if($s[1] == 'TV') $TVs++;
						if($s[1] == 'CTable') $CTables++;
						if($s[1] != 'Door')
							echo '<tr style="font-size:10pt;"><td style="border-right:1px solid black;">'.BLK().$s[1].'</td><td style="border-right:1px solid black;">'.BLK().$s[2].'</td><td style="border-right:1px solid black;">'.BLK().$s[3].'</td><td style="border-right:1px solid black;"align="center">'.BLK().$coords[0].'</td><td style="border-right:1px solid black;"align="center">'.BLK().$coords[1].'</td><td>'.$IO.'</td></tr>';
						//echo $s[1].'- Model: '.$s[2].', Serial Tag: '.$s[3].'<br />';
						//$GR=$GR+"<area shape='rect' alt='hi!' coords='".($coords[0]*100).",".($coords[1]*100).",".(($coords[0]*100)+50).",".(($coords[1]*100)+50)."' href='?page=".ILink('ATCDATA')."&class<form style="display: inline;"class']."&device="."' />";
					}
				}
				$O=0;
				foreach($my_img as $R)
				{
					imagepng($R,'./cache/room'.$_REQUEST['class'].'-'.$O.'.png');
					$O++;
				}
				//$GR=$GR+"</map>";
				?></table></div></td><td></td></tr></table><?php 
			
				CloseTable();
				?></td><td><?php 
				OpenTable("Room Data","400px");
				echo 'Estimated Size: ['.$SIZE.']<br />';
				echo 'PCs: '.$PCs.'<br />';
				echo 'MACs: '.$MACs.'<br />';
				echo 'Projectors: '.$Projectors.'<br />';
				echo 'Printers: '.$Printers.'<br />';
				echo 'Computer Tables: '.$CTables.'<br />';
				CloseTable();
				?></td><td style="width:10%;"></td></tr></table><?php 
				//Lp Voad the just-saved image...
				OpenTable("Visual Class map (Room ".$_REQUEST['class'].")","90%");
				$PT=0;
				?><table cellspacing=0><tr><?php 
				foreach($my_img as $T)
				{
					$RR = "";
					if($ioff[$PT] != '') $RR = ' valign="'.str_replace('\n','',$ioff[$PT]).'"';
					echo '<td'.$RR.'><img src="./cache/room'.$_REQUEST['class'].'-'.$PT.'.png" style="border:0;" /></td>';$PT++;
				}
				?></tr></table><?php 
				//echo $GR;
				CloseTable();
				
				OpenTable("Information","90%");
				echo '<b>Abbreviations:</b><br />Proj - Projector<br />Print - Printer<br />DJ - DeskJet (Large)<br />DDJ - DesignJet (Large)<br />LJ - LaserJet<br />SJ - ScanJet<br /><br /><b>Model / Producer: (* = unique model ID digit)</b>'.
				'<br />Optiplex *** / <b>Dell</b> (Abv. as: Opti ***)<br />Optiplex GX*** / <b>Dell</b> (Abv. as: GX***)<br />2300-2400MP Projector / <b>Dell</b>';
				CloseTable();
			}
		}
		else{
			?><table style="width:100%"><tr><td><?php 
			OpenTable2("Search","400px");
			DRWSearch();
			CloseTable();
			?></td><td><?php 
			OpenTable("Data","200px");
			?><a href="?page=export">Get Data (Export)</a><?php 
			CloseTable();
			?></td></tr></table><?php 
			
			OpenTable("Room Layout","90%");
	
			if(isset($_COOKIE['ELAYOUT']))
			{
				if($_COOKIE['ELAYOUT'] == 0){DRWButtons(0); DRWCompact();} 
				if($_COOKIE['ELAYOUT'] == 1){DRWButtons(1); DRWList();}
				if($_COOKIE['ELAYOUT'] == 2){DRWButtons(2); DRWMap();}
			}
			else 
			{
				DRWButtons(2); DRWCompact();
				setcookie("ELAYOUT", 2);
			}	
			CloseTable();
		}
	}
}
?>