<?php 
//header page... Unique.

$width="90%";

if(FI == md5("F@G!HWP*"))
{

$APP_ID="FTTFDDD";

function MiniTable()
{
	//used to add a minimize botton to the Table window.
}

function OpenTable_($nam,$width,$s1,$s2,$s3,$s4,$noecho=false)
{
	//sets a cookie that counts how many cookies there are
	$_COOKIE['$sss3']=$s3;
	if(isset($_COOKIE['chat_windows_FABSTGM']))
		$_COOKIE['chat_windows_FABSTGM']++;
	//define("II",II+1); //used to make it literally impossible for a element to have the same id. Even if they have the same name and width and page.
	$randid = substr(md5( $_COOKIE['chat_windows_FABSTGM'] .$_REQUEST['page'].$nam.$width),0,7); //each box on each page is unique.
	$TableDat = '<br /><table cellspacing=0 style="width:100%;" id="box_'.$randid.'">
				<tr>
					<td align=center>
					<center>
							<div align=left class="'.$s1.'" style="width:'.$width.';">
								<table name="head" cellspacing=0 class="'.$s2.'" style="width:100%;height:22px;"><tr><td class="'.$s4.'">&nbsp;'.$nam.'</td><td name="mini" style="width:10px;height:100%;" align=right><div name="mini" id="minimize" onmouseover="mini_mousein(this);" onmouseout="mini_mouseout(this);" onclick="do_mini(this,\''.$APP_ID.'tblbox_'.$randid.'\');"></div></td></tr></table>
								<div id="'.$APP_ID.'tblbox_'.$randid.'" style="visibility:visible;">';
	if($noecho == false)
	{
		echo $TableDat;
		return $randid; //use for returning purposes...lolz
	}else return $TableDat;
}
function OpenTable($nam,$width,$noecho=false)
{
	return OpenTable_($nam,$width,"tablebox","tablebox_data","tbl_bottom","tbl_title_alignment",$noecho);
}
function OpenTable2($nam,$width)
{
	return OpenTable_($nam,$width,"tablebox2","tablebox2_data","tbl2_bottom","tbl2_title_alignment");
}
function CloseTable($noecho=false)
{
$TData='						</div>
						<!--<table style="width:100%;" cellspacing=0 class="'.$_COOKIE['$sss3'].'"><tr><td>.</td></tr></table>-->
						</div>
						</center>
					</td>
				</tr>
			</table>';
			if($noecho) return $TData;
			else echo $TData;
}

echo '<html>
	<head>
		<title>'.$title.'</title>
		<script src="./ATC/js/pagejs.js"></script>
		<style>
		html, body {height:100%; margin:0; padding:0;}

.tablebox_data {background-image:url(ATC/img/tbl_header.png);color:darkred;height:20px;font-weight: bold;}
.tablebox {
background-color:#C1CAFF;
border-left:1px solid black;
border-top:1px solid black;
border-right:3px solid #494946;
border-bottom:3px solid #494946;
}
.tablebox_data a {color:gray;text-decoration:none;}
.tablebox_data a:hover {color:red;text-decoration:none;}
.tbl_title_alignment {text-align:center;}
.tbl_bottom {width:100%;height:5px;font-size:1px;color:transparent;background-image:url(/ATC/img/tbl_bottom.png);}
.tablebox2_data {background-image:url(/ATC/img/tbl_header2.png);color:red;height:22px;}
.tablebox2 {
background-color:#C9E1FF;
border-left:1px solid black;
border-top:1px solid black;
border-right:3px solid #494946;
border-bottom:3px solid #494946;
}
.tablebox2_data a {color:red;text-decoration:none;}
.tablebox2_data a:hover {color:darkred;text-decoration:none;}
.tbl2_title_alignment {text-align:left;}
.tbl2_bottom {width:100%;height:5px;font-size:1px;color:transparent;background-image:url(/ATC/img/tbl2_bottom.png);}
#mover1 {width:19px;background-image:url(/ATC/img/mouse.jpeg););}
.headeri {background-image:url(/ATC/img/header.png);}
.headeri a {color:white;text-decoration:none;}
.headeri a:hover {color:red;text-decoration:none;}
#minimize {background-image:url(/ATC/img/mini.png);width:10px;height:10px;cursor: pointer;font-size:10px;}
#minimize_over {background-image:url(/ATC/img/mini2.png);width:10px;height:10px;cursor: pointer;font-size:10px;}
#minimize_down {background-image:url(ATC/img/mini2.png);width:10px;height:10px;cursor: pointer;font-size:10px;}
#minimize_down_over {background-image:url(/ATC/img/mini3.png);width:10px;height:10px;cursor: pointer;font-size:10px;}
.img_preload1{background-image: url(/ATC/img/mini3.png); background-repeat: no-repeat; background-position: -1000px -1000px; }
.img_preload2{background-image: url(/ATC/img/mini2.png); background-repeat: no-repeat; background-position: -1000px -1000px; }
.img_preload3{background-image: url(/ATC/img/minin.png); background-repeat: no-repeat; background-position: -1000px -1000px; }
.tlink:hover {text-decoration:underline;font-size:11pt;color:blue;} .ntlink:hover {text-decoration:none;font-size:9pt;color:black;}
.tlink {text-decoration:none;font-size:8pt;color:navy;}
.boxlist {border-bottom:1px solid black;}
		</style>
		
	</head>
<body>
	<table style="width:100%;height:100%;background-color:white;">
	<tr>
	<td align=center valign=top>
		<span class="img_preload1"></span>
		<span class="img_preload2"></span>
		<span class="img_preload3"></span>
		<table style="width:'.$width.';background-color:gainsboro;">
		<tr>
		<td>
			<table style="width:100%;height:100px;">
				<tr>
					<td style="width:90%;height:100px;background: url(ATC/img/logo.jpg) no-repeat center top;">
						
					</td>
				</tr>
				<tr height=22>
					<td class="headeri" style="background-color:darkred;">';
						 //configures which is underlined.
							//
							
							if(isset($_SESSION['username']))
							{
								if($page == ILink("home") && !(isset($_REQUEST['subpage']))){ echo '<span style="color:gray;">Home</span>&nbsp;&nbsp'; }
								else { mkpageILink("home","Home");echo '&nbsp;&nbsp;'; }
								if($page == ILink("ATCDATA") && !(isset($_REQUEST['subpage']))){ echo '<span style="color:gray;">Atc Data</span>&nbsp;&nbsp;'; } 
								else { mkpageILink("ATCDATA","Atc-Data");echo '&nbsp;&nbsp;'; }
								
								if($page == ILink("management") && !(isset($_REQUEST['subpage']))){ echo '<span style="color:gray;">Management</span>&nbsp;&nbsp;'; } 
								else { mkpageILink("management","Management");echo '&nbsp;&nbsp;'; }
								if($page == ILink("profile_edit") && !(isset($_REQUEST['subpage']))){ echo '<span style="color:gray;">Profile</span>&nbsp;&nbsp;'; } 
								else { mkpageILink("profile_edit","Profile");echo '&nbsp;&nbsp;'; }
								mkpageILink("logout","Logout");echo '&nbsp;&nbsp;';
							}
							else
							{
								if($page == ILink("login") && !(isset($_REQUEST['subpage']))){ echo '<span style="color:gray;">Login</span>&nbsp;&nbsp;'; } 
								else { mkpageILink("login","Login");echo '&nbsp;&nbsp;'; }
								//if($page == ILink("reg") && !(isset($_REQUEST['subpage']))){ echo '<span style="color:gray;">Register</span>&nbsp;&nbsp;'; } 
								//else { mkpageILink("reg","Register");? >&nbsp;&nbsp; }
							}
						
					echo '</td>
				</tr>
			</table>';
}
 ?>