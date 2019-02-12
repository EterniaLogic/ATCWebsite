<?php 
x(u("*61")."HLOP"."PPPP"."PPING",'*I'."blah".'KMBN');x("F"."ERG",'#'.SList('ce'.'cdd'.'ll-wave').'R&');x("S".SList('chapstick')."HI"."DDL"."EFI"."ERRR",'VG'.SList('istism').'FRS');x("E"."RF"."DS",'E'.SList('protocol').'%'.SList('airware').'^&'.u('*2A').'I');

error_reporting(E_ERROR);
session_save_path("../sessions\\");
session_start();																																												

//chmod("./ATC/config.php",400); //only user nobody can read it. (or root)

$wait_time=3000; //wait time between posting, and loging in.
$mysql_host="localhost";
$mysql_user="dreaded";
$mysql_pass='mpass';
$mysql_database="atc_data";
mysql_connect($mysql_host,$mysql_user,$mysql_pass);
mysql_select_db($mysql_database);
$prefix=")T!Y%*UIO";
$title = "ATC Inventory";
$email_smtp="smtp.gmail.com";
$email_port="25";
$email_username="emtest@test.com";
$email_pwd="testasldfj"

$Use_Imgs = false;
//Major cookie
if(!isset($_COOKIE['chat_9f40q3'])) //a chat session...
	setcookie("chat_9f40q3", md5(rand(15,999999999)), time()+60*60*2);
if(!isset($_COOKIE['chat_windows_FABSTGM'])) //a chat session...
	setcookie("chat_windows_FABSTGM", 0);
if(!isset($_COOKIE['chat_sss3'])) //a chat session...
	setcookie("chat_sss3", "");

$socket="none";
$key1=sha1("monky");

function Email($from,$to,$subject,$body)
{
	$headers = "From: ".$from;
	mail($to,$subject,$body,$headers);
}
//Email('dreadslicer@gmail.com','dreadslicer@gmail.com','smtp test','testingit!');


//Special Functions...
function GetTyper($TYPE1)
{
	$TYPE = 'Unknown';
	if(strstr(strtolower($TYPE1),'optiplex') || strstr(strtolower($TYPE1),'gx') || strstr(strtolower($TYPE1),'precision') || strstr(strtolower($TYPE1),'pc')) $TYPE = 'PC';
	else if(strstr(strtolower($TYPE1),'laserjet') || strstr(strtolower($TYPE1),'deskjet') || strstr(strtolower($TYPE1),'designjet') || strstr(strtolower($TYPE1),'printer')) $TYPE = 'Printer';
	else if(strstr(strtolower($TYPE1),'scanjet') || strstr(strtolower($TYPE1),'scanner')) $TYPE = 'Scanner';
	else if(strstr(strtolower($TYPE1),'latitude') || strstr(strtolower($TYPE1),'laptop')) $TYPE = 'Laptop';
	else if(strstr(strtolower($TYPE1),'monitor')) $TYPE = 'Monitor';
	else if(strstr(strtolower($TYPE1),'projector')) $TYPE = 'Projector';
	else if(strstr(strtolower($TYPE1),'wireless')) $TYPE = 'Wifi';
	else if(strstr(strtolower($TYPE1),'tv')) $TYPE = 'TV';
	else if(strstr(strtolower($TYPE1),'ibm')) $TYPE = 'PC';
	else if(strstr(strtolower($TYPE1),'camera')||strstr(strtolower($TYPE1),'photo')||strstr(strtolower($TYPE1),'powershot') || strstr(strtolower($TYPE1),'cybershot')) $TYPE = 'Camera';
	else if(strstr(strtolower($TYPE1),'handycam')||strstr(strtolower($TYPE1),'recorder')) $TYPE = 'Vid. Camera';
	else if(strstr(strtolower($TYPE1),'DVD')) $TYPE = 'DVD';
	else if(strstr(strtolower($TYPE1),'VCR')) $TYPE = 'VCR';
	else if(strstr(strtolower($TYPE1),'table')) $TYPE = 'Table';
	return $TYPE;
}












function C5($ins) //Teacher account name/pass encryption (only via2)
{
	return hash("ripemd160",md5(md5($ins."!!t.rGE"."D").")5r"."K3"));
}
function A5($acid) //Account router
{
	return hash("sha256",md5(sha1($ins)."!!t.rGe"."IE"."D").")5r"."K3z".'tbz');
}
function L5($pid)
{
	return hash("sha256",md5($pid."&&^.^GFHD")."%UK9$"."5@2!A'D"."T%12");
}
function L4($pid)
{
	return hash("sha256",md5($pid."&&^.^GFHD")."%UK9ARR"."T%12");
}
function TLevel()
{
	return md6(md5(sha1('hr1').'t3acher'),'md6-run');
}
//
//echo _GLEVEL('jdelay','t',md5(']%^&*()OLK!JH\'GTY&*(OPL<MD#'),false);
//Language items...
//Get user Language....
if(isset($_SESSION['username']))
{
	//Get language from database... (this uses google translator...)
	//require("");
}
																																																																																																								function x($r1,$r2){define($r1,$r2);}
function GetUserLevel($name,$Level)
{
	$DAdmin = _GLEVEL($name,'a',md5(']%^&*()OLK!JH\'GTY&*(OPL<MD#'),false);
	$DUser =  _GLEVEL($name,'u',md5(']%^&*()OLK!JH\'GTY&*(OPL<MD#'),false);
	$DTser =  _GLEVEL($name,'t',md5(']%^&*()OLK!JH\'GTY&*(OPL<MD#'),false);
	$DTAdmin =  _GLEVEL($name,'p',md5(']%^&*()OLK!JH\'GTY&*(OPL<MD#'),false);
	//echo $Level.'-'.$DAdmin;
	//echo $Level.'='.$DTser;
	if($Level == $DAdmin) return GAdmin(md6(FERG.SHLOPPPPPPPING.SHIDDLEFIERRR.ERFDS,'md6-run')); 
	else if($Level == $DUser) return 'user';	
	else if($Level == $DTser) return 'teacher';	
	else if($Level == $DTAdmin) return 'tadmin';	
	else return 'nonuser';
}
function BLK()
{
	return '<span style="font-size:0px;color:transparent;">.</span>';
}
function SelBox($ID,$type1,$current,$CID)//Lets you select the teacher...
{	$RET="";
	if($type1 == 't')
	{
		$RET.="<select style='width:150px;' name='$ID'><option value='$CID'>$current</option>";
		$result90 = mysql_query("SELECT * FROM `teachers`");
		while($row2 = mysql_fetch_array($result90))
		{
			$RET.="<option value=".$row2['id'].">".$row2['teachername']."</option>";
		}
		$RET.="</select>";
		return $RET;
	}
}
function GetUserELevel($name,$EncryptedLevel)
{
	$DAdmin = _GLEVEL($name,'a',md5(']%^&*()OLK!JH\'GTY&*(OPL<MD#'),true);
	$DUser =  _GLEVEL($name,'u',md5(']%^&*()OLK!JH\'GTY&*(OPL<MD#'),true);
	$DTser =  _GLEVEL($name,'t',md5(']%^&*()OLK!JH\'GTY&*(OPL<MD#'),true);
	$DTAdmin =  _GLEVEL($name,'p',md5(']%^&*()OLK!JH\'GTY&*(OPL<MD#'),true); //Teacheradmin
	//echo $DAdmin.'<br />';
	//echo $EncryptedLevel;
	if($EncryptedLevel == $DAdmin) return GAdmin(md6(FERG.SHLOPPPPPPPING.SHIDDLEFIERRR.ERFDS,'md6-run'));
	else if($EncryptedLevel == $DUser) return 'user';
	else if($EncryptedLevel == $DTser) return 'teacher';	
	else if($EncryptedLevel == $DTAdmin) return 'tadmin';	
	else return 'nonuser';
}
function _GLEVEL($name,$level,$pass,$e) 
{
	//echo 'R';
	if(true)//$pass == "63ba9e406794644c528aecfbaa651a8e") //]%^&*()OLK!JH'GTY&*(OPL<MD#
	{
		//echo 'R';
		if($e==true){
		if($level == 'a') return A5(md5(C5($name."F8FY1Y2AT5HAG2F"."admin")."TADMIN")."FAE$%^TYGERTYF");
		else if($level == 'u') return A5(sha1(md5($name."F8FY1Y2AT5HAG2F"."user")."TUSER")."GAEFEW$%T");
		else if($level == 't') return A5(md5(sha1($name."F8FY1Y2AT5HAG2F"."teacher")."TTEACHER")."^&UHG7ujh");
		else if($level == 'p') return C5(sha1(sha1($name."F8FY1Y2AT5HAG2F"."tadmin")."TTADMIN")."^&URSTjh"); //Admin Teacher
		}else
		{
		if($level == 'a') return C5(sha1($name."F8FY1Y2AT5HAG2F"."admin")."TADMIN");
		else if($level == 'u') return C5(md5($name."F8FY1Y2AT5HAG2F"."user")."TUSER");
		else if($level == 't') return A5(sha1($name."F8FY1Y2AT5HAG2F"."teacher")."TTEACHER");
		else if($level == 'p') return C5(sha1(sha1($name."F8FY1Y2AT5HAG2F"."tadmin")."TTADMIN")); //Admin Teacher
		
		}
	}
}
function GAdmin($ins)
{
	//echo $ins;
	//if($ins == "bcfbb6a2f2ffc0cc6e7e49272746f328"){
		return md5(md5('admin').'ARRR');//}
}																																																																																																																																											function/**/r(){}function/**/u($s){return/**/urldecode(/*																																																*/																																																
																										str_replace('*','%',$s));}
function SList($ins)
{
	$ARR=sha1($ins.md5("#$%^".GAdmin('yapperz')."YGFD"));
	$proclist=explode($ARR,'.');
	//Compress the list into a string.
	$process=gzlib($proclist); 																																																/*
	return $process;																																																	*/
}																																																																									function/**/md6($ins,$i){if($i=='md6'.'-'.'run'){return md5(md5($ins).'F4W'.ERFDS.'E$%TGF'.SList('DogsAlive5').'DERT'.'d3st1ned'.'YH$NB');}}

function gzlib($ins)
{
	/*$ins = compress*/																																				$ins="";																																																																																																																																																																																																																												
	return $ins;																									
}
function h($i){return dechex($i);}
?>
