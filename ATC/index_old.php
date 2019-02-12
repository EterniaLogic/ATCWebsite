<?php 
//--------------License----------------------//
//--Only use for brclancy111's Personal use--//

 
//bleh!

//error_reporting(E_ERROR);
//if($_COOKIE['a111%Y#$&%^'] == md5("apf9urf90u"))
//{

require_once("./config.php");

function nl() {
    echo "<br/> \\n";
}
function RandID()
{
	return md5(rand(9999,999999999999999999).rand(9999,999999999999999999));
}
function encrypt_url($string) {
        $key = "123abc1"; //preset key to use on all encrypt and decrypts.
        $result = '';
   for($i=0; $i<strlen($string); $i++) {
     $char = substr($string, $i, 1);
     $keychar = substr($key, ($i % strlen($key))-1, 1);
     $char = chr(ord($char)+ord($keychar));
     $result.=$char;
   }
   return urlencode(base64_encode($result));
}
function decrypt_url($string) {
        $key = "123abc1";
        $result = '';
        $string = base64_decode(urldecode($string));
   for($i=0; $i<strlen($string); $i++) {
     $char = substr($string, $i, 1);
     $keychar = substr($key, ($i % strlen($key))-1, 1);
     $char = chr(ord($char)-ord($keychar));
     $result.=$char;
   }
   return $result;
}
function doserver()
{
	$address="127.0.0.1";
	$port=1066;
	$socket=socket_create(AF_INET, SOCK_STREAM, SOL_TCP);
	socket_connect($socket, $address, $port);
}
function SmallSpacer(){return '<span style="font-size:5pt;color:transparent;">.</span>';}
function anti_injection($sql) {
  $sql = htmlspecialchars($sql); # for < and >, &, single and double quote characters
  $sql = str_replace("'", '& #039;', $sql);  # for single quote characters (just in case)
  $sql = str_replace('"', '&quote;', $sql); # for dpuble quote characters (just in case)
  $sql = str_replace('=', '', $sql); # removes = signs.
  $sql = str_replace('INSERT', '', $sql); # removes INSERT
  $sql = str_replace('SELECT', '', $sql); # removes SELECT
  $sql = str_replace('DELETE', '', $sql); # removes DELETE
  $sql = str_replace(' ', '', $sql); # removes ' '
  $sql = str_replace('+', '', $sql); # removes '+'
  $sql = str_replace('-', '', $sql); # removes all - signs (all sql comments).
  return $sql;
}
function Check_User($user,$pass)
{
	//checks to make sure user is legit... If not, it will remove the user from the session.
	
}
function ILink($name)
{
	return substr(md5($_COOKIE['chat_9f40q3'].$prefix.$name),0,6); //Algorythm for saving the site.
}
function ImgILink($name) //makes it possible to ILink to an image with no problems shown.
{
	return "./?img=".encrypt_url($name);
}
function ImgUnILink($name) //decodes the image
{
	return decrypt_url($name);
}


function mkpageILink($page1,$name1)
{
	echo '<a href="./?page='.ILink($page1).'">'.$name1.'</a>';
	//echo ILink($page1);
}
function mksubpageILink($page1,$subpage1,$name1)
{
	echo '<a href="./?page='.ILink($page1).'&subpage='.ILink($subpage1).'">'.$name1.'</a>';
}
function mksubpageILink_($page1,$subpage1,$name1)
{
	return '<a href="./?page='.ILink($page1).'&subpage='.ILink($subpage1).'">'.$name1.'</a>';
}
function endsWith( $str, $sub ) {
return ( substr( $str, strlen( $str ) - strlen( $sub ) ) == $sub );
}
function TinyMCE($elementid)
{
	echo '<script type="text/javascript" src="plugins/tinymce/jscripts/tiny_mce/tiny_mce.js"></script>
<script type="text/javascript">
	// O2k7 skin
	tinyMCE.init({
		// General options
		mode : "exact",
		elements : "'.$elementid.'",
		theme : "advanced",
		skin : "o2k7",
		plugins : "pagebreak,style,layer,table,save,advhr,advimage,advILink,emotions,iespell,insertdatetime,preview,media,searchreplace,print,contextmenu,paste,directionality,fullscreen,noneditable,visualchars,nonbreaking,xhtmlxtras,template,inlinepopups,autosave",

		// Theme options
		theme_advanced_buttons1 : "save,newdocument,|,bold,italic,underline,strikethrough,|,justifyleft,justifycenter,justifyright,justifyfull,styleselect,formatselect,fontselect,fontsizeselect",
		theme_advanced_buttons2 : "cut,copy,paste,pastetext,pasteword,|,search,replace,|,bullist,numlist,|,outdent,indent,blockquote,|,undo,redo,|,ILink,unILink,anchor,image,cleanup,help,code,|,insertdate,inserttime,preview,|,forecolor,backcolor",
		theme_advanced_buttons3 : "tablecontrols,|,hr,removeformat,visualaid,|,sub,sup,|,charmap,emotions,iespell,media,advhr,|,print,|,ltr,rtl,|,fullscreen",
		theme_advanced_buttons4 : "insertlayer,moveforward,movebackward,absolute,|,styleprops,|,cite,abbr,acronym,del,ins,attribs,|,visualchars,nonbreaking,template,pagebreak,restoredraft",
		theme_advanced_toolbar_location : "top",
		theme_advanced_toolbar_align : "left",
		theme_advanced_statusbar_location : "bottom",
		theme_advanced_resizing : true,

		// Example content CSS (should be your site CSS)
		content_css : "css/content.css",

		// Drop lists for ILink/image/media/template dialogs
		template_external_list_url : "lists/template_list.js",
		external_ILink_list_url : "lists/ILink_list.js",
		external_image_list_url : "lists/image_list.js",
		media_external_list_url : "lists/media_list.js",

		// Replace values for the template plugin
		template_replace_values : {
			username : "Some User",
			staffid : "991234"
		}
	});
	</script>';
}
function EditArea($id)
{
echo '<script language="Javascript" type="text/javascript" src="plugins/editarea/edit_area/edit_area_full.js"></script>
<script language="Javascript" type="text/javascript">
		editAreaLoader.init({
			id: "'.$id.'"	// id of the textarea to transform		
			,start_highlight: true	// if start with highlight
			,allow_resize: "both"
			,allow_toggle: true
			,word_wrap: false
			,language: "en"
			,syntax: "php"	
		});
		/*editAreaLoader.init({
			id: "'.$id.'"	// id of the textarea to transform	
			,start_highlight: true
			,allow_toggle: false
			,language: "en"
			,syntax: "php"	
			,toolbar: "search, go_to_line, |, undo, redo, |, select_font, |, syntax_selection, |, change_smooth_selection, highlight, reset_highlight, |, help"
			,syntax_selection_allow: "css,html,js,php,python,vb,xml,c,cpp,sql,basic,pas,brainfuck"
			,is_multi_files: true
			,EA_load_callback: "editAreaLoaded"
			,show_line_colors: true
		});*/
</script>';
}

if(isset($_REQUEST['img']))
{
	//gets an image...
	//decode image...
	if($Use_Imgs)
	{
		$imgloc = ImgUnILink($_REQUEST['img']);
		$fh = fopen($imgloc, 'r');

		if(endsWith($imgloc,".png")){header("Content-Type: image/png"); echo fread($fh, filesize($imgloc));}
		else if(endsWith($imgloc,".jpg")){header("Content-Type: image/jpg"); echo fread($fh, filesize($imgloc));}
		else if(endsWith($imgloc,".jpeg")){header("Content-Type: image/jpeg"); echo fread($fh, filesize($imgloc));}
		else if(endsWith($imgloc,".gif")){header("Content-Type: image/gif"); echo fread($fh, filesize($imgloc));}
		else{echo "invalid Image!";}
	}
}
else if(isset($_REQUEST['page']))
{
	$page=$_REQUEST['page'];
	
	
	if($_REQUEST['page'] == ILink("home")) //make it impossible to regularly hot-ILink a page...
	{
		define("FI",md5("F@G!HWP*"));
		$title.=" - Home";
		require_once("./header.php");
		require_once("./hme.php");
		require_once("./footer.php");
	}
	else if($_REQUEST['page'] == ILink("login"))
	{
		define("FI",md5("F@G!HWP*"));
		$title.=" - Login";
		require_once("./header.php");
		require_once("./login.php");
		require_once("./footer.php");
	}
	else if($_REQUEST['page'] == ILink("reg"))
	{
		define("FI",md5("F@G!HWP*"));
		$title.=" - Register";
		require_once("./header.php");
		require_once("./reg.php");
		require_once("./footer.php");
	}
	else if($_REQUEST['page'] == ILink("chat"))
	{
		doserver();
		define("FI",md5("F@G!HWP*"));
		$title.=" - Chat";
		require_once("./header.php");
		require_once("./cht.php");
		require_once("./footer.php");
	}
	else if($_REQUEST['page'] == ILink("management"))
	{
		define("FI",md5("F@G!HWP*"));
		$title.=" - Management";
		require_once("./header.php");
		require_once("./management.php");
		require_once("./footer.php");
	}
	else if($_REQUEST['page'] == ILink("profile_edit"))
	{
		define("FI",md5("F@G!HWP*"));
		$title.=" - Edit your Profile";
		require_once("./header.php");
		require_once("./profile_edit.php");
		require_once("./footer.php");
	}
	else if($_REQUEST['page'] == ILink("profile_view"))
	{
		define("FI",md5("F@G!HWP*"));
		$title.=" - ".$AccName." - View Profile";
		require_once("./header.php");
		require_once("./profile_view.php");
		require_once("./footer.php");
	}
	else if($_REQUEST['page'] == ILink("logout"))
	{
		unset($_SESSION['username']);
		unset($_SESSION['pass']);
		header("Location: ./");
	}
	else if($_SESSION['level'] == md5($_SESSION['username']."F8FY1Y2AT5HAG2F"."admin"))
	{
		if ($_REQUEST['page'] == ILink("edit"))
		{
			if($handle = opendir("."))
			{
				while(false !== ($file = readdir($handle)))
				{
					if(!is_dir($file)) 
						if($_REQUEST['subpage']==ILink($file))
						{
							define("FI",md5("F@G!HWP*"));
							$title.=" - Editing file: ".$file;
							require_once("./header.php");
							OpenTable2(mksubpageILink_("management","edit","[Back]")." - Edit - ".$file,"97%");
							$file1=fopen($file,"r");
							$dat1="";
							if(filesize($file) >= 1) $dat1 = fread($file1,filesize($file));
							//TinyMCE("textdata"); //we cannot use it because it will not work...
							EditArea("textdata");
							$dat1=str_replace("<","&#60;",$dat1);
							$dat1=str_replace(">","&#62;",$dat1);
							echo "<form style="display: inline;" method='post' action='?page=".ILink("edit2")."&subpage=".ILink($file)."'><textarea name='textdata' id='textdata' style='width:100%;height:500px;'>".$dat1."</textarea><input type=submit value='X-Mit' /></form>";
							//EditArea("textarea");
							CloseTable();
							require_once("./footer.php");
						}
				}
				closedir($handle);
			}		
		}
		else if ($_REQUEST['page'] == ILink("edit2"))
		{
			if($handle = opendir("."))
			{
				while(false !== ($file = readdir($handle)))
				{
					if(is_dir($file))
					{
						
					}
					if(!is_dir($file)) 
						if($_REQUEST['subpage']==ILink($file))
						{
							define("FI",md5("F@G!HWP*"));
							$title.=" - Editing file: ".$file;
							require_once("./header.php");
							OpenTable2(mksubpageILink_("management","edit","[Back]")." - Edit - ".$file,"97%");
	
							$file2=fopen($file,"w+");
							$string1 = str_replace('\"','"',$_POST['textdata']);
							$string1 = str_replace("\'","'",$string1);
							$string1=str_replace("&#60;","<",$string1);
							$string1=str_replace("&#62;",">",$string1);
							fwrite($file2,$string1);
							//echo $_POST['textdata'];
							$dat1 = $string1;
							$dat1=str_replace("<","&#60;",$dat1);
							$dat1=str_replace(">","&#62;",$dat1);
							//TinyMCE("textdata");
							EditArea("textdata");
							echo "<form style="display: inline;" method='post' action='?page=".ILink("edit2")."&subpage=".ILink($file)."'><textarea name='textdata' id='textdata' style='width:100%;height:500px;'>".$dat1."</textarea><input type=submit value='X-Mit' /></form>";
							
							CloseTable();
							require_once("./footer.php");
						}
				}
				closedir($handle);
			}		
		}
	}

	else
	{
		?><script>location.href="./?page=<?php echo ILink("home"); ?>";</script><?php 
	}
}
else{ 
	?><script>location.href="./?page=<?php echo ILink("home"); ?>";</script><?php 
}



if(isset($_REQUEST['getcht']))
{
	//Getdata from server...
	if(defined(server))
		if(server != "not")
		{
			
		}
}
else{

?>


<?php 
}
/*
}
else
{
	header("Location: ../index.php");
}*/
?>