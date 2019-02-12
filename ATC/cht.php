<script type="text/javascript">

var xmlhttp=null; 

function showChat()
{
if (str.length==0)
  { 
  document.getElementById("chat").innerHTML="";
  return;
  }
if (window.XMLHttpRequest)
  {
  // code for IE7+, Firefox, Chrome, Opera, Safari
  xmlhttp=new XMLHttpRequest();
  }
else
  {
  // code for IE6, IE5
  xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
  }
var url="index.php?getcht="+"<?php echo $gil; ?>";
xmlhttp.open("GET",url,false);
xmlhttp.send(null);
document.getElementById("chat").innerHTML=xmlhttp.responseText;
}
function SendMessage()
{
	//send the message to be processed...
}
function rev()
{
	showChat();
	wait(1000);
	rev();
}
rev();
//while(true){wait(1000);showChat();}

</script>