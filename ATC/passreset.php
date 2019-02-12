<?php
if(FI == md5("F@G!HWP*"))
{
	OpenTable("Send Password Reset Email","350px");
	echo '<form style="display: inline;" method="post" action="./?page='.ILink("preset").'">
	Email: <input type="text" /><br />
	<input type="submit" value="Send" />
	</form>';
	CloseTable();
}
?>