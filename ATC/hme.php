<?php 
//footer page... Unique.
if(FI == md5("F@G!HWP*"))
{
	OpenTable("Welcome Notice","90%");
	?>Welcome to the ATC Map layout Database Engine!<?php 
	CloseTable();
?><hr /><?php 
$result = mysql_query("SELECT * FROM `news` ORDER BY `id` DESC LIMIT 0 , 10");

	$I=0;
while($row = mysql_fetch_array($result))
{
	$I++;
	$post_id=$row['id'];
	$post_name=$row['article'];
	$post_author=$row['user'];
	$post_data=$row['data'];
	$result2 = mysql_query("SELECT * FROM `account` WHERE name = '".$post_author."'");
	if($row2 = mysql_fetch_array($result2)){
		OpenTable2("#$post_id: $post_name</td><td align=right><a href='./?page=".ILink("profile_view")."&data=".encrypt_url($post_author)."'>".$row2['rlname']."</a>","98%");}
	else{OpenTable2("#$post_id: $post_name</td><td align=right><a href='./?page=".ILink("profile_view")."&data=".encrypt_url($post_author)."'>$post_author</a>","98%");}
	echo $post_data;
	CloseTable();
}

if($I == 0)
{
	?><br />No Articles...<br /><?php 
}
}
?>