<?php 
/*require_once './Spreadsheet/Excel/Writer.php';

$workbook = new Spreadsheet_Excel_Writer();
$workbook->send('atc_rooms.xls');
$worksheet =& $workbook->addWorksheet('ATCRooms');
$worksheet->writeString(0, 0, 'Room#');
$worksheet->writeString(1, 0, 'Item');
$worksheet->writeString(2, 0, 'IModel');
$worksheet->writeString(3, 0, 'ISerial');
$worksheet->writeString(4, 0, 'Column X');
$worksheet->writeString(5, 0, 'Row Y');
$worksheet->writeString(6, 0, 'Room Map (Image)');*/


?>
You can get the data from this page using Excel. <br />
<a href="http://office.microsoft.com/en-us/excel-help/getting-data-from-the-web-in-excel-HA001045085.aspx">Learn howto here (Excel)</a><br />
<a href="http://wiki.services.openoffice.org/wiki/Documentation/OOo3_User_Guides/Calc_Guide/Linking_to_external_data">Learn howto here (Openoffice.org Calc)</a><br /><br />
<table style="width:100%;" name="ATC_ROOM_DATA">
<tr><td>Type</td><td>Description</td><td>Serial Tag=Barcode</td><td>Coord</td><td>Location</td><td>Date</td></tr>

<?php 

$result = mysql_query("SELECT * FROM `rooms`");
$I = 0;
while($row = mysql_fetch_array($result))
{
	$setm = explode("\n",$row['data']);
	$Projectors = 0; $PCs=0; $MACs=0; $Printers=0;$TVs=0;
	$RFormat_Size=""; $my_img='';
	$Projectors = 0; $PCs=0; $MACs=0; $Printers=0;$TVs=0;
	/*?><table style="width:100%"><tr><td style="width:10%;"></td><td><?php */
	//OpenTable("Room <form style="display: inline;"class'], '458px');
	
	if($row['room'] != ' ')
	{
		$dat = '';
		$UI = 0;
		while(true)
		{
			if(file_exists("./cache/room".$row['room']."-".$UI.".png"))
				$dat .= 'https://74.196.209.131/cache/room'.$row['room'].'-'.$UI.'.png ';
			else break;
			$UI++;
		}
		echo '<tr><td></td><td><b>'.$row['room'].'</b></td><td>'.$dat.'</td><td>< - Room '.$row['room'].' Room Map</td><td></td><td></td></tr>';
	}
	
	//$worksheet->writeString(0, $I, $row['room']);
	//$worksheet->insertBitmap(6, $I, './cache/room'.$row['room'].'.png');
	/*?><table style="width:100%;height:200px;"><tr><td><div style="width:454px;height:200px;overflow:scroll;background-color:gainsboro;"><table style="wdith:100%;" cellspacing=0><tr style="background-color:darkgray;"><td style="border-right:1px solid black;">Device Type</td><td style="border-right:1px solid black;">Model</td><td style="border-right:1px solid black;">Serial</td><td style="border-right:1px solid black;">Column(x)</td><td>Row(y)</td></tr><?php */
	$SIZE='';
	
	foreach($setm as $A)
	{
		if(substr($A,0,12) == 'RFormat_Size'){
			$s = explode(': ',$A);
			$RFormat_Size = $s[1];
			$e = explode('x', $RFormat_Size);
			$SIZE = $e;

		}
		else if(substr($A,0,1) == '+')
		{
			$A1 = substr($A,1,strlen($A)-1);
			$s = explode(':',$A1);
			$RR = $s[1].$s[2];
			if($s[1] != '' && $s[2] != '') $RR = $s[1].' / '.$s[2];
			$TYPE1 = $s[0];
			$date = $row['Date'];
			$TYPE=GetTyper($row['Description']);
			echo '<tr><td>'.$TYPE.'</td><td>'.$s[0].'</td><td>'.$RR.'</td><td>[none]</td><td>'.$s[3].'</td><td>'.$s[4].'</td></tr>';
			$I++;
		}
		else if(substr($A,0,1) == '(')
		{
			$s = explode(':',$A);
			//$s[0] == coords
			//$s[1] == type
			//$s[2] == model
			//$s[3] == Serial
			//$s[4] == DATE!!!
			$R22 = str_replace(')','',$s[0]);
			$coords = explode(',',str_replace("(",'',$R22));
			/*$worksheet->writeString(1, $I, $s[1]);
			$worksheet->writeString(2, $I, $s[2]);
			$worksheet->writeString(3, $I, $s[3]);
			$worksheet->writeString(4, $I, $coords[0]);
			$worksheet->writeString(5, $I, $coords[1]);*/
			if(str_replace(' ','',$s[2])=='2300MP') $s[2] = 'Dell 2300MP Projector';
			if(str_replace(' ','',$s[2])=='2400MP') $s[2] = 'Dell 2400MP Projector';
			if(substr($s[2],0,2) == 'GX') $s[2] = 'Dell Optiplex '.$s[2];
			if(substr($s[2],0,9) == 'Precision') $s[2] = 'Dell '.$s[2];
			$s[2] = str_replace('Opti ','Dell Optiplex ',$s[2]);
			$s[2] = str_replace('LJ','LaserJet',$s[2]);
			$s[2] = str_replace('DDJ','DesignJet',$s[2]);
			$s[2] = str_replace('DJ','DeskJet',$s[2]);
			$s[2] = str_replace('SJ','ScanJet',$s[2]);
			$s[2] = str_replace('LP','Laser Printer',$s[2]);
			$s[2] = str_replace('PSmart','Photo Smart',$s[2]);
			$RS = $s[3];
			if($s[1] == '') $TYPE=GetTyper($s[2]);
			$RRR = '('.($coords[0]+1).','.($coords[1]+1).')';
			if($RRR == '(-1,-1)') $RRR = "[none]";
			if($s[2] == '') $s[2] = $s[1].'-unrecorded';
			if(substr($RS,0,14) == "SERVICETAGHERE") $RS = '';
			if($s[1] != 'Door')
				echo '<tr><td>'.$s[1].'</td><td>'.$s[2].'</td><td>'.$RS.'</td><td>'.$RRR.'</td><td>'.$row['room'].'</td><td>'.$s[4].'</td></tr>';
			$I++;
		}
		else if(substr($A,0,1) == '[')
		{
			$s = explode(':',$A);
			//$s[0] == coords
			//$s[1] == type
			//$s[2] == model
			//$s[3] == Serial
			//$s[4] == DATE!!!
			$R22 = str_replace(')','',$s[0]);
			$coords = explode(',',str_replace("(",'',$R22));
			/*$worksheet->writeString(1, $I, $s[1]);
			$worksheet->writeString(2, $I, $s[2]);
			$worksheet->writeString(3, $I, $s[3]);
			$worksheet->writeString(4, $I, $coords[0]);
			$worksheet->writeString(5, $I, $coords[1]);*/
			if(str_replace(' ','',$s[2])=='2300MP') $s[2] = 'Dell 2300MP Projector';
			if(str_replace(' ','',$s[2])=='2400MP') $s[2] = 'Dell 2400MP Projector';
			if(substr($s[2],0,2) == 'GX') $s[2] = 'Dell Optiplex '.$s[2];
			if(substr($s[2],0,9) == 'Precision') $s[2] = 'Dell '.$s[2];
			$s[2] = str_replace('Opti ','Dell Optiplex ',$s[2]);
			$s[2] = str_replace('LJ','LaserJet',$s[2]);
			$s[2] = str_replace('DDJ','DesignJet',$s[2]);
			$s[2] = str_replace('DJ','DeskJet',$s[2]);
			$s[2] = str_replace('SJ','ScanJet',$s[2]);
			$s[2] = str_replace('LP','Laser Printer',$s[2]);
			$s[2] = str_replace('PSmart','Photo Smart',$s[2]);
			$RS = $s[3];
			if($s[1] == '') $TYPE=GetTyper($s[2]);
			if(substr($RS,0,14) == "SERVICETAGHERE") $RS = '';
			if($s[1] != 'Door')
				echo '<tr><td>'.$s[1].'</td><td>'.$s[2].'</td><td>'.$RS.'</td><td>('.($coords[0]+1).','.($coords[1]+1).')</td><td>'.$row['room'].'</td><td>'.$s[4].'</td></tr>';
			$I++;
		}
		
	}
	/*?></tr><?php */
	$I++;
}

?></table><?php 
//$workbook->close();
?>
