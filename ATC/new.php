<?php 
function exec_enabled() {
  $disabled = explode(', ', ini_get('disable_functions'));
  return !in_array('exec', $disabled);
}
echo exec_enabled();
function _exec($cmd) 
{ 
   $WshShell = new COM("WScript.Shell"); 
   $oExec = $WshShell->Run($cmd, 0,false); 
   echo $cmd;
   return $oExec == 0 ? true : false; 
}
echo _exec("C:\\Program Files (x86)\\Notepad++\\notepad++.exe");

?>