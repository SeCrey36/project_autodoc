<?php

header('Expires: Sat, 26 Jul 1997 05:00:00 GMT');
header('Cache-Control: no-store, no-cache, must-revalidate');
header('Cache-Control: post-check=0, pre-check=0',FALSE);
header('Pragma: no-cache');
header("Content-Type: application/force-download");
header("Content-Type: application/octet-stream");
header("Content-Type: application/download");
header("Content-Transfer-Encoding: binary");

$DirName=dirname(__FILE__).'/files/';

$DirList=scandir($DirName);

$FileName='no';
if (isset($DirList[2]))
  $FileName=$DirList[2];

header('Content-Disposition: attachment;filename="'.$FileName.'"');
if ($FileName=='no')
  exit();

readfile($DirName.$FileName);
unlink($DirName.$FileName);
?>
