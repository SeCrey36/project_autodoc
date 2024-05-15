<?php
  ini_set('display_errors',1);
  error_reporting(E_ALL);

  echo '<pre>'.print_r($_FILES,1).'</pre>';

  $SrcFileName=$_FILES['exceldoc']['tmp_name'];
  $DstFileName=dirname(__FILE__).'/files/'.uniqid('doc_').'.xlsx';

  move_uploaded_file($SrcFileName,$DstFileName);

?>