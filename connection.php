<?php
  $server = "Server1";
  $user = "root";
  $pass = "moth34board";
  $name = "ptc_ims";
  
  $con = mysqli_connect($server, $user, $pass, $name);
  if(mysqli_connect_errno()){
    die("Database connection failed." .
      mysqli_connect_errno() .
      "(" . mysqli_connect_errno() . ")"
    );
  }

?>