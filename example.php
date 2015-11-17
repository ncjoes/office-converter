<?php
require_once('OfficeConverter.php');

$tempPath = "E:\\tmp\\libreoffice";
$bin = '"C:\Program Files (x86)\LibreOffice 5\program\soffice.exe"';
$convert2 = new OfficeConverter('test.xlsx', $bin, $tempPath);
var_dump($convert2->convertTo('test3.pdf'));
