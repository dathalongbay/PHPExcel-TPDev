<?php
require_once('vendor/autoload.php');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;

$extension = "xlsx";

if('csv' == $extension) {
    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Csv();
} else {
    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
}

//$spreadsheet = $reader->load($_FILES['file']['tmp_name']);
$spreadsheet = $reader->load("my_excel_file.xlsx");

$sheetData = $spreadsheet->getActiveSheet()->toArray();

echo "<pre>";
print_r($sheetData);
echo "</pre>";