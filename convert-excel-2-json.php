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
$fieldName = array();
$json = array();
if ($sheetData) {
    foreach ($sheetData as $dataKey => $data) {
        if ($dataKey == 0) {
            $fieldName = array_values($data);
        } else {
            $json[] = "{'$fieldName[0]':'$data[0]','$fieldName[1]':'$data[1]','$fieldName[2]':'$data[2]','$fieldName[3]':'$data[3]','$fieldName[4]':'$data[4]'}";
        }

    }
}
$jsonStr = "[".implode(",", $json)."]";

echo "<pre>";
print_r($fieldName);
echo "</pre>";
echo "<pre>";
print_r($jsonStr);
echo "</pre>";
file_put_contents("subject.json", $jsonStr);