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



$fieldName = array();

$final = array();
if ($sheetData) {
    $currentClass = 0;
    foreach ($sheetData as $dataKey => $data) {
        if ($dataKey == 0) {
            $final = array();
            $fieldName = array_values($data);
        } else {
            if (!empty($data[0])) {
                $newId = $data[0];
                $class = $data[1];
                $order = $data[2];
                $title = $data[3];
                $picture = $data[4];
                $oldId = $data[5];
                if ($currentClass != $class) {
                    $final[$class] = array();
                    $currentClass = $class;
                }

                $final[$class][$oldId] = array("id" => $newId,"order" => $order,"title"=>$title,"picture"=>$picture, "old_id" => $oldId);
            }

        }

    }
}
foreach ($final as $finalClass => $finalData) {
    $price = array();
    foreach ($finalData as $key => $row)
    {
        $price[$key] = $row['order'];
    }
    array_multisort($price, SORT_ASC, $finalData);

    $final[$finalClass] = $finalData;
}
foreach ($final as $finalClass => $finalData) {
    $newArr = array();
    foreach ($finalData as $key => $row)
    {
        $oldId = $row["old_id"];
        $newArr[$oldId] = $row;
    }
    $final[$finalClass] = $newArr;
}

// gen code for copy
$code = '';
$code .= '$configSubjects = array();';
$code .= "<br>";
if ($final) {
    $currentClass = 0;
    foreach ($final as $class => $classData) {
        if (!empty($classData)) {
            foreach ($classData as $data) {

                if ($currentClass != $class) {
                    $code .= '$configSubjects['.$class.'] = array();';
                    $code .= "<br>";
                    $currentClass = $class;
                }
                $oldId = $data['old_id'];
                $code .= '$configSubjects['.$class.']['.$oldId.'] = array("id" => '.$data['id'].',"title"=>"'.$data['title'].'","picture"=>"'.$data['picture'].'", "old_id" => '.$data['old_id'].');';
                $code .= "<br>";
            }

        }
    }
}

echo $code;
