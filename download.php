<?php
require_once('vendor/autoload.php');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;

$spreadsheet = new Spreadsheet();

//Specify the properties for this document
$spreadsheet->getProperties()
    ->setTitle('PHP Download Example')
    ->setSubject('A PHPExcel example')
    ->setDescription('A simple example for PhpSpreadsheet. This class replaces the PHPExcel class')
    ->setCreator('php.com')
    ->setLastModifiedBy('php.com');
$i = 1;
$spreadsheet->setActiveSheetIndex(0)
    ->setCellValue('A'.$i, 'id')
    ->setCellValue('B'.$i, 'class')
    ->setCellValue('C'.$i, 'order')
    ->setCellValue('D'.$i, 'title')
    ->setCellValue('E'.$i, 'picture')
    ->setCellValue('F'.$i, 'old_id');

$configSubjects = array();
$configSubjects[1] = array();
$configSubjects[1][251] = array("id" => 1, "title" => "Tự nhiên và xã hội", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 251);
$configSubjects[1][246] = array("id" => 2, "title" => "Đạo đức", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 246);
$configSubjects[1][72] = array("id" => 3, "title" => "Tiếng Anh", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 72);
$configSubjects[1][40] = array("id" => 4, "title" => "Toán học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 40);
$configSubjects[2] = array();
$configSubjects[2][251] = array("id" => 5, "title" => "Tự nhiên và xã hội", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 251);
$configSubjects[2][29] = array("id" => 6, "title" => "Ngữ văn", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 29);
$configSubjects[2][40] = array("id" => 7, "title" => "Toán học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 40);
$configSubjects[2][72] = array("id" => 8, "title" => "Tiếng Anh", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 72);
$configSubjects[3] = array();
$configSubjects[3][251] = array("id" => 9, "title" => "Tự nhiên và xã hội", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 251);
$configSubjects[3][246] = array("id" => 10, "title" => "Đạo đức", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 246);
$configSubjects[3][29] = array("id" => 11, "title" => "Ngữ văn", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 29);
$configSubjects[3][72] = array("id" => 12, "title" => "Tiếng Anh", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 72);
$configSubjects[3][40] = array("id" => 13, "title" => "Toán học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 40);
$configSubjects[4] = array();
$configSubjects[4][80] = array("id" => 14, "title" => "Lịch sử", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 80);
$configSubjects[4][88] = array("id" => 15, "title" => "Địa lí", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 88);
$configSubjects[4][175] = array("id" => 16, "title" => "Khoa học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 175);
$configSubjects[4][29] = array("id" => 17, "title" => "Ngữ văn", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 29);
$configSubjects[4][40] = array("id" => 18, "title" => "Toán học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 40);
$configSubjects[4][153] = array("id" => 19, "title" => "Tin học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 153);
$configSubjects[4][246] = array("id" => 20, "title" => "Đạo đức", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 246);
$configSubjects[4][72] = array("id" => 21, "title" => "Tiếng Anh", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 72);
$configSubjects[5] = array();
$configSubjects[5][80] = array("id" => 22, "title" => "Lịch sử", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 80);
$configSubjects[5][88] = array("id" => 23, "title" => "Địa lí", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 88);
$configSubjects[5][175] = array("id" => 24, "title" => "Khoa học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 175);
$configSubjects[5][40] = array("id" => 25, "title" => "Toán học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 40);
$configSubjects[5][29] = array("id" => 26, "title" => "Ngữ văn", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 29);
$configSubjects[5][153] = array("id" => 27, "title" => "Tin học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 153);
$configSubjects[5][72] = array("id" => 28, "title" => "Tiếng Anh", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 72);
$configSubjects[5][246] = array("id" => 29, "title" => "Đạo đức", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 246);
$configSubjects[6] = array();
$configSubjects[6][166] = array("id" => 30, "title" => "Công nghệ", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 166);
$configSubjects[6][153] = array("id" => 31, "title" => "Tin học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 153);
$configSubjects[6][443] = array("id" => 32, "title" => "Âm nhạc và mỹ thuật", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 443);
$configSubjects[6][611] = array("id" => 33, "title" => "Lịch sử và Địa lí", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 611);
$configSubjects[6][144] = array("id" => 34, "title" => "GDCD", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 144);
$configSubjects[6][29] = array("id" => 35, "title" => "Ngữ văn", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 29);
$configSubjects[6][40] = array("id" => 36, "title" => "Toán học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 40);
$configSubjects[6][612] = array("id" => 37, "title" => "Khoa học tự nhiên", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 612);
$configSubjects[6][72] = array("id" => 38, "title" => "Tiếng Anh", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 72);
$configSubjects[7] = array();
$configSubjects[7][72] = array("id" => 39, "title" => "Tiếng Anh", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 72);
$configSubjects[7][144] = array("id" => 40, "title" => "GDCD", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 144);
$configSubjects[7][56] = array("id" => 41, "title" => "Vật lí", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 56);
$configSubjects[7][40] = array("id" => 42, "title" => "Toán học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 40);
$configSubjects[7][166] = array("id" => 43, "title" => "Công nghệ", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 166);
$configSubjects[7][64] = array("id" => 44, "title" => "Sinh học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 64);
$configSubjects[7][153] = array("id" => 45, "title" => "Tin học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 153);
$configSubjects[7][443] = array("id" => 46, "title" => "Âm nhạc và mỹ thuật", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 443);
$configSubjects[7][80] = array("id" => 47, "title" => "Lịch sử", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 80);
$configSubjects[7][88] = array("id" => 48, "title" => "Địa lí", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 88);
$configSubjects[7][29] = array("id" => 49, "title" => "Ngữ văn", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 29);
$configSubjects[8] = array();
$configSubjects[8][72] = array("id" => 50, "title" => "Tiếng Anh", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 72);
$configSubjects[8][88] = array("id" => 51, "title" => "Địa lí", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 88);
$configSubjects[8][144] = array("id" => 52, "title" => "GDCD", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 144);
$configSubjects[8][56] = array("id" => 53, "title" => "Vật lí", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 56);
$configSubjects[8][40] = array("id" => 54, "title" => "Toán học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 40);
$configSubjects[8][166] = array("id" => 55, "title" => "Công nghệ", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 166);
$configSubjects[8][29] = array("id" => 56, "title" => "Ngữ văn", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 29);
$configSubjects[8][64] = array("id" => 57, "title" => "Sinh học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 64);
$configSubjects[8][153] = array("id" => 58, "title" => "Tin học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 153);
$configSubjects[8][443] = array("id" => 59, "title" => "Âm nhạc và mỹ thuật", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 443);
$configSubjects[8][80] = array("id" => 60, "title" => "Lịch sử", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 80);
$configSubjects[8][50] = array("id" => 61, "title" => "Hóa học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 50);
$configSubjects[9] = array();
$configSubjects[9][50] = array("id" => 62, "title" => "Hóa học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 50);
$configSubjects[9][88] = array("id" => 63, "title" => "Địa lí", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 88);
$configSubjects[9][72] = array("id" => 64, "title" => "Tiếng Anh", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 72);
$configSubjects[9][144] = array("id" => 65, "title" => "GDCD", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 144);
$configSubjects[9][40] = array("id" => 66, "title" => "Toán học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 40);
$configSubjects[9][166] = array("id" => 67, "title" => "Công nghệ", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 166);
$configSubjects[9][153] = array("id" => 68, "title" => "Tin học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 153);
$configSubjects[9][64] = array("id" => 69, "title" => "Sinh học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 64);
$configSubjects[9][443] = array("id" => 70, "title" => "Âm nhạc và mỹ thuật", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 443);
$configSubjects[9][29] = array("id" => 71, "title" => "Ngữ văn", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 29);
$configSubjects[9][56] = array("id" => 72, "title" => "Vật lí", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 56);
$configSubjects[9][80] = array("id" => 73, "title" => "Lịch sử", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 80);
$configSubjects[10] = array();
$configSubjects[10][50] = array("id" => 74, "title" => "Hóa học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 50);
$configSubjects[10][72] = array("id" => 75, "title" => "Tiếng Anh", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 72);
$configSubjects[10][80] = array("id" => 76, "title" => "Lịch sử", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 80);
$configSubjects[10][64] = array("id" => 77, "title" => "Sinh học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 64);
$configSubjects[10][88] = array("id" => 78, "title" => "Địa lí", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 88);
$configSubjects[10][56] = array("id" => 79, "title" => "Vật lí", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 56);
$configSubjects[10][153] = array("id" => 80, "title" => "Tin học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 153);
$configSubjects[10][40] = array("id" => 81, "title" => "Toán học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 40);
$configSubjects[10][144] = array("id" => 82, "title" => "GDCD", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 144);
$configSubjects[10][166] = array("id" => 83, "title" => "Công nghệ", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 166);
$configSubjects[10][29] = array("id" => 84, "title" => "Ngữ văn", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 29);
$configSubjects[11] = array();
$configSubjects[11][50] = array("id" => 85, "title" => "Hóa học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 50);
$configSubjects[11][72] = array("id" => 86, "title" => "Tiếng Anh", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 72);
$configSubjects[11][56] = array("id" => 87, "title" => "Vật lí", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 56);
$configSubjects[11][153] = array("id" => 88, "title" => "Tin học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 153);
$configSubjects[11][144] = array("id" => 89, "title" => "GDCD", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 144);
$configSubjects[11][40] = array("id" => 90, "title" => "Toán học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 40);
$configSubjects[11][88] = array("id" => 91, "title" => "Địa lí", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 88);
$configSubjects[11][166] = array("id" => 92, "title" => "Công nghệ", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 166);
$configSubjects[11][80] = array("id" => 93, "title" => "Lịch sử", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 80);
$configSubjects[11][29] = array("id" => 94, "title" => "Ngữ văn", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 29);
$configSubjects[11][64] = array("id" => 95, "title" => "Sinh học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 64);
$configSubjects[12] = array();
$configSubjects[12][64] = array("id" => 96, "title" => "Sinh học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 64);
$configSubjects[12][72] = array("id" => 97, "title" => "Tiếng Anh", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 72);
$configSubjects[12][50] = array("id" => 98, "title" => "Hóa học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 50);
$configSubjects[12][144] = array("id" => 99, "title" => "GDCD", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 144);
$configSubjects[12][40] = array("id" => 100, "title" => "Toán học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 40);
$configSubjects[12][88] = array("id" => 101, "title" => "Địa lí", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 88);
$configSubjects[12][153] = array("id" => 102, "title" => "Tin học", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 153);
$configSubjects[12][166] = array("id" => 103, "title" => "Công nghệ", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 166);
$configSubjects[12][80] = array("id" => 104, "title" => "Lịch sử", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 80);
$configSubjects[12][29] = array("id" => 105, "title" => "Ngữ văn", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 29);
$configSubjects[12][56] = array("id" => 106, "title" => "Vật lí", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 56);
$configSubjects[16] = array();
$configSubjects[16][123] = array("id" => 107, "title" => "Môn Đại Cương", "picture" => "https://img.loigiaihay.com/picture/2016/0712/monvan-0.png", "old_id" => 123);
//Adding data to the excel sheet

$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(10);
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(50);
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(50);

$styleArray = array(
    'font'  => array(
        'bold'  => true,
        'color' => array('rgb' => '000000'),
        'size'  => 12,
        'name'  => 'Verdana',
    ),
    'fill' => array(
        'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
        'startColor' => array('argb' => 'FF4F81BD')
    ));

$spreadsheet->getActiveSheet()->getStyle('A1:F1')->applyFromArray($styleArray);


$i = 2;
$currentClass = 0;
foreach($configSubjects as $class => $configSubjectItem) {

    if ($currentClass != $class) {
        $currentClass = $class;

        $spreadsheet->setActiveSheetIndex(0)
            ->setCellValue('A'.$i, '')
            ->setCellValue('B'.$i, '')
            ->setCellValue('C'.$i, '')
            ->setCellValue('D'.$i, '')
            ->setCellValue('E'.$i, '')
            ->setCellValue('F'.$i, '');

        $i++;
    }

    $order = 1;
    foreach ($configSubjectItem as $subject) {
        $spreadsheet->setActiveSheetIndex(0)
            ->setCellValue('A'.$i, $subject["id"])
            ->setCellValue('B'.$i, $class)
            ->setCellValue('C'.$i, $order)
            ->setCellValue('D'.$i, $subject["title"])
            ->setCellValue('E'.$i, $subject["picture"])
            ->setCellValue('F'.$i, $subject["old_id"]);
        $i++;
        $order++;
    }
}


$fileName = "struc-".time().'.xlsx';
$writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="'. urlencode($fileName).'"');
$writer->save('php://output');