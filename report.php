<?php
$_SERVER["DOCUMENT_ROOT"] = "/home/bitrix/www";
$DOCUMENT_ROOT = $_SERVER["DOCUMENT_ROOT"];

define("NO_KEEP_STATISTIC", true);
define("NOT_CHECK_PERMISSIONS", true);
require($_SERVER["DOCUMENT_ROOT"]."/bitrix/modules/main/include/prolog_before.php");
set_time_limit(0);

global $USER;
global $DB;
echo "test";
CModule::IncludeModule("nkhost.phpexcel");
global $PHPEXCELPATH;
require_once($PHPEXCELPATH.'/PHPExcel.php');
require_once($PHPEXCELPATH.'/PHPExcel/Writer/Excel5.php');
echo "test1";

CModule::IncludeModule("iblock");
CModule::IncludeModule("crm");
CModule::IncludeModule("tasks");
CModule::IncludeModule("sale");
CModule::IncludeModule("socialnetwork");

$date1 = date("Y-m-d", time()-(60*60*24));
$date = date("Y-m-d", time()-(60*60*24*30));
$fileName = "Report_c_".$date."_до_".$date1.".xls";

$objPHPExcel = new PHPExcel();
$objPHPExcel->setActiveSheetIndex(0);
$rowCount = 2;
$sheet = $objPHPExcel->getActiveSheet();
$tt = date('d.m.Y',strtotime($date));
$sheet->setCellValue("A1", 'Period');
$sheet->getStyle("A1")->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->setCellValue("B1", 'Lead');
$sheet->getStyle("B1")->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->setCellValue("C1", 'Issued');
$sheet->getStyle("C1")->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->setCellValue("D1", 'Canceled');
$sheet->getStyle("D1")->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->setCellValue("E1", 'Rejected');
$sheet->getStyle("E1")->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->setCellValue("F1", 'Extensions');
$sheet->getStyle("F1")->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->setCellValue("G1", 'Closed');
$sheet->getStyle("G1")->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);

$objPHPExcel->getActiveSheet()->getStyle("A1:G1")->getFont()->setBold(true);

$used_days = strtotime($date1) - strtotime($date);
$used_days = $used_days/3600/24;
echo $used_days."\n";


$sheet->setTitle('REPORT');

$sheet->getDefaultStyle()->getFont()->setName('Times New Roman');
$sheet->getDefaultStyle()->getFont()->setSize(10);

for ($i = 0; $i <= $used_days; $i++) {
    $sql = "SELECT COUNT(b.ID) FROM b_crm_lead AS b JOIN b_uts_crm_lead AS bu ON b.ID = bu.VALUE_ID WHERE DATE_CREATE LIKE '".$date."%'";
    $result = $DB->Query($sql)->Fetch();
    foreach($result as $key=>$val) {
        $sheet->SetCellValue('B'.$rowCount, $val);
    }
    $sql = "SELECT COUNT(UF_CRM_1573668162) FROM b_crm_deal AS b JOIN b_uts_crm_deal AS bu ON b.ID = bu.VALUE_ID WHERE bu.UF_CRM_1575338848 = '".$date."' AND b.CATEGORY_ID = 2";
    $result = $DB->Query($sql)->Fetch();
    foreach($result as $key=>$val) {
        $sheet->SetCellValue('C'.$rowCount, $val);
    }
    $sql = "SELECT COUNT(b.ID) FROM b_crm_lead AS b JOIN b_uts_crm_lead AS bu ON b.ID = bu.VALUE_ID WHERE DATE_CREATE LIKE '".$date."%' AND STATUS_ID = 3";
    $result = $DB->Query($sql)->Fetch();
    foreach($result as $key=>$val) {
        $sheet->SetCellValue('D' . $rowCount, $val);
    }
    $sql = "SELECT COUNT(b.ID) FROM b_crm_lead AS b JOIN b_uts_crm_lead AS bu ON b.ID = bu.VALUE_ID WHERE DATE_CREATE LIKE '".$date."%' AND STATUS_ID = 'JUNK'";
    $result = $DB->Query($sql)->Fetch();

    foreach($result as $key=>$val) {
        $sheet->SetCellValue('E'.$rowCount, $val);
    }
    $sql = "SELECT COUNT(UF_CRM_1573668162) FROM b_crm_deal AS b JOIN b_uts_crm_deal AS bu ON b.ID = bu.VALUE_ID WHERE UF_CRM_1613644897 = '".$date."'";
    $result = $DB->Query($sql)->Fetch();
    foreach($result as $key=>$val) {
        $sheet->SetCellValue('F'.$rowCount, $val);
    }
    $sql = "SELECT COUNT(UF_CRM_1573668162) FROM b_crm_deal AS b JOIN b_uts_crm_deal AS bu ON b.ID = bu.VALUE_ID WHERE UF_CRM_1610348437 = '".$date."'";
    $result = $DB->Query($sql)->Fetch();
    foreach($result as $key=>$val) {
        $sheet->SetCellValue('G'.$rowCount, $val);
    }
    $a = PHPExcel_Shared_Date::PHPToExcel($date);
    $sheet->SetCellValue('A'.$rowCount, $a);
    $sheet->getStyle('A'.$rowCount)->getNumberFormat()->setFormatCode('dd/mm/yyyy'); //date
    $date = date("Y-m-d", strtotime($date)+(60*60*24));
    $rowCount++;
}
echo $rowCount;
echo 'lol';

$objWriter = new PHPExcel_Writer_Excel5($objPHPExcel);
$uploadDir = '/upload/report';

if (!is_dir($_SERVER["DOCUMENT_ROOT"] . $uploadDir)) {
    mkdir($_SERVER["DOCUMENT_ROOT"] . $uploadDir, 0664);
}

$fileFullName = $_SERVER["DOCUMENT_ROOT"] . $uploadDir . '/' . $fileName;
$objWriter->save($fileFullName);

$fields = [
    "EMAIL_TO" => 'r.sharafutdinov@i-credit.kz',
    "TEXT" => 'Report',
];
$arr_file=Array(
    "name" => $fileName,
    "tmp_name" => $fileFullName,
    "type" => "",
    "old_file" => "",
    "del" => "Y",
    "MODULE_ID" => "iblock"
);
$fid = CFile::SaveFile($arr_file, "landings");
$file['FILE_ID'] = $fid;
$sPath = CFile::GetPath($fid);
$iFileId = \CFile::SaveFile(\CFile::MakeFileArray($sPath), "main");
$s = \CEvent::Send("IMAGE_FEEDBACK", 's1', $fields,'Y',8,[$iFileId]);
var_dump($s);
?>