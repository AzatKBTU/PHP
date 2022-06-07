<?php
$_SERVER["DOCUMENT_ROOT"] = "/home/bitrix/www";
$DOCUMENT_ROOT = $_SERVER["DOCUMENT_ROOT"];

define("NO_KEEP_STATISTIC", true);
define("NOT_CHECK_PERMISSIONS", true);
require($_SERVER["DOCUMENT_ROOT"]."/bitrix/modules/main/include/prolog_before.php");
set_time_limit(0);

global $USER;
global $DB;
echo "tes";
CModule::IncludeModule("nkhost.phpexcel");
global $PHPEXCELPATH;
require_once($PHPEXCELPATH.'/PHPExcel.php');
require_once($PHPEXCELPATH.'/PHPExcel/Writer/Excel5.php');
echo "ues";
echo "\n";

CModule::IncludeModule("iblock");
CModule::IncludeModule("crm");
CModule::IncludeModule("tasks");
CModule::IncludeModule("sale");
CModule::IncludeModule("socialnetwork");

$objPHPExcel = new PHPExcel();
$objPHPExcel->setActiveSheetIndex(0);
$rowCount = 2;
$sheet = $objPHPExcel->getActiveSheet();
$sheet->setCellValue("A2", 'месяц выдачи');
$sheet->setCellValue("B2", 'сумма');
$sheet->setCellValue("C2", 'доля по сумме');
$sheet->setCellValue("D2", 'количество');
$sheet->setCellValue("E2", 'доля по количеству');
$sheet->setCellValue("H2", 'сумма');
$sheet->setCellValue("I2", 'доля по сумме');
$sheet->setCellValue("J2", 'количество');
$sheet->setCellValue("K2", 'доля по количеству');
$sheet->setCellValue("L2", 'сумма');
$sheet->setCellValue("M2", 'доля по сумме');
$sheet->setCellValue("N2", 'количество');
$sheet->setCellValue("O2", 'доля по количеству');
$sheet->setCellValue("P2", 'сумма');
$sheet->setCellValue("Q2", 'доля по сумме');
$sheet->setCellValue("R2", 'количество');
$sheet->setCellValue("S2", 'доля по количеству');
$sheet->setCellValue("T2", 'сумма');
$sheet->setCellValue("U2", 'доля по сумме');
$sheet->setCellValue("V2", 'количество');
$sheet->setCellValue("W2", 'доля по количеству');
$sheet->setCellValue("X2", 'сумма');
$sheet->setCellValue("Y2", 'количество');
$sheet->setCellValue("Z2", 'сумма');
$sheet->setCellValue("AA2", 'количество');

$sheet->setCellValue("A3", 'Январь 2022');
$sheet->setCellValue("A4", 'Февраль 2022');
$sheet->setCellValue("A5", 'Март 2022');
$sheet->setCellValue("A6", 'Апрель 2022');
$sheet->setCellValue("A7", 'Май 2022');
$sheet->setCellValue("A8", 'Июнь 2022');
$sheet->setCellValue("A9", 'Июль 2022');
$sheet->setCellValue("A10", 'Август 2022');
$sheet->setCellValue("A11", 'Сентябрь 2022');
$sheet->setCellValue("A12", 'Октябрь 2022');
$sheet->setCellValue("A13", 'Ноябрь 2022');
$sheet->setCellValue("A14", 'Декабрь 2022');

$objPHPExcel->getActiveSheet()->getStyle("A1:AA2")->getFont()->setBold(true);
$sheet->mergeCells('B1:E1');
$sheet->setCellValue("B1", 'стандартный займ');
$sheet->mergeCells('H1:K1');
$sheet->setCellValue("H1", 'реструктуризация');
$sheet->mergeCells('L1:O1');
$sheet->setCellValue("L1", 'просрочен софт');
$sheet->mergeCells('P1:S1');
$sheet->setCellValue("P1", 'просрочен хард');
$sheet->mergeCells('T1:W1');
$sheet->setCellValue("T1", 'погашен');
$sheet->mergeCells('X1:Y1');
$sheet->setCellValue("X1", 'Общий итог');
$sheet->mergeCells('Z1:AA1');
$sheet->setCellValue("Z1", 'Продажа');

$sheet->setTitle('REPORT');

$sheet->getDefaultStyle()->getFont()->setName('Times New Roman');
$sheet->getDefaultStyle()->getFont()->setSize(10);

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(13);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(9);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(14);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(11);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(18);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(9);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(14);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(11);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(18);
$objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth(9);
$objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth(14);
$objPHPExcel->getActiveSheet()->getColumnDimension('N')->setWidth(11);
$objPHPExcel->getActiveSheet()->getColumnDimension('O')->setWidth(18);
$objPHPExcel->getActiveSheet()->getColumnDimension('P')->setWidth(9);
$objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setWidth(14);
$objPHPExcel->getActiveSheet()->getColumnDimension('R')->setWidth(11);
$objPHPExcel->getActiveSheet()->getColumnDimension('S')->setWidth(18);
$objPHPExcel->getActiveSheet()->getColumnDimension('T')->setWidth(9);
$objPHPExcel->getActiveSheet()->getColumnDimension('U')->setWidth(14);
$objPHPExcel->getActiveSheet()->getColumnDimension('V')->setWidth(11);
$objPHPExcel->getActiveSheet()->getColumnDimension('W')->setWidth(18);
$objPHPExcel->getActiveSheet()->getColumnDimension('X')->setWidth(11);
$objPHPExcel->getActiveSheet()->getColumnDimension('Y')->setWidth(11);
$objPHPExcel->getActiveSheet()->getColumnDimension('Z')->setWidth(11);
$objPHPExcel->getActiveSheet()->getColumnDimension('AA')->setWidth(11);

$style = array(
    'alignment' => array(
        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
        'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
    )
);
$sheet->getStyle("A1:AA14")->applyFromArray($style);

$border = array(
	'borders'=>array(
		'outline' => array(
			'style' => PHPExcel_Style_Border::BORDER_THICK,
			'color' => array('rgb' => '000000')
		),
		'allborders' => array(
			'style' => PHPExcel_Style_Border::BORDER_THIN,
			'color' => array('rgb' => '000000')
		)
	)
);
$sheet->getStyle("A1:AA14")->applyFromArray($border);

$bg = array(
	'fill' => array(
		'type' => PHPExcel_Style_Fill::FILL_SOLID,
		'color' => array('rgb' => 'BDD7EE')
	)
);
$sheet->getStyle("A1:AA14")->applyFromArray($bg);

echo $a;
echo "\n";
$day = '01';
$month = '01';
$year = '2022';
$rowCount = 3;
$now = date("Y-m-d");

for ($i = 1; $i <= 12; $i++) {
	$date1 = $year.'-'.$month.'-'.$day;
	$a = date('t', mktime(0, 0, 0, $month, 1, $year));
	$day = intval($day)+$a-1;
	$date2 = $year.'-'.$month.'-'.$day;
	if($date1 <= $now){
		//Стандарт
		$sql = "SELECT COUNT(UF_CRM_1573668162) FROM b_crm_deal AS b JOIN b_uts_crm_deal AS bu ON b.ID = bu.VALUE_ID WHERE UF_CRM_1575338848 BETWEEN '".$date1."' AND '".$date2."' AND b.CATEGORY_ID = 2 AND (STAGE_ID = 'C2:PREPARATION' or STAGE_ID = 'C2:PREPAYMENT_INVOICE' or STAGE_ID = 'C2:14')";
		$result = $DB->Query($sql)->Fetch();
		foreach($result as $key=>$val) {
			$sheet->SetCellValue('D'.$rowCount, $val);
		}
		$sql = "SELECT UF_CRM_1589873735 FROM b_crm_deal AS b JOIN b_uts_crm_deal AS bu ON b.ID = bu.VALUE_ID WHERE UF_CRM_1575338848 BETWEEN '".$date1."' AND '".$date2."' AND b.CATEGORY_ID = 2 AND (STAGE_ID = 'C2:PREPARATION' or STAGE_ID = 'C2:PREPAYMENT_INVOICE' or STAGE_ID = 'C2:14')";
		$result = $DB->Query($sql);
		$money = 0;
		while ($s = $result->fetch()){
			$money += $s['UF_CRM_1589873735'];
		}
		$sheet->SetCellValue('B'.$rowCount, $money);
		//Ресструктур
		$sql = "SELECT COUNT(UF_CRM_1573668162) FROM b_crm_deal AS b JOIN b_uts_crm_deal AS bu ON b.ID = bu.VALUE_ID WHERE UF_CRM_1575338848 BETWEEN '".$date1."' AND '".$date2."' AND b.CATEGORY_ID = 2 AND (STAGE_ID = 'C2:1' or STAGE_ID = 'C2:6')";
		$result = $DB->Query($sql)->Fetch();
		foreach($result as $key=>$val) {
			$sheet->SetCellValue('J'.$rowCount, $val);
		}
		$sql = "SELECT UF_CRM_1589873735 FROM b_crm_deal AS b JOIN b_uts_crm_deal AS bu ON b.ID = bu.VALUE_ID WHERE UF_CRM_1575338848 BETWEEN '".$date1."' AND '".$date2."' AND b.CATEGORY_ID = 2 AND (STAGE_ID = 'C2:1' or STAGE_ID = 'C2:6')";
		$result = $DB->Query($sql);
		$money = 0;
		while ($s = $result->fetch()){
			$money += $s['UF_CRM_1589873735'];
		}
		$sheet->SetCellValue('H'.$rowCount, $money);
		//Софт
		$sql = "SELECT COUNT(UF_CRM_1573668162) FROM b_crm_deal AS b JOIN b_uts_crm_deal AS bu ON b.ID = bu.VALUE_ID WHERE UF_CRM_1575338848 BETWEEN '".$date1."' AND '".$date2."' AND b.CATEGORY_ID = 2 AND STAGE_ID = 'C2:2'";
		$result = $DB->Query($sql)->Fetch();
		foreach($result as $key=>$val) {
			$sheet->SetCellValue('N'.$rowCount, $val);
		}
		$sql = "SELECT UF_CRM_1589873735 FROM b_crm_deal AS b JOIN b_uts_crm_deal AS bu ON b.ID = bu.VALUE_ID WHERE UF_CRM_1575338848 BETWEEN '".$date1."' AND '".$date2."' AND b.CATEGORY_ID = 2 AND STAGE_ID = 'C2:2'";
		$result = $DB->Query($sql);
		$money = 0;
		while ($s = $result->fetch()){
			$money += $s['UF_CRM_1589873735'];
		}
		$sheet->SetCellValue('L'.$rowCount, $money);
		//Хард
		$sql = "SELECT COUNT(UF_CRM_1573668162) FROM b_crm_deal AS b JOIN b_uts_crm_deal AS bu ON b.ID = bu.VALUE_ID WHERE UF_CRM_1575338848 BETWEEN '".$date1."' AND '".$date2."' AND b.CATEGORY_ID = 2 AND (STAGE_ID = 'C2:3' or STAGE_ID = 'C2:7' or STAGE_ID = 'C2:8' or STAGE_ID = 'C2:9' or STAGE_ID = 'C2:10' or STAGE_ID = 'C2:13' or STAGE_ID = 'C2:LOSE')";
		$result = $DB->Query($sql)->Fetch();
		foreach($result as $key=>$val) {
			$sheet->SetCellValue('R'.$rowCount, $val);
		}
		$sql = "SELECT UF_CRM_1589873735 FROM b_crm_deal AS b JOIN b_uts_crm_deal AS bu ON b.ID = bu.VALUE_ID WHERE UF_CRM_1575338848 BETWEEN '".$date1."' AND '".$date2."' AND b.CATEGORY_ID = 2 AND (STAGE_ID = 'C2:3' or STAGE_ID = 'C2:7' or STAGE_ID = 'C2:8' or STAGE_ID = 'C2:9' or STAGE_ID = 'C2:10' or STAGE_ID = 'C2:13' or STAGE_ID = 'C2:LOSE')";
		$result = $DB->Query($sql);
		$money = 0;
		while ($s = $result->fetch()){
			$money += $s['UF_CRM_1589873735'];
		}
		$sheet->SetCellValue('P'.$rowCount, $money);
		//Погашен
		$sql = "SELECT COUNT(UF_CRM_1573668162) FROM b_crm_deal AS b JOIN b_uts_crm_deal AS bu ON b.ID = bu.VALUE_ID WHERE UF_CRM_1575338848 BETWEEN '".$date1."' AND '".$date2."' AND b.CATEGORY_ID = 2 AND STAGE_ID = 'C2:WON'";
		$result = $DB->Query($sql)->Fetch();
		foreach($result as $key=>$val) {
			$sheet->SetCellValue('V'.$rowCount, $val);
		}
		$sql = "SELECT UF_CRM_1589873735 FROM b_crm_deal AS b JOIN b_uts_crm_deal AS bu ON b.ID = bu.VALUE_ID WHERE UF_CRM_1575338848 BETWEEN '".$date1."' AND '".$date2."' AND b.CATEGORY_ID = 2 AND STAGE_ID = 'C2:WON'";
		$result = $DB->Query($sql);
		$money = 0;
		while ($s = $result->fetch()){
			$money += $s['UF_CRM_1589873735'];
		}
		$sheet->SetCellValue('T'.$rowCount, $money);
		//Продажа
		$sql = "SELECT COUNT(UF_CRM_1573668162) FROM b_crm_deal AS b JOIN b_uts_crm_deal AS bu ON b.ID = bu.VALUE_ID WHERE UF_CRM_1575338848 BETWEEN '".$date1."' AND '".$date2."' AND b.CATEGORY_ID = 2 AND (STAGE_ID = 'C2:11' OR STAGE_ID = 'C2:12')";
		$result = $DB->Query($sql)->Fetch();
		foreach($result as $key=>$val) {
			$sheet->SetCellValue('AA'.$rowCount, $val);
		}
		$sql = "SELECT UF_CRM_1589873735 FROM b_crm_deal AS b JOIN b_uts_crm_deal AS bu ON b.ID = bu.VALUE_ID WHERE UF_CRM_1575338848 BETWEEN '".$date1."' AND '".$date2."' AND b.CATEGORY_ID = 2 AND (STAGE_ID = 'C2:11' OR STAGE_ID = 'C2:12')";
		$result = $DB->Query($sql);
		$money = 0;
		while ($s = $result->fetch()){
			$money += $s['UF_CRM_1589873735'];
		}
		$sheet->SetCellValue('Z'.$rowCount, $money);

		$sheet->SetCellValue('C'.$rowCount, '=B'.$rowCount.'/X'.$rowCount);
		$sheet->SetCellValue('E'.$rowCount, '=D'.$rowCount.'/Y'.$rowCount);
		$sheet->SetCellValue('I'.$rowCount, '=H'.$rowCount.'/X'.$rowCount);
		$sheet->SetCellValue('K'.$rowCount, '=J'.$rowCount.'/Y'.$rowCount);
		$sheet->SetCellValue('M'.$rowCount, '=L'.$rowCount.'/X'.$rowCount);
		$sheet->SetCellValue('O'.$rowCount, '=N'.$rowCount.'/Y'.$rowCount);
		$sheet->SetCellValue('Q'.$rowCount, '=P'.$rowCount.'/X'.$rowCount);
		$sheet->SetCellValue('S'.$rowCount, '=R'.$rowCount.'/Y'.$rowCount);
		$sheet->SetCellValue('U'.$rowCount, '=T'.$rowCount.'/X'.$rowCount);
		$sheet->SetCellValue('W'.$rowCount, '=V'.$rowCount.'/Y'.$rowCount);

		$sheet->SetCellValue('X'.$rowCount, '=B'.$rowCount.'+H'.$rowCount.'+L'.$rowCount.'+P'.$rowCount.'+T'.$rowCount);
		$sheet->SetCellValue('Y'.$rowCount, '=D'.$rowCount.'+J'.$rowCount.'+N'.$rowCount.'+R'.$rowCount.'+V'.$rowCount);
	}

	var_dump($result);
	$day = '01';
	$lol = intval($month) + 1;
	if($lol < 10){
		$month = '0'.$lol;
	}else{
		$month = $lol;
	}
	$year = '2022';
	$rowCount++;
}
$objPHPExcel->getActiveSheet()->getStyle('C3:C14')->getNumberFormat()->applyFromArray(array('code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00));
$objPHPExcel->getActiveSheet()->getStyle('E3:E14')->getNumberFormat()->applyFromArray(array('code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00));
$objPHPExcel->getActiveSheet()->getStyle('I3:I14')->getNumberFormat()->applyFromArray(array('code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00));
$objPHPExcel->getActiveSheet()->getStyle('K3:K14')->getNumberFormat()->applyFromArray(array('code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00));
$objPHPExcel->getActiveSheet()->getStyle('M3:M14')->getNumberFormat()->applyFromArray(array('code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00));
$objPHPExcel->getActiveSheet()->getStyle('O3:O14')->getNumberFormat()->applyFromArray(array('code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00));
$objPHPExcel->getActiveSheet()->getStyle('Q3:Q14')->getNumberFormat()->applyFromArray(array('code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00));
$objPHPExcel->getActiveSheet()->getStyle('S3:S14')->getNumberFormat()->applyFromArray(array('code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00));
$objPHPExcel->getActiveSheet()->getStyle('U3:U14')->getNumberFormat()->applyFromArray(array('code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00));
$objPHPExcel->getActiveSheet()->getStyle('W3:W14')->getNumberFormat()->applyFromArray(array('code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00));

$today = date("d.m.y");  

$objWriter = new PHPExcel_Writer_Excel5($objPHPExcel);
$fileName = "Report_c_01.01.22_до_".$today.".xls";
$uploadDir = '/upload/report';

if (!is_dir($_SERVER["DOCUMENT_ROOT"] . $uploadDir)) {
    mkdir($_SERVER["DOCUMENT_ROOT"] . $uploadDir, 0664);
}

$fileFullName = $_SERVER["DOCUMENT_ROOT"] . $uploadDir . '/' . $fileName;
$objWriter->save($fileFullName);
?>