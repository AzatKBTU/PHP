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
echo "ues1";

CModule::IncludeModule("iblock");
CModule::IncludeModule("crm");
CModule::IncludeModule("tasks");
CModule::IncludeModule("sale");
CModule::IncludeModule("socialnetwork");

$date = date("Y-m-d", time()-(60*60*24));
echo $date;
$date1 = date('d.m.Y',strtotime($date));
$sql = "SELECT 
			bu.UF_CRM_1573668162,
			bu.UF_CRM_1575338848,
			bu.UF_CRM_1575338914,
			bu.UF_CRM_1577671976335,
			bu.UF_CRM_1577672007511,
			bu.UF_CRM_1577672053366,
			bu.UF_CRM_1577672434509,
			bu.UF_CRM_1578082787898,
			bu.UF_CRM_5E9FBD498C467,
			bu.UF_CRM_5E9FBD4A162DE,
			bu.UF_CRM_1610348437,
			bu.UF_CRM_1622005475,
			bu.UF_CRM_1589873611,
			bu.UF_CRM_1589873735,
			bu.UF_CRM_1589873821 
			FROM 
			b_crm_deal AS b 
			JOIN 
			b_uts_crm_deal AS bu 
			ON 
			b.ID = bu.VALUE_ID 
			WHERE 
			(bu.UF_CRM_1575338848 = '$date' OR 
			bu.UF_CRM_1610348437 = '$date' OR 
			bu.UF_CRM_1610348437 IS NULL) AND 
			bu.UF_CRM_1575338848 >= '2021-10-01' AND b.CATEGORY_ID = 2";
echo $sql;
$result = $DB->Query($sql);
while ($s = $result->fetch()){
    set_time_limit(0);
    $endDate = $s['UF_CRM_1575338914'];
    $used_days = strtotime($date) - strtotime($endDate);
    $used_days = $used_days/3600/24;
    $amountInsurance = $s['UF_CRM_1589873821'];
    $amountOnHand = $s['UF_CRM_1589873735'];
    $od = $amountInsurance + $amountOnHand;
    $amountCredit = $amountOnHand + $amountInsurance * 2;
    $factPaymentDate = $s['UF_CRM_1610348437'];
    $idCardNumber = $s['UF_CRM_5E9FBD4A162DE'];
    $fatherName = $s['UF_CRM_1577672053366'];
    $gender = $s['UF_CRM_1577672434509'];
    $birthDate = $s['UF_CRM_5E9FBD498C467'];
    $street = $s['UF_CRM_1578082787898'];
    $parentContractNumber = $s['UF_CRM_1622005475'];
    if(empty($parentContractNumber)){
        $parentContractNumber = $s['UF_CRM_1589873611'];
    }
    if($street == ''){
        $street = '1';
    }
    if(intval(substr($birthDate, 6, 4)) >= 2010){
        $birthDate = substr($birthDate, 0, 6).'19'.substr($birthDate, 8, 2);
    }
    if($gender == 'Ж' or $gender == 'Ж '){
        $gender = 'F';
    }
    if($gender == 'М'){
        $gender = 'M';
    }
    if($fatherName == '-' or $fatherName == 'Нет' or $fatherName == 'Нету'){
        $fatherName = '';
    }
    if(strlen($idCardNumber) == 7){
        $idCardNumber = "00".$idCardNumber;
    }
    if(strlen($idCardNumber) == 8){
        $idCardNumber = "0".$idCardNumber;
    }
    if(strlen($idCardNumber) == 10){
        $idCardNumber = substr($idCardNumber,1);
    }
	//echo substr($factPaymentDate, 0, 10);
    if(substr($factPaymentDate, 0, 10) == $date){
        $date_out = date('d.m.Y', strtotime($s['UF_CRM_1610348437']));
        $upcomingPayments = 0;
        $amountOutstanding = 0;
        $daysOverdue = 0;
        $amountOverdue = 0;
        $contractPhase = 5;
        $contractStatus = 1;
    }
    else{
        $date_out = '';
        if($used_days <= 0){
            $upcomingPayments = 1;
            $amountOutstanding = $amountCredit;
            $daysOverdue = 0;
            $amountOverdue = 0;
            $contractPhase = 4;
            $contractStatus = 1;
        }
        else{
            $upcomingPayments = 0;
            $amountOutstanding = 0;
            $daysOverdue = $used_days;
            $amountOverdue = $amountCredit;
            $contractPhase = 4;
            if($used_days>=1 and $used_days<=6){
                $contractStatus = 10;
            }
            elseif($used_days>=7 and $used_days<=30){
                $contractStatus = 11;
            }
            elseif($used_days>=31 and $used_days<=60){
                $contractStatus = 12;
            }
            elseif($used_days>=61 and $used_days<=90){
                $contractStatus = 13;
            }
            elseif($used_days>=91 and $used_days<=360){
                $contractStatus = 15;
            }
            else{
                $contractStatus = 16;
            }
        }
    }
    $array[] = [
        'iin' => $s['UF_CRM_1573668162'],
        'start' => $s['UF_CRM_1575338848'],
        'end' => $s['UF_CRM_1575338914'],
        'surname' => $s['UF_CRM_1577671976335'],
        'name' => $s['UF_CRM_1577672007511'],
        'fatherName' => $fatherName,
        'gender' => $gender,
        'street' => $street,
        'birthDate' => $birthDate,
        'identityCardNumber' => $idCardNumber,
        'factPaymentDate' => $date_out,
        'parentContractNumber' => $parentContractNumber,
        'contractNumber' => $s['UF_CRM_1589873611'],
        'amountOnHand' => $s['UF_CRM_1589873735'],
        'amountInsurance' => $s['UF_CRM_1589873821'],
        'od' => $od,
        'amountCredit' => $amountCredit,
        'upcomingPayments' => $upcomingPayments,
        'amountOutstanding' => $amountOutstanding,
        'daysOverdue' => $daysOverdue,
        'amountOverdue' => $amountOverdue,
        'contractPhase' => $contractPhase,
        'contractStatus' => $contractStatus,
    ];
}

$objPHPExcel = new PHPExcel();
$objPHPExcel->setActiveSheetIndex(0);
$rowCount = 4;
$sheet = $objPHPExcel->getActiveSheet();
$tt = date('d.m.Y',strtotime($date));
$sheet->setCellValue("A1", 'КОНТРАКТ');
$sheet->getStyle("A1")->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->mergeCells('A1:AZ1');
$sheet->setCellValue("BC1", 'РОЛЬ ДЛЯ ФИЗ И ЮР ЛИЦА');
$sheet->getStyle("BC1")->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->setCellValue("BD1", 'ФИЗ ЛИЦО');
$sheet->getStyle("BD1")->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->mergeCells('BD1:CG1');
$bg1 = array(
    'fill' => array(
        'type' => PHPExcel_Style_Fill::FILL_SOLID,
        'color' => array('rgb' => 'CCFF66')
    )
);
$sheet->getStyle("A1:AZ1")->applyFromArray($bg1);
$bg2 = array(
    'fill' => array(
        'type' => PHPExcel_Style_Fill::FILL_SOLID,
        'color' => array('rgb' => 'FF99CC')
    )
);
$sheet->getStyle("BC1")->applyFromArray($bg2);
$bg3 = array(
    'fill' => array(
        'type' => PHPExcel_Style_Fill::FILL_SOLID,
        'color' => array('rgb' => '99CCFF')
    )
);
$sheet->getStyle("BD1:CG1")->applyFromArray($bg3);

$border1 = array(
    'borders'=>array(
        'outline' => array(
            'style' => PHPExcel_Style_Border::BORDER_THIN,
            'color' => array('rgb' => '000000')
        ),
    )
);
$sheet->getStyle("A1:CE3")->applyFromArray($border1);
$border2 = array(
    'borders'=>array(
        'inside' => array(
            'style' => PHPExcel_Style_Border::BORDER_THIN,
            'color' => array('rgb' => '000000')
        ),
    )
);
$sheet->getStyle("A1:CE3")->applyFromArray($border2);


$objPHPExcel->getActiveSheet()->getStyle("A1:CE1")->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle("A2:CE2")->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle("A3:CE3")->getFont()->setBold(true);

$sheet->SetCellValue('A2','Contract Type');
$sheet->SetCellValue('B2','ContractCode');
$sheet->SetCellValue('C2','AgreementNumber');
$sheet->SetCellValue('D2','FundingType');
$sheet->SetCellValue('E2','CreditPurpose2');
$sheet->SetCellValue('F2','CreditObject');
$sheet->SetCellValue('G2','ContractPhase');
$sheet->SetCellValue('H2','ContractStatus');
$sheet->SetCellValue('I2','ThirdPartyHolder');
$sheet->SetCellValue('J2','ApplicationDate');
$sheet->SetCellValue('K2','StartDate');
$sheet->SetCellValue('L2','EndDate');
$sheet->SetCellValue('M2','ActualDate');
$sheet->SetCellValue('N2','RealPaymentDate');
$sheet->SetCellValue('O2','SpecialRelationship');
$sheet->SetCellValue('P2','AnnualEffectiveRate');
$sheet->SetCellValue('Q2','NominalRate');
$sheet->SetCellValue('R2','AmountProvisions');
$sheet->SetCellValue('S2','AMOUNTPROVISIONS_CUR');
$sheet->SetCellValue('T2','LoanAcount');
$sheet->SetCellValue('U2','GracePrincipal');
$sheet->SetCellValue('V2','GracePay');
$sheet->SetCellValue('W2','PlaceOfDisbursement');
$sheet->SetCellValue('X2','Classification');
$sheet->SetCellValue('Y2','ParentContractcode');
$sheet->SetCellValue('Z2','ParentProvider');
$sheet->SetCellValue('AA2','ParentContractStatus');
$sheet->SetCellValue('AB2','ParentOperationDate');
$sheet->SetCellValue('AC2','ProlongationCount');
$sheet->SetCellValue('AD2','BranchLocation');
$sheet->SetCellValue('AE2','paymentMethod');
$sheet->SetCellValue('AF2','paymentPeriodId');
$sheet->SetCellValue('AG2','TotalAmount');
$sheet->SetCellValue('AH2','TotalAmountCur');
$sheet->SetCellValue('AI2','InstalmentAmount');
$sheet->SetCellValue('AJ2','INSTALMENTAMOUNT_CUR');
$sheet->SetCellValue('AK2','InstalmentCount');
$sheet->SetCellValue('AL2','CreditLimit');
$sheet->SetCellValue('AM2','CreditLimit_Cur');
$sheet->SetCellValue('AN2','FundingSource');
$sheet->SetCellValue('AO2','AvailableDate');
$sheet->SetCellValue('AP2','AccountingDate');
$sheet->SetCellValue('AQ2','OutstandingInstallmentCount');
$sheet->SetCellValue('AR2','OutstandingAmount');
$sheet->SetCellValue('AS2','OverdueInstallmentCount');
$sheet->SetCellValue('AT2','OverdueAmount');
$sheet->SetCellValue('AU2','Fine');
$sheet->SetCellValue('AV2','FINE_CUR');
$sheet->SetCellValue('AW2','Penalty');
$sheet->SetCellValue('AX2','Penalty_cur');
$sheet->SetCellValue('AY2','ProlongationStartDate');
$sheet->SetCellValue('AZ2','ProlongationEndDate');
$sheet->SetCellValue('BA2','AvailableLimit');
$sheet->SetCellValue('BB2','AvailableLimit_Cur');
$sheet->SetCellValue('BC2','SubjectRole');
$sheet->SetCellValue('BD2','Surname');
$sheet->SetCellValue('BE2','FirstName');
$sheet->SetCellValue('BF2','FathersName (при наличии)');
$sheet->SetCellValue('BG2','BirthSurname');
$sheet->SetCellValue('BH2','Gender');
$sheet->SetCellValue('BI2','Classification');
$sheet->SetCellValue('BJ2','Residency');
$sheet->SetCellValue('BK2','Education');
$sheet->SetCellValue('BL2','MaritalStatus');
$sheet->SetCellValue('BM2','DateOfBirth');
$sheet->SetCellValue('BN2','NegativeStatus');
$sheet->SetCellValue('BO2','Profession');
$sheet->SetCellValue('BP2','EconomyActivityGroup');
$sheet->SetCellValue('BQ2','EmploymentNature');
$sheet->SetCellValue('BR2','Citizenship');
$sheet->SetCellValue('BS2','Salary');
$sheet->SetCellValue('BT2','IIN');
$sheet->SetCellValue('BU2','MobilePhone');
$sheet->SetCellValue('BV2','DependantLt18');
$sheet->SetCellValue('BW2','DependantGt18');
$sheet->SetCellValue('BX2','Number');
$sheet->SetCellValue('BY2','RegistrationDate');
$sheet->SetCellValue('BZ2','IssueDate');
$sheet->SetCellValue('CA2','ExpirationDate');
$sheet->SetCellValue('CB2','IssuedBy');
$sheet->SetCellValue('CC2','DocTypeId');
$sheet->SetCellValue('CD2','katoId');
$sheet->SetCellValue('CE2','StreetName');

$sheet->getRowDimension("1")->setRowHeight(37);
$sheet->getRowDimension("2")->setRowHeight(50);
$sheet->getRowDimension("3")->setRowHeight(60);

$sheet->getColumnDimension('A')->setWidth(15);
$sheet->getColumnDimension('B')->setWidth(15);
$sheet->getColumnDimension('C')->setWidth(15);
$sheet->getColumnDimension('D')->setWidth(15);
$sheet->getColumnDimension('E')->setWidth(15);
$sheet->getColumnDimension('F')->setWidth(15);
$sheet->getColumnDimension('G')->setWidth(15);
$sheet->getColumnDimension('H')->setWidth(15);
$sheet->getColumnDimension('I')->setWidth(15);
$sheet->getColumnDimension('J')->setWidth(15);
$sheet->getColumnDimension('K')->setWidth(15);
$sheet->getColumnDimension('L')->setWidth(15);
$sheet->getColumnDimension('M')->setWidth(15);
$sheet->getColumnDimension('N')->setWidth(15);
$sheet->getColumnDimension('O')->setWidth(15);
$sheet->getColumnDimension('P')->setWidth(15);
$sheet->getColumnDimension('Q')->setWidth(15);
$sheet->getColumnDimension('R')->setWidth(15);
$sheet->getColumnDimension('S')->setWidth(15);
$sheet->getColumnDimension('T')->setWidth(15);
$sheet->getColumnDimension('U')->setWidth(15);
$sheet->getColumnDimension('V')->setWidth(15);
$sheet->getColumnDimension('W')->setWidth(15);
$sheet->getColumnDimension('X')->setWidth(15);
$sheet->getColumnDimension('Y')->setWidth(15);
$sheet->getColumnDimension('Z')->setWidth(15);
$sheet->getColumnDimension('AA')->setWidth(15);
$sheet->getColumnDimension('AB')->setWidth(15);
$sheet->getColumnDimension('AC')->setWidth(15);
$sheet->getColumnDimension('AD')->setWidth(15);
$sheet->getColumnDimension('AE')->setWidth(15);
$sheet->getColumnDimension('AF')->setWidth(15);
$sheet->getColumnDimension('AG')->setWidth(15);
$sheet->getColumnDimension('AH')->setWidth(15);
$sheet->getColumnDimension('AI')->setWidth(15);
$sheet->getColumnDimension('AJ')->setWidth(15);
$sheet->getColumnDimension('AK')->setWidth(15);
$sheet->getColumnDimension('AL')->setWidth(15);
$sheet->getColumnDimension('AM')->setWidth(15);
$sheet->getColumnDimension('AN')->setWidth(15);
$sheet->getColumnDimension('AO')->setWidth(15);
$sheet->getColumnDimension('AP')->setWidth(15);
$sheet->getColumnDimension('AQ')->setWidth(15);
$sheet->getColumnDimension('AR')->setWidth(15);
$sheet->getColumnDimension('AS')->setWidth(15);
$sheet->getColumnDimension('AT')->setWidth(15);
$sheet->getColumnDimension('AU')->setWidth(15);
$sheet->getColumnDimension('AV')->setWidth(15);
$sheet->getColumnDimension('AW')->setWidth(15);
$sheet->getColumnDimension('AX')->setWidth(15);
$sheet->getColumnDimension('AY')->setWidth(15);
$sheet->getColumnDimension('AZ')->setWidth(15);
$sheet->getColumnDimension('BA')->setWidth(15);
$sheet->getColumnDimension('BB')->setWidth(15);
$sheet->getColumnDimension('BC')->setWidth(15);
$sheet->getColumnDimension('BD')->setWidth(15);
$sheet->getColumnDimension('BE')->setWidth(15);
$sheet->getColumnDimension('BF')->setWidth(15);
$sheet->getColumnDimension('BG')->setWidth(15);
$sheet->getColumnDimension('BH')->setWidth(15);
$sheet->getColumnDimension('BI')->setWidth(15);
$sheet->getColumnDimension('BJ')->setWidth(15);
$sheet->getColumnDimension('BK')->setWidth(15);
$sheet->getColumnDimension('BL')->setWidth(15);
$sheet->getColumnDimension('BM')->setWidth(15);
$sheet->getColumnDimension('BN')->setWidth(15);
$sheet->getColumnDimension('BO')->setWidth(15);
$sheet->getColumnDimension('BP')->setWidth(15);
$sheet->getColumnDimension('BQ')->setWidth(15);
$sheet->getColumnDimension('BR')->setWidth(15);
$sheet->getColumnDimension('BS')->setWidth(15);
$sheet->getColumnDimension('BT')->setWidth(15);
$sheet->getColumnDimension('BU')->setWidth(15);
$sheet->getColumnDimension('BV')->setWidth(15);
$sheet->getColumnDimension('BW')->setWidth(15);
$sheet->getColumnDimension('BX')->setWidth(15);
$sheet->getColumnDimension('BY')->setWidth(15);
$sheet->getColumnDimension('BZ')->setWidth(15);
$sheet->getColumnDimension('CA')->setWidth(15);
$sheet->getColumnDimension('CB')->setWidth(15);
$sheet->getColumnDimension('CC')->setWidth(15);
$sheet->getColumnDimension('CD')->setWidth(15);
$sheet->getColumnDimension('CE')->setWidth(15);


$sheet->getStyle("A2:H3")->getFont()->getColor()->setRGB('ff0000');
$sheet->getStyle("K2:L3")->getFont()->getColor()->setRGB('ff0000');
$sheet->getStyle("N2:N3")->getFont()->getColor()->setRGB('ff0000');
$sheet->getStyle("P2:Q3")->getFont()->getColor()->setRGB('ff0000');
$sheet->getStyle("X2:X3")->getFont()->getColor()->setRGB('ff0000');
$sheet->getStyle("AE2:AK3")->getFont()->getColor()->setRGB('ff0000');
$sheet->getStyle("AP2:AT3")->getFont()->getColor()->setRGB('ff0000');
$sheet->getStyle("BC2:BF3")->getFont()->getColor()->setRGB('ff0000');
$sheet->getStyle("BH2:BJ3")->getFont()->getColor()->setRGB('ff0000');
$sheet->getStyle("BM2:BM3")->getFont()->getColor()->setRGB('ff0000');
$sheet->getStyle("BR2:BR3")->getFont()->getColor()->setRGB('ff0000');
$sheet->getStyle("BT2:BT3")->getFont()->getColor()->setRGB('ff0000');
$sheet->getStyle("BX2:BY3")->getFont()->getColor()->setRGB('ff0000');
$sheet->getStyle("CC2:CE3")->getFont()->getColor()->setRGB('ff0000');

$sheet->SetCellValue('A3','Contract Type');
$sheet->SetCellValue('B3','ContractCode');
$sheet->SetCellValue('C3','AgreementNumber');
$sheet->SetCellValue('D3','FundingType');
$sheet->SetCellValue('E3','CreditPurpose2');
$sheet->SetCellValue('F3','CreditObject');
$sheet->SetCellValue('G3','ContractPhase');
$sheet->SetCellValue('H3','ContractStatus');
$sheet->SetCellValue('I3','ThirdPartyHolder');
$sheet->SetCellValue('J3','ApplicationDate');
$sheet->SetCellValue('K3','StartDate');
$sheet->SetCellValue('L3','EndDate');
$sheet->SetCellValue('M3','ActualDate');
$sheet->SetCellValue('N3','RealPaymentDate');
$sheet->SetCellValue('O3','SpecialRelationship');
$sheet->SetCellValue('P3','AnnualEffectiveRate');
$sheet->SetCellValue('Q3','NominalRate');
$sheet->SetCellValue('R3','AmountProvisions');
$sheet->SetCellValue('S3','AMOUNTPROVISIONS_CUR');
$sheet->SetCellValue('T3','LoanAcount');
$sheet->SetCellValue('U3','GracePrincipal');
$sheet->SetCellValue('V3','GracePay');
$sheet->SetCellValue('W3','PlaceOfDisbursement');
$sheet->SetCellValue('X3','Classification');
$sheet->SetCellValue('Y3','ParentContractcode');
$sheet->SetCellValue('Z3','ParentProvider');
$sheet->SetCellValue('AA3','ParentContractStatus');
$sheet->SetCellValue('AB3','ParentOperationDate');
$sheet->SetCellValue('AC3','ProlongationCount');
$sheet->SetCellValue('AD3','BranchLocation');
$sheet->SetCellValue('AE3','paymentMethod');
$sheet->SetCellValue('AF3','paymentPeriodId');
$sheet->SetCellValue('AG3','TotalAmount');
$sheet->SetCellValue('AH3','TotalAmountCur');
$sheet->SetCellValue('AI3','InstalmentAmount');
$sheet->SetCellValue('AJ3','INSTALMENTAMOUNT_CUR');
$sheet->SetCellValue('AK3','InstalmentCount');
$sheet->SetCellValue('AL3','CreditLimit');
$sheet->SetCellValue('AM3','CreditLimit_Cur');
$sheet->SetCellValue('AN3','FundingSource');
$sheet->SetCellValue('AO3','AvailableDate');
$sheet->SetCellValue('AP3','AccountingDate');
$sheet->SetCellValue('AQ3','OutstandingInstallmentCount');
$sheet->SetCellValue('AR3','OutstandingAmount');
$sheet->SetCellValue('AS3','OverdueInstallmentCount');
$sheet->SetCellValue('AT3','OverdueAmount');
$sheet->SetCellValue('AU3','Fine');
$sheet->SetCellValue('AV3','FINE_CUR');
$sheet->SetCellValue('AW3','Penalty');
$sheet->SetCellValue('AX3','Penalty_cur');
$sheet->SetCellValue('AY3','ProlongationStartDate');
$sheet->SetCellValue('AZ3','ProlongationEndDate');
$sheet->SetCellValue('BA3','AvailableLimit');
$sheet->SetCellValue('BB3','AvailableLimit_Cur');
$sheet->SetCellValue('BC3','SubjectRole');
$sheet->SetCellValue('BD3','Surname');
$sheet->SetCellValue('BE3','FirstName');
$sheet->SetCellValue('BF3','FathersName (при наличии)');
$sheet->SetCellValue('BG3','BirthSurname');
$sheet->SetCellValue('BH3','Gender');
$sheet->SetCellValue('BI3','Classification');
$sheet->SetCellValue('BJ3','Residency');
$sheet->SetCellValue('BK3','Education');
$sheet->SetCellValue('BL3','MaritalStatus');
$sheet->SetCellValue('BM3','DateOfBirth');
$sheet->SetCellValue('BN3','NegativeStatus');
$sheet->SetCellValue('BO3','Profession');
$sheet->SetCellValue('BP3','EconomyActivityGroup');
$sheet->SetCellValue('BQ3','EmploymentNature');
$sheet->SetCellValue('BR3','Citizenship');
$sheet->SetCellValue('BS3','Salary');
$sheet->SetCellValue('BT3','IIN');
$sheet->SetCellValue('BU3','MobilePhone');
$sheet->SetCellValue('BV3','DependantLt18');
$sheet->SetCellValue('BW3','DependantGt18');
$sheet->SetCellValue('BX3','Number');
$sheet->SetCellValue('BY3','RegistrationDate');
$sheet->SetCellValue('BZ3','IssueDate');
$sheet->SetCellValue('CA3','ExpirationDate');
$sheet->SetCellValue('CB3','IssuedBy');
$sheet->SetCellValue('CC3','DocTypeId');
$sheet->SetCellValue('CD3','katoId');
$sheet->SetCellValue('CE3','StreetName');

$sheet->setTitle('Batch');

$sheet->getDefaultStyle()->getFont()->setName('Times New Roman');
$sheet->getDefaultStyle()->getFont()->setSize(10);

foreach ($array as $list){
    $k = PHPExcel_Shared_Date::PHPToExcel($list['start']);
    $l = PHPExcel_Shared_Date::PHPToExcel($list['end']);
    $bm = PHPExcel_Shared_Date::PHPToExcel($list['birthDate']);
//    if (!empty($list['factPaymentDate'])){
        $n = PHPExcel_Shared_Date::PHPToExcel($list['factPaymentDate']);
//    }
    $ap = PHPExcel_Shared_Date::PHPToExcel($date1);

    $sheet->SetCellValue('A'.$rowCount, '1');
    $sheet->SetCellValue('B'.$rowCount, $list['parentContractNumber']);
    $sheet->SetCellValue('C'.$rowCount, $list['parentContractNumber']);
    $sheet->SetCellValue('D'.$rowCount, '2');
    $sheet->SetCellValue('E'.$rowCount, '09');
    $sheet->SetCellValue('F'.$rowCount, '10');
    $sheet->SetCellValue('G'.$rowCount, $list['contractPhase']);
    $sheet->SetCellValue('H'.$rowCount, $list['contractStatus']);
    $sheet->SetCellValue('K'.$rowCount, $k);
    $sheet->SetCellValue('L'.$rowCount, $l);
    $sheet->SetCellValue('N'.$rowCount, iconv("windows-1251", "UTF-8", $n));
    $sheet->SetCellValue('P'.$rowCount, '54');
    $sheet->SetCellValue('Q'.$rowCount, '23,068');
    $sheet->SetCellValue('X'.$rowCount, '1');
    $sheet->SetCellValue('AE'.$rowCount, '6');
    $sheet->SetCellValue('AF'.$rowCount, '9');
    $sheet->SetCellValue('AG'.$rowCount, $list['od']);
    $sheet->SetCellValue('AH'.$rowCount, 'KZT');
    $sheet->SetCellValue('AI'.$rowCount, $list['amountCredit']);
    $sheet->SetCellValue('AJ'.$rowCount, 'KZT');
    $sheet->SetCellValue('AK'.$rowCount, '1');
    $sheet->SetCellValue('AP'.$rowCount, $ap);
    $sheet->SetCellValue('AQ'.$rowCount, $list['upcomingPayments']);
    $sheet->SetCellValue('AR'.$rowCount, $list['amountOutstanding']);
    $sheet->SetCellValue('AS'.$rowCount, $list['daysOverdue']);
    $sheet->SetCellValue('AT'.$rowCount, $list['amountOverdue']);
    $sheet->SetCellValue('BC'.$rowCount, '1');
    $sheet->SetCellValue('BD'.$rowCount, $list['surname']);
    $sheet->SetCellValue('BE'.$rowCount, $list['name']);
    $sheet->SetCellValue('BF'.$rowCount, $list['fatherName']);
    $sheet->SetCellValue('BH'.$rowCount, $list['gender']);
    $sheet->SetCellValue('BI'.$rowCount, '1');
    $sheet->SetCellValue('BJ'.$rowCount, '1');
    $sheet->SetCellValue('BM'.$rowCount, $bm);
    $sheet->SetCellValue('BR'.$rowCount, '110');
    $sheet->SetCellValue('BT'.$rowCount, $list['iin']);
    $sheet->SetCellValue('BX'.$rowCount, $list['identityCardNumber']);
    $sheet->SetCellValue('BY'.$rowCount, $k);
    $sheet->SetCellValue('CC'.$rowCount, '7');
    $sheet->SetCellValue('CD'.$rowCount, '431000000');
    $sheet->SetCellValue('CE'.$rowCount, $list['street']);
    //$sheet->SetCellValue('Y'.$rowCount, '=I'.$rowCount.'*1');
    /*'iin' => $s['UF_CRM_1573668162'],
    'start' => $s['UF_CRM_1575338848'],
    'end' => $s['UF_CRM_1575338914'],
    'surname' => $s['UF_CRM_1577671976335'],
    'name' => $s['UF_CRM_1577672007511'],
    'fatherName' => $s['UF_CRM_1577672053366'],
    'gender' => $s['UF_CRM_1577672434509'],
    'street' => $s['UF_CRM_1578082787898'],
    'birthDate' => $s['UF_CRM_5E9FBD498C467'],
    'identityCardNumber' => $s['UF_CRM_5E9FBD4A162DE'],
    'factPaymentDate' => $s['UF_CRM_1610348437'],
    'parentContractNumber' => $s['UF_CRM_1622005475'],
    'contractNumber' => $s['UF_CRM_1589873611'],
    'amountOnHand' => $s['UF_CRM_1589873735'],
    'amountInsurance' => $s['UF_CRM_1589873821'],*/

    $sheet->getStyle('E'.$rowCount)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_TEXT); //text
    $sheet->getStyle('BT'.$rowCount)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_TEXT); //text
    $sheet->getStyle('BX'.$rowCount)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_TEXT); //text
    $sheet->getStyle('CD'.$rowCount)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_TEXT); //text
    $sheet->getStyle('K'.$rowCount)->getNumberFormat()->setFormatCode('dd/mm/yyyy'); //date
    $sheet->getStyle('L'.$rowCount)->getNumberFormat()->setFormatCode('dd/mm/yyyy'); //date
    $sheet->getStyle('N'.$rowCount)->getNumberFormat()->setFormatCode('dd/mm/yyyy'); //date
    $sheet->getStyle('AP'.$rowCount)->getNumberFormat()->setFormatCode('dd/mm/yyyy'); //date
    $sheet->getStyle('BM'.$rowCount)->getNumberFormat()->setFormatCode('dd/mm/yyyy'); //date
    $sheet->getStyle('BY'.$rowCount)->getNumberFormat()->setFormatCode('dd/mm/yyyy'); //date
    $sheet->getStyle('Q'.$rowCount)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER); //number
    $rowCount++;
}

echo $rowCount;
echo 'lol';

$objWriter = new PHPExcel_Writer_Excel5($objPHPExcel);
$uploadDir = '/upload/report';

if (!is_dir($_SERVER["DOCUMENT_ROOT"] . $uploadDir)) {
    mkdir($_SERVER["DOCUMENT_ROOT"] . $uploadDir, 0664);
}
$fileName = "Данные_для_загрузки_в_ПКБ_ГКБ_за_".$date.".xls";
$fileFullName = $_SERVER["DOCUMENT_ROOT"] . $uploadDir . '/' . $fileName;
$objWriter->save($fileFullName);


$fields = [
    "EMAIL_TO" => 'a.abzal@i-credit.kz',
    "TEXT" => 'TEST',
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