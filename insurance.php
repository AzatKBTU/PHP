<?

	$_SERVER["DOCUMENT_ROOT"] = "/home/bitrix/www";
	$DOCUMENT_ROOT = $_SERVER["DOCUMENT_ROOT"];
	
	define("NO_KEEP_STATISTIC", true);
	define("NOT_CHECK_PERMISSIONS", true);
	require($_SERVER["DOCUMENT_ROOT"]."/bitrix/modules/main/include/prolog_before.php");
	set_time_limit(0); 
	
	global $USER;
	global $DB;

	CModule::IncludeModule("iblock");
	CModule::IncludeModule("crm");
	CModule::IncludeModule("tasks");
	CModule::IncludeModule("sale");
	CModule::IncludeModule("socialnetwork");

	$url = "https://app.ffin.life/ws/ffinlife.wsdl";
	$username = "i-credit";
	$password = "HH67bnsdq12n";
  	$xmlBody = '
		<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ws="http://ffinlife/ws">
   <soapenv:Header/>
   <soapenv:Body>
      <ws:authorizationWSRequest>
         <ws:login>i-credit</ws:login>
         <ws:password>HH67bnsdq12n</ws:password>
      </ws:authorizationWSRequest>
   </soapenv:Body>
</soapenv:Envelope>
	';

	$ch = curl_init($url);
	$headers = array(
		'Content-Type:application/xop+xml',
		'Authorization: Basic a2FzcGk6a2FzcGk=' // <---
	);
	curl_setopt($ch, CURLOPT_HTTPHEADER, $headers);
	curl_setopt($ch, CURLOPT_POST, 1);
	curl_setopt($ch, CURLOPT_SSL_VERIFYHOST, 0);
	curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, 0);
	curl_setopt($ch, CURLOPT_POSTFIELDS, $xmlBody);
	curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
	$output = curl_exec($ch);
	curl_close($ch);
	$response = preg_replace("/(<\/?)(\w+):([^>]*>)/", "$1$2$3", $output);
	$output = explode('Content-Type: application/xop+xml; charset=utf-8; type="text/xml"',$output);
	$res = explode("------=",$output[1]);
	$s = preg_replace("/(<\/?)(\w+):([^>]*>)/", "$1$2$3", $res[0]);
	$session = explode('<ns2:sessionId>',$res[0]);
	$sessionID = explode('</ns2:sessionId>',$session[1]);
	$sessionID = $sessionID[0];
	$dd = date('Y-m-d');

	$all = "104729,104730,104731,104732,104733,104734,104735,104736,104737,104738,104927,104928,104929,104930,104931,104932,104933,104934,104935,104936,104937,104938,104939,104940,104941,104942,104943,104944,104945,104946,104947,104948,104949,104950,104951,104952,104953,104954,104955,104956,104957,104958,104959,104960,104961,104962,104963,104964,104965,104966,104967,104968,104969,104970,104971,104972,104973,104974,104975,104976,104977,104978,104979,104980,104981,104982,104983,104984,104985,104986,104987,104988,104989,104990,104991,104992,104993,104994,104995,104996,104997,104998,104999,105000,105001,105002,105003,105004,105005,105006,105007,105008,105009,105010,105011,105012,105013,105014,105015,105016,105017,105018,105019,105020,105021,105022,105023,105024,105025,105026,105027,105028,105029,105030,105031,105032,105033,105034,105035,105036,105037,105038,105039,105040,105041,105042,105043,105044,105045,105046,105047,105048,105049,105050,105051,105052,105053,105054,105055,105056,105057,105058,105059,105060,105179,105180,105181,105182,105183,105184,105185,105186,105187,105188,105189,105190,105191,105192,105193,105194,105195,105196";
	$sql = "SELECT * FROM b_crm_deal JOIN b_uts_crm_deal ON b_crm_deal.ID = b_uts_crm_deal.VALUE_ID WHERE b_crm_deal.CATEGORY_ID = 2 AND b_crm_deal.ID IN ($all)";


	$result = $DB->Query($sql);

	while ($freedom = $result->Fetch()){

	$surname = $freedom['UF_CRM_1577671976335'];
		$dealID = $freedom['ID'];
		$mainAmount = intval($freedom['OPPORTUNITY']);
		$name = $freedom['UF_CRM_1577672007511'];
		$fatherName = $freedom['UF_CRM_1577672053366'];
		$fullName = $surname.' '.$name.' '.$fatherName;
		$fullName = trim($fullName);
		$iin = $freedom['UF_CRM_1573668162'];
		$birth = str_split($iin);

		$date = $birth[0].$birth[1].$birth[2].$birth[3].$birth[4].$birth[5];

		$century = $birth[6];

		if ($century == 3 || $century == 4){
			$birthDay = $birth[4].$birth[5].".".$birth[2].$birth[3]."."."19".$birth[0].$birth[1];
		}

		if ($century == 5 || $century == 6){
			$birthDay = $birth[4].$birth[5].".".$birth[2].$birth[3]."."."20".$birth[0].$birth[1];
		}
		$contractNumber = 'AR'.$freedom['UF_CRM_1589873611'];
		$phone = $freedom['UF_CRM_1577672738323'];
		//$start = $freedom['UF_CRM_1575338848'];
		//$end = $freedom['UF_CRM_1575338914'];
		//$start = date('d.m.Y',strtotime($start));
		$start = date('d.m.Y');

		$end = date('d.m.Y',strtotime($end));
		$email = $freedom['UF_CRM_5E9FBD48C0802'];
		$period = $freedom['UF_CRM_1574846030'];
		//$period = $period-1;
		$end = date('d.m.Y',strtotime($start." + $period day"));
		$insurance = $freedom['UF_CRM_1589873821'];
		$amountHand = $freedom['UF_CRM_1589873735'];
		$insuranceAmount = $freedom['UF_CRM_1589873735'];
		$issueDate = $freedom['UF_CRM_5E9FBD4A2FCF4'];
		$docNumber = $freedom['UF_CRM_5E9FBD4A162DE'];
		$issueGiven = $freedom['UF_CRM_5E9FBD4A6007B'];
	$fields = "
			<soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/' xmlns:ws='http://ffinlife/ws'>
	   <soapenv:Header/>
	   <soapenv:Body>
						<ws:creditLifeOnlineRequest>
							<ws:sessionId>$sessionID</ws:sessionId>
							<ws:fullName>$fullName</ws:fullName>
							<ws:iin>$iin</ws:iin>
							<ws:birthDate>$birthDay</ws:birthDate>
							<ws:docNumber>$docNumber</ws:docNumber>
							<ws:docDate>$issueDate</ws:docDate>
							<ws:docIssue>$issueGiven</ws:docIssue>
							<ws:eMail>$email</ws:eMail>
							<ws:loanAgreementNumber>$contractNumber</ws:loanAgreementNumber>
							<ws:phoneNumber>$phone</ws:phoneNumber>
							<ws:resident>1</ws:resident>
							<ws:startDate>$start</ws:startDate>
							<ws:endDate>$end</ws:endDate>
							<ws:term>$period</ws:term>
							<ws:creditSum>$insuranceAmount</ws:creditSum>
							<ws:percent>$period</ws:percent>
							<ws:login>i-credit</ws:login>
							<ws:insuranceType>life</ws:insuranceType>
						</ws:creditLifeOnlineRequest>
	</soapenv:Body>
	</soapenv:Envelope> 
	";	
							   print_r($fields);
	$url = "https://app.ffin.life/ws/ffinlife.wsdl";
	$basic = base64_encode('test_life'.'test_life');

	$headers = [
			'Content-Type:application/xop+xml',
			'Authorization: Basic a2FzcGk6a2FzcGk=',
		];
	$ch = curl_init($url);
	curl_setopt($ch, CURLOPT_HTTPHEADER, $headers);
	curl_setopt($ch, CURLOPT_POST, 1);
	curl_setopt($ch, CURLOPT_SSL_VERIFYHOST, 0);
	curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, 0);
	curl_setopt($ch, CURLOPT_POSTFIELDS, $fields);
	curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
	$output = curl_exec($ch);
	var_dump($output);


	$success = explode('<ns2:success>',$output);
	$success = explode('</ns2:success>',$success[1]);

	$resp = explode('<ns2:response>',$output);
	$resp = explode('</ns2:response>',$resp[1]);
	$resp = $resp[0];
	$success = $success[0];
	$time = date('d.m.Y');
	$sql = "INSERT INTO freedom_manual(fullName,iin,phone,dealID,contractNumber,dateStart,dateEnd,period,mainAmount,insurance,insuranceAmount,response,success,created_at) 
			VALUES('$fullName','$iin','$phone',$dealID,'$contractNumber','$start','$end',$period,$mainAmount,$insurance,$insuranceAmount,'$resp','$success','$time')";
	$DB->Query($sql);

	}
