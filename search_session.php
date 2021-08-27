<?php
date_default_timezone_set('Asia/Seoul');
require_once './PHPExcel-1.8/Classes/PHPExcel.php';

//첫번째 변수 - 시작날짜, 두번째 변수- 종료날짜, 세번째 변수 - 지역 이름(GURO 고정)
$STARTDATE=$_GET['start_date'];
$ENDDATE=$_GET['end_date'];
$STAT=$_GET['stat'];

if($STAT == "GURO")
    {

$objPHPExcel = new PHPExcel();

 // 엑셀의 전체적인 양식 설정
 
$sheet = $objPHPExcel->getActiveSheet();
$sheetIndex = $objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(30);
$sheetIndex->getStyle('B2:BL100')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$sheetIndex->getStyle('A1:A100')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$sheetIndex->getStyle('A1:BK1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$sheetIndex->setCellValue('A1','날짜');
$sheetIndex->setCellValue('B1','PGW#1');
$sheetIndex->setCellValue('C1','PGW#3');
$sheetIndex->setCellValue('D1','PGW#4');
$sheetIndex->setCellValue('E1','PGW#5');
$sheetIndex->setCellValue('F1','PGW#11');
$sheetIndex->setCellValue('G1','PGW#12');
$sheetIndex->setCellValue('H1','PGW#13');
$sheetIndex->setCellValue('I1','PGW#15');
$sheetIndex->setCellValue('J1','PGW#16');
$sheetIndex->setCellValue('K1','PGW#17');
$sheetIndex->setCellValue('L1','PGW#18');
$sheetIndex->setCellValue('M1','PGW#19');
$sheetIndex->setCellValue('N1','PGW#20');
$sheetIndex->setCellValue('O1','SGW#11');
$sheetIndex->setCellValue('P1','SGW#12');
$sheetIndex->setCellValue('Q1','SGW#13');
$sheetIndex->setCellValue('R1','SGW#14');
$sheetIndex->setCellValue('S1','SGW#15');
$sheetIndex->setCellValue('T1','UPF#2-1');
$sheetIndex->setCellValue('U1','UPF#2-2');
$sheetIndex->setCellValue('V1','UPF#2-3');
$sheetIndex->setCellValue('W1','UPF#2-4');
$sheetIndex->setCellValue('X1','UPF#2-5');
$sheetIndex->setCellValue('Y1','UPF#2-6');
$sheetIndex->setCellValue('Z1','UPF#2-7');
$sheetIndex->setCellValue('AA1','UPF#2-8');
$sheetIndex->setCellValue('AB1','UPF#2-9');
$sheetIndex->setCellValue('AC1','UPF#2-10');
$sheetIndex->setCellValue('AD1','UPF#2-11');
$sheetIndex->setCellValue('AE1','UPF#2-12');
$sheetIndex->setCellValue('AF1','UPF#11-1');
$sheetIndex->setCellValue('AG1','UPF#11-2');
$sheetIndex->setCellValue('AH1','UPF#11-3');
$sheetIndex->setCellValue('AI1','UPF#11-4');
$sheetIndex->setCellValue('AJ1','UPF#11-5');
$sheetIndex->setCellValue('AK1','UPF#11-6');
$sheetIndex->setCellValue('AL1','UPF#11-7');
$sheetIndex->setCellValue('AM1','UPF#11-8');
$sheetIndex->setCellValue('AN1','UPF#12-1');
$sheetIndex->setCellValue('AO1','UPF#12-2');
$sheetIndex->setCellValue('AP1','UPF#12-3');
$sheetIndex->setCellValue('AQ1','UPF#12-4');
$sheetIndex->setCellValue('AR1','UPF#12-5');
$sheetIndex->setCellValue('AS1','UPF#12-6');
$sheetIndex->setCellValue('AT1','UPF#12-7');
$sheetIndex->setCellValue('AU1','UPF#12-8');
$sheetIndex->setCellValue('AV1','UPF#12-9');
$sheetIndex->setCellValue('AW1','UPF#12-10');
$sheetIndex->setCellValue('AX1','UPF#12-11');
$sheetIndex->setCellValue('AY1','UPF#12-12');
$sheetIndex->setCellValue('AZ1','UPF#13-1');
$sheetIndex->setCellValue('BA1','UPF#13-2');
$sheetIndex->setCellValue('BB1','UPF#13-3');
$sheetIndex->setCellValue('BC1','UPF#13-4');
$sheetIndex->setCellValue('BD1','UPF#13-5');
$sheetIndex->setCellValue('BE1','UPF#13-6');
$sheetIndex->setCellValue('BF1','UPF#13-7');
$sheetIndex->setCellValue('BG1','UPF#13-8');
$sheetIndex->setCellValue('BH1','UPF#13-9');
$sheetIndex->setCellValue('BI1','UPF#13-10');
$sheetIndex->setCellValue('BJ1','UPF#13-11');
$sheetIndex->setCellValue('BK1','UPF#13-12');
$sheetIndex->setCellValue('BL1','SMF02');

$new_date = date("Y-m-d", strtotime("-1 day", strtotime($STARTDATE)));

$i=2;

$conn = mysqli_connect(
  '172.21.223.167',
  'guro',
  'Drcore12345!',
  'MACA_RAW');



// 시작 날짜부터 종료날짜 까지 반복

while(true) {
	
	// 하루 씩 추가하여 계산
$new_date = date("Y-m-d", strtotime("+1 day", strtotime($new_date)));

// MYSQL 에 접속해서 하루 씩 가입자 수 DATA 파싱

$PGW01_SESSION="SELECT MAX(SESSSTAT_BEARACT_DEF) from ".$STAT."_PGW01_PGWAttachStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$PGW03_SESSION="SELECT MAX(SESSSTAT_BEARACT_DEF) from ".$STAT."_PGW03_PGWAttachStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$PGW04_SESSION="SELECT MAX(SESSSTAT_BEARACT_DEF) from ".$STAT."_PGW04_PGWAttachStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$PGW05_SESSION="SELECT MAX(SESSSTAT_BEARACT_DEF) from ".$STAT."_PGW05_PGWAttachStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$PGW11_SESSION="SELECT MAX(SESSSTAT_BEARACT_DEF) from ".$STAT."_PGW11_PGWAttachStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$PGW12_SESSION="SELECT MAX(SESSSTAT_BEARACT_DEF) from ".$STAT."_PGW12_PGWAttachStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$PGW13_SESSION="SELECT MAX(SESSSTAT_BEARACT_DEF+SESSSTAT_BEARACT_DED) from ".$STAT."_PGW13_PGWAttachStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$PGW15_SESSION="SELECT MAX(SESSSTAT_BEARACT_DEF) from ".$STAT."_PGW15_PGWAttachStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$PGW16_SESSION="SELECT MAX(SESSSTAT_BEARACT_DEF) from ".$STAT."_PGW16_PGWAttachStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$PGW17_SESSION="SELECT MAX(SESSSTAT_BEARACT_DEF) from ".$STAT."_PGW17_PGWAttachStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$PGW18_SESSION="SELECT MAX(SESSSTAT_BEARACT_DEF) from ".$STAT."_PGW18_PGWAttachStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$PGW19_SESSION="SELECT MAX(SESSSTAT_BEARACT_DEF) from ".$STAT."_PGW19_PGWAttachStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$PGW20_SESSION="SELECT MAX(SESSSTAT_BEARACT_DEF) from ".$STAT."_PGW20_PGWAttachStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";

$SGW11_SESSION="SELECT MAX(SESSSTAT_TOTCUR_BEARERS) from ".$STAT."_SGW11_SGWAttachStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$SGW12_SESSION="SELECT MAX(SESSSTAT_TOTCUR_BEARERS) from ".$STAT."_SGW12_SGWAttachStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$SGW13_SESSION="SELECT MAX(SESSSTAT_TOTCUR_BEARERS) from ".$STAT."_SGW13_SGWAttachStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$SGW14_SESSION="SELECT MAX(SESSSTAT_TOTCUR_BEARERS) from ".$STAT."_SGW14_SGWAttachStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$SGW15_SESSION="SELECT MAX(SESSSTAT_TOTCUR_BEARERS) from ".$STAT."_SGW15_SGWAttachStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";

$UPF02_01_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF02_SxPeerStat where SYSTEM='GURO_UPF02-01' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF02_02_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF02_SxPeerStat where SYSTEM='GURO_UPF02-02' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF02_03_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF02_SxPeerStat where SYSTEM='GURO_UPF02-03' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF02_04_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF02_SxPeerStat where SYSTEM='GURO_UPF02-04' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF02_05_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF02_SxPeerStat where SYSTEM='GURO_UPF02-05' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF02_06_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF02_SxPeerStat where SYSTEM='GURO_UPF02-06' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF02_07_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF02_SxPeerStat where SYSTEM='GURO_UPF02-07' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF02_08_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF02_SxPeerStat where SYSTEM='GURO_UPF02-08' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF02_09_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF02_SxPeerStat where SYSTEM='GURO_UPF02-09' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF02_10_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF02_SxPeerStat where SYSTEM='GURO_UPF02-10' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF02_11_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF02_SxPeerStat where SYSTEM='GURO_UPF02-11' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF02_12_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF02_SxPeerStat where SYSTEM='GURO_UPF02-12' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";

$UPF11_01_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF11_SxPeerStat where SYSTEM='GURO_UPF11-01' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF11_02_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF11_SxPeerStat where SYSTEM='GURO_UPF11-02' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF11_03_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF11_SxPeerStat where SYSTEM='GURO_UPF11-03' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF11_04_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF11_SxPeerStat where SYSTEM='GURO_UPF11-04' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF11_05_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF11_SxPeerStat where SYSTEM='GURO_UPF11-05' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF11_06_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF11_SxPeerStat where SYSTEM='GURO_UPF11-06' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF11_07_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF11_SxPeerStat where SYSTEM='GURO_UPF11-07' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF11_08_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF11_SxPeerStat where SYSTEM='GURO_UPF11-08' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";

$UPF12_01_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF12_SxPeerStat where SYSTEM='GURO_UPF12-01' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF12_02_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF12_SxPeerStat where SYSTEM='GURO_UPF12-02' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF12_03_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF12_SxPeerStat where SYSTEM='GURO_UPF12-03' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF12_04_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF12_SxPeerStat where SYSTEM='GURO_UPF12-04' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF12_05_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF12_SxPeerStat where SYSTEM='GURO_UPF12-05' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF12_06_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF12_SxPeerStat where SYSTEM='GURO_UPF12-06' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF12_07_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF12_SxPeerStat where SYSTEM='GURO_UPF12-07' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF12_08_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF12_SxPeerStat where SYSTEM='GURO_UPF12-08' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF12_09_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF12_SxPeerStat where SYSTEM='GURO_UPF12-09' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF12_10_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF12_SxPeerStat where SYSTEM='GURO_UPF12-10' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF12_11_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF12_SxPeerStat where SYSTEM='GURO_UPF12-11' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF12_12_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF12_SxPeerStat where SYSTEM='GURO_UPF12-12' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";

$UPF13_01_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF13_SxPeerStat where SYSTEM='GURO_UPF13-01' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF13_02_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF13_SxPeerStat where SYSTEM='GURO_UPF13-02' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF13_03_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF13_SxPeerStat where SYSTEM='GURO_UPF13-03' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF13_04_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF13_SxPeerStat where SYSTEM='GURO_UPF13-04' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF13_05_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF13_SxPeerStat where SYSTEM='GURO_UPF13-05' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF13_06_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF13_SxPeerStat where SYSTEM='GURO_UPF13-06' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF13_07_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF13_SxPeerStat where SYSTEM='GURO_UPF13-07' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF13_08_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF13_SxPeerStat where SYSTEM='GURO_UPF13-08' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF13_09_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF13_SxPeerStat where SYSTEM='GURO_UPF13-09' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF13_10_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF13_SxPeerStat where SYSTEM='GURO_UPF13-10' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF13_11_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF13_SxPeerStat where SYSTEM='GURO_UPF13-11' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";
$UPF13_12_SESSION = "SELECT MAX(CURRENT_SESSION) from ".$STAT."_UPF13_SxPeerStat where SYSTEM='GURO_UPF13-12' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55');";


$SMF02_SESSION="SELECT MAX(SESSSTAT_TOTCUR_BEARERS) from ".$STAT."_SMF02_SGWAttachStat where DATETIME BETWEEN '".$STARTDATE."' AND '".$ENDDATE."';";

// 행을 한 줄씩 추가하면서, SQL 결과를 저장

$sheetIndex->setCellValue('A'.$i,$new_date);

$result = mysqli_query($conn, $PGW01_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(SESSSTAT_BEARACT_DEF)']/1000000*100,2);

$sheetIndex->setCellValue('B'.$i,$LOAD.' %');

$result = mysqli_query($conn, $PGW03_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(SESSSTAT_BEARACT_DEF)']/1000000*100,2);

$sheetIndex->setCellValue('C'.$i,$LOAD.' %');

$result = mysqli_query($conn, $PGW04_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(SESSSTAT_BEARACT_DEF)']/1000000*100,2);

$sheetIndex->setCellValue('D'.$i,$LOAD.' %');

$result = mysqli_query($conn, $PGW05_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(SESSSTAT_BEARACT_DEF)']/700000*100,2);

$sheetIndex->setCellValue('E'.$i,$LOAD.' %');

$result = mysqli_query($conn, $PGW11_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(SESSSTAT_BEARACT_DEF)']/1800000*100,2);

$sheetIndex->setCellValue('F'.$i,$LOAD.' %');

$result = mysqli_query($conn, $PGW12_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(SESSSTAT_BEARACT_DEF)']/1800000*100,2);

$sheetIndex->setCellValue('G'.$i,$LOAD.' %');

$result = mysqli_query($conn, $PGW13_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(SESSSTAT_BEARACT_DEF+SESSSTAT_BEARACT_DED)']/4000000*100,2);

$sheetIndex->setCellValue('H'.$i,$LOAD.' %');

$result = mysqli_query($conn, $PGW15_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(SESSSTAT_BEARACT_DEF)']/1800000*100,2);

$sheetIndex->setCellValue('I'.$i,$LOAD.' %');

$result = mysqli_query($conn, $PGW16_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(SESSSTAT_BEARACT_DEF)']/1800000*100,2);

$sheetIndex->setCellValue('J'.$i,$LOAD.' %');

$result = mysqli_query($conn, $PGW17_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(SESSSTAT_BEARACT_DEF)']/1800000*100,2);

$sheetIndex->setCellValue('K'.$i,$LOAD.' %');

$result = mysqli_query($conn, $PGW18_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(SESSSTAT_BEARACT_DEF)']/1800000*100,2);

$sheetIndex->setCellValue('L'.$i,$LOAD.' %');

$result = mysqli_query($conn, $PGW19_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(SESSSTAT_BEARACT_DEF)']/1800000*100,2);

$sheetIndex->setCellValue('M'.$i,$LOAD.' %');

$result = mysqli_query($conn, $PGW20_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(SESSSTAT_BEARACT_DEF)']/1800000*100,2);

$sheetIndex->setCellValue('N'.$i,$LOAD.' %');

$result = mysqli_query($conn, $SGW11_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(SESSSTAT_TOTCUR_BEARERS)']/3750000*100,2);

$sheetIndex->setCellValue('O'.$i,$LOAD.' %');

$result = mysqli_query($conn, $SGW12_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(SESSSTAT_TOTCUR_BEARERS)']/3750000*100,2);

$sheetIndex->setCellValue('P'.$i,$LOAD.' %');


$result = mysqli_query($conn, $SGW13_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(SESSSTAT_TOTCUR_BEARERS)']/3750000*100,2);

$sheetIndex->setCellValue('Q'.$i,$LOAD.' %');

$result = mysqli_query($conn, $SGW14_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(SESSSTAT_TOTCUR_BEARERS)']/3750000*100,2);

$sheetIndex->setCellValue('R'.$i,$LOAD.' %');

$result = mysqli_query($conn, $SGW15_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(SESSSTAT_TOTCUR_BEARERS)']/3750000*100,2);

$sheetIndex->setCellValue('S'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF02_01_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('T'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF02_02_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('U'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF02_03_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('V'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF02_04_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('W'.$i,$LOAD.' %');


$result = mysqli_query($conn, $UPF02_05_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('X'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF02_06_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('Y'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF02_07_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('Z'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF02_08_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('AA'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF02_09_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('AB'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF02_10_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('AC'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF02_11_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('AD'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF02_12_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('AE'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF11_01_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('AF'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF11_02_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('AG'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF11_03_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('AH'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF11_04_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);;

$sheetIndex->setCellValue('AI'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF11_05_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('AJ'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF11_06_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('AK'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF11_07_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('AL'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF11_08_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('AM'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF12_01_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('AN'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF12_02_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('AO'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF12_03_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('AP'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF12_04_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('AQ'.$i,$LOAD.' %');


$result = mysqli_query($conn, $UPF12_05_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('AR'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF12_06_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('AS'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF12_07_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('AT'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF12_08_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('AU'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF12_09_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('AV'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF12_10_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('AW'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF12_11_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('AX'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF12_12_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('AY'.$i,$LOAD.' %');


$result = mysqli_query($conn, $UPF13_01_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('AZ'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF13_02_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('BA'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF13_03_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('BB'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF13_04_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('BC'.$i,$LOAD.' %');


$result = mysqli_query($conn, $UPF13_05_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('BD'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF13_06_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('BE'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF13_07_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('BF'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF13_08_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('BG'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF13_09_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('BH'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF13_10_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('BI'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF13_11_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('BJ'.$i,$LOAD.' %');

$result = mysqli_query($conn, $UPF13_12_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(CURRENT_SESSION)']/1200000*100,2);

$sheetIndex->setCellValue('BK'.$i,$LOAD.' %');



$result = mysqli_query($conn, $SMF02_SESSION);
$row = mysqli_fetch_array($result);

$LOAD = round($row['MAX(SESSSTAT_TOTCUR_BEARERS)']/2000000*100,2);

$sheetIndex->setCellValue('BL'.$i,$LOAD.' %');

// 종료날짜에 다르면 종료

if($new_date == $ENDDATE) break;
$i=$i+1;


}

mysqli_close($conn);

// 엑셀파일로 저장

header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename='.$STARTDATE.'___'.$ENDDATE.'_세션 추이.xls');
header('Cache-Control: max-age=0');
 
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');
 
exit;

   }
 
?>
