<?php
date_default_timezone_set('Asia/Seoul');
require_once './PHPExcel-1.8/Classes/PHPExcel.php';

//첫번째 변수 - 시작날짜, 두번째 변수- 종료날짜, 세번째 변수 - 지역 이름(GURO 고정)
$STARTDATE=$_GET['start_date'];
$ENDDATE=$_GET['end_date'];
$STAT=$_GET['stat'];

if($STAT == "GURO")
    {

 // 엑셀의 전체적인 양식 설정
$objPHPExcel = new PHPExcel();

$sheet = $objPHPExcel->getActiveSheet();
$sheetIndex = $objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(15);
$sheetIndex->getStyle('B2:BL100')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$sheetIndex->getStyle('A1:A100')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$sheetIndex->getStyle('A1:BL1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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

$PGW01_TRAFFIC="SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(UTIL_RX_5MIN_MBPS) AS RX,SUM(UTIL_TX_5MIN_MBPS) AS TX from ".$STAT."_PGW01_PortStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$PGW03_TRAFFIC="SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(UTIL_RX_5MIN_MBPS) AS RX,SUM(UTIL_TX_5MIN_MBPS) AS TX from ".$STAT."_PGW03_PortStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$PGW04_TRAFFIC="SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(UTIL_RX_5MIN_MBPS) AS RX,SUM(UTIL_TX_5MIN_MBPS) AS TX from ".$STAT."_PGW04_PortStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$PGW05_TRAFFIC="SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(UTIL_RX_5MIN_MBPS) AS RX,SUM(UTIL_TX_5MIN_MBPS) AS TX from ".$STAT."_PGW05_PortStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$PGW11_TRAFFIC="SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(UTIL_RX_5MIN_MBPS) AS RX,SUM(UTIL_TX_5MIN_MBPS) AS TX from ".$STAT."_PGW11_PortStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$PGW12_TRAFFIC="SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(UTIL_RX_5MIN_MBPS) AS RX,SUM(UTIL_TX_5MIN_MBPS) AS TX from ".$STAT."_PGW12_PortStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$PGW13_TRAFFIC="SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(UTIL_RX_5MIN_MBPS) AS RX,SUM(UTIL_TX_5MIN_MBPS) AS TX from ".$STAT."_PGW13_PortStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$PGW15_TRAFFIC="SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(UTIL_RX_5MIN_MBPS) AS RX,SUM(UTIL_TX_5MIN_MBPS) AS TX from ".$STAT."_PGW15_PortStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$PGW16_TRAFFIC="SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(UTIL_RX_5MIN_MBPS) AS RX,SUM(UTIL_TX_5MIN_MBPS) AS TX from ".$STAT."_PGW16_PortStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$PGW17_TRAFFIC="SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(UTIL_RX_5MIN_MBPS) AS RX,SUM(UTIL_TX_5MIN_MBPS) AS TX from ".$STAT."_PGW17_PortStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$PGW18_TRAFFIC="SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(UTIL_RX_5MIN_MBPS) AS RX,SUM(UTIL_TX_5MIN_MBPS) AS TX from ".$STAT."_PGW18_PortStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$PGW19_TRAFFIC="SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(UTIL_RX_5MIN_MBPS) AS RX,SUM(UTIL_TX_5MIN_MBPS) AS TX from ".$STAT."_PGW19_PortStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$PGW20_TRAFFIC="SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(UTIL_RX_5MIN_MBPS) AS RX,SUM(UTIL_TX_5MIN_MBPS) AS TX from ".$STAT."_PGW20_PortStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";

$SGW11_TRAFFIC="SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(UTIL_RX_5MIN_MBPS) AS RX,SUM(UTIL_TX_5MIN_MBPS) AS TX from ".$STAT."_SGW11_PortStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$SGW12_TRAFFIC="SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(UTIL_RX_5MIN_MBPS) AS RX,SUM(UTIL_TX_5MIN_MBPS) AS TX from ".$STAT."_SGW12_PortStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$SGW13_TRAFFIC="SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(UTIL_RX_5MIN_MBPS) AS RX,SUM(UTIL_TX_5MIN_MBPS) AS TX from ".$STAT."_SGW13_PortStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$SGW14_TRAFFIC="SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(UTIL_RX_5MIN_MBPS) AS RX,SUM(UTIL_TX_5MIN_MBPS) AS TX from ".$STAT."_SGW14_PortStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$SGW15_TRAFFIC="SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(UTIL_RX_5MIN_MBPS) AS RX,SUM(UTIL_TX_5MIN_MBPS) AS TX from ".$STAT."_SGW15_PortStat where (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";

$UPF02_01_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF02_PortUtilStat where SYSTEM = 'GURO_UPF02-01' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF02_02_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF02_PortUtilStat where SYSTEM = 'GURO_UPF02-02' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF02_03_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF02_PortUtilStat where SYSTEM = 'GURO_UPF02-03' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF02_04_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF02_PortUtilStat where SYSTEM = 'GURO_UPF02-04' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF02_05_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF02_PortUtilStat where SYSTEM = 'GURO_UPF02-05' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF02_06_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF02_PortUtilStat where SYSTEM = 'GURO_UPF02-06' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF02_07_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF02_PortUtilStat where SYSTEM = 'GURO_UPF02-07' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF02_08_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF02_PortUtilStat where SYSTEM = 'GURO_UPF02-08' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF02_09_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF02_PortUtilStat where SYSTEM = 'GURO_UPF02-09' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF02_10_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF02_PortUtilStat where SYSTEM = 'GURO_UPF02-10' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF02_11_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF02_PortUtilStat where SYSTEM = 'GURO_UPF02-11' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF02_12_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF02_PortUtilStat where SYSTEM = 'GURO_UPF02-12' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";

$UPF11_01_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF11_PortUtilStat where SYSTEM = 'GURO_UPF11-01' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF11_02_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF11_PortUtilStat where SYSTEM = 'GURO_UPF11-02' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF11_03_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF11_PortUtilStat where SYSTEM = 'GURO_UPF11-03' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF11_04_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF11_PortUtilStat where SYSTEM = 'GURO_UPF11-04' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF11_05_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF11_PortUtilStat where SYSTEM = 'GURO_UPF11-05' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF11_06_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF11_PortUtilStat where SYSTEM = 'GURO_UPF11-06' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF11_07_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF11_PortUtilStat where SYSTEM = 'GURO_UPF11-07' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF11_08_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF11_PortUtilStat where SYSTEM = 'GURO_UPF11-08' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";

$UPF12_01_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF12_PortUtilStat where SYSTEM = 'GURO_UPF12-01' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF12_02_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF12_PortUtilStat where SYSTEM = 'GURO_UPF12-02' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF12_03_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF12_PortUtilStat where SYSTEM = 'GURO_UPF12-03' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF12_04_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF12_PortUtilStat where SYSTEM = 'GURO_UPF12-04' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF12_05_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF12_PortUtilStat where SYSTEM = 'GURO_UPF12-05' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF12_06_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF12_PortUtilStat where SYSTEM = 'GURO_UPF12-06' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF12_07_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF12_PortUtilStat where SYSTEM = 'GURO_UPF12-07' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF12_08_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF12_PortUtilStat where SYSTEM = 'GURO_UPF12-08' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF12_09_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF12_PortUtilStat where SYSTEM = 'GURO_UPF12-09' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF12_10_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF12_PortUtilStat where SYSTEM = 'GURO_UPF12-10' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF12_11_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF12_PortUtilStat where SYSTEM = 'GURO_UPF12-11' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF12_12_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF12_PortUtilStat where SYSTEM = 'GURO_UPF12-12' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";

$UPF13_01_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF13_PortUtilStat where SYSTEM = 'GURO_UPF13-01' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF13_02_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF13_PortUtilStat where SYSTEM = 'GURO_UPF13-02' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF13_03_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF13_PortUtilStat where SYSTEM = 'GURO_UPF13-03' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF13_04_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF13_PortUtilStat where SYSTEM = 'GURO_UPF13-04' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF13_05_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF13_PortUtilStat where SYSTEM = 'GURO_UPF13-05' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF13_06_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF13_PortUtilStat where SYSTEM = 'GURO_UPF13-06' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF13_07_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF13_PortUtilStat where SYSTEM = 'GURO_UPF13-07' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF13_08_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF13_PortUtilStat where SYSTEM = 'GURO_UPF13-08' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF13_09_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF13_PortUtilStat where SYSTEM = 'GURO_UPF13-09' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF13_10_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF13_PortUtilStat where SYSTEM = 'GURO_UPF13-10' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF13_11_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF13_PortUtilStat where SYSTEM = 'GURO_UPF13-11' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";
$UPF13_12_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(RX_AVG_5MIN_MBPS) AS RX,SUM(TX_AVG_5MIN_MBPS) AS TX from ".$STAT."_UPF13_PortUtilStat where SYSTEM = 'GURO_UPF13-12' AND (DATETIME >= '".$new_date." 00:00:00') AND (DATETIME <= '".$new_date." 23:55:55') GROUP BY DATETIME) subquery1;";

$SMF02_TRAFFIC="SELECT SYSTEM,MAX(subquery1.RX) FROM (SELECT SYSTEM,DATETIME,SUM(UTIL_RX_5MIN_MBPS) AS RX,SUM(UTIL_TX_5MIN_MBPS) AS TX from ".$STAT."_SMF02_PortStat where DATETIME BETWEEN '".$STARTDATE."' AND '".$ENDDATE."' GROUP BY DATETIME) subquery1;";

// 행을 한 줄씩 추가하면서, SQL 결과를 저장

$sheetIndex->setCellValue('A'.$i,$new_date);

$result = mysqli_query($conn, $PGW01_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/14.4*100,2);

$sheetIndex->setCellValue('B'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $PGW03_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/14.4*100,2);

$sheetIndex->setCellValue('C'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $PGW04_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/14.4*100,2);

$sheetIndex->setCellValue('D'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $PGW05_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/9*100,2);

$sheetIndex->setCellValue('E'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $PGW11_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/43*100,2);

$sheetIndex->setCellValue('F'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $PGW12_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/43*100,2);

$sheetIndex->setCellValue('G'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $PGW13_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/43*100,2);

$sheetIndex->setCellValue('H'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $PGW15_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/43*100,2);

$sheetIndex->setCellValue('I'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $PGW16_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/43*100,2);

$sheetIndex->setCellValue('J'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $PGW17_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/43*100,2);

$sheetIndex->setCellValue('K'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $PGW18_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/43*100,2);

$sheetIndex->setCellValue('L'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $PGW19_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/43*100,2);

$sheetIndex->setCellValue('M'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $PGW20_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/43*100,2);

$sheetIndex->setCellValue('N'.$i,$TEMP.'G ('.$LOAD.' %)');


$result = mysqli_query($conn, $SGW11_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/70*100,2);

$sheetIndex->setCellValue('O'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $SGW12_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/70*100,2);

$sheetIndex->setCellValue('P'.$i,$TEMP.'G ('.$LOAD.' %)');


$result = mysqli_query($conn, $SGW13_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/70*100,2);

$sheetIndex->setCellValue('Q'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $SGW14_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/70*100,2);

$sheetIndex->setCellValue('R'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $SGW15_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/70*100,2);

$sheetIndex->setCellValue('S'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF02_01_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('T'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF02_02_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('U'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF02_03_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('V'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF02_04_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('W'.$i,$TEMP.'G ('.$LOAD.' %)');


$result = mysqli_query($conn, $UPF02_05_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('X'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF02_06_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('Y'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF02_07_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('Z'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF02_08_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('AA'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF02_09_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('AB'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF02_10_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('AC'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF02_11_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('AD'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF02_12_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('AE'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF11_01_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('AF'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF11_02_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('AG'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF11_03_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('AH'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF11_04_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('AI'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF11_05_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('AJ'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF11_06_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('AK'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF11_07_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('AL'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF11_08_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('AM'.$i,$TEMP.'G ('.$LOAD.' %)');


$result = mysqli_query($conn, $UPF12_01_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('AN'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF12_02_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('AO'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF12_03_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('AP'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF12_04_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('AQ'.$i,$TEMP.'G ('.$LOAD.' %)');


$result = mysqli_query($conn, $UPF12_05_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('AR'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF12_06_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('AS'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF12_07_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('AT'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF12_08_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('AU'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF12_09_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('AV'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF12_10_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('AW'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF12_11_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('AX'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF12_12_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('AY'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF13_01_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('AZ'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF13_02_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('BA'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF13_03_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('BB'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF13_04_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('BC'.$i,$TEMP.'G ('.$LOAD.' %)');


$result = mysqli_query($conn, $UPF13_05_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('BD'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF13_06_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('BE'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF13_07_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('BF'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF13_08_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('BG'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF13_09_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('BH'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF13_10_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('BI'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF13_11_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('BJ'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $UPF13_12_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('BK'.$i,$TEMP.'G ('.$LOAD.' %)');

$result = mysqli_query($conn, $SMF02_TRAFFIC);
$row = mysqli_fetch_array($result);
$TEMP = $row['MAX(subquery1.RX)']/1000;
$LOAD = round($TEMP/30*100,2);

$sheetIndex->setCellValue('BL'.$i,$TEMP.'G ('.$LOAD.' %)');

// 종료날짜에 다르면 종료

if($new_date == $ENDDATE) break;
$i=$i+1;


}

// 엑셀 셀 옵션 설정

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('N')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('O')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('P')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('R')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('S')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('T')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('U')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('V')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('W')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('X')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('Y')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('Z')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AA')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AB')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AC')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AD')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AE')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AF')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AG')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AH')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AI')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AJ')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AK')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AL')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AM')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AN')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AO')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AP')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AQ')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AR')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AS')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AT')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AU')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AV')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AW')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AX')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AY')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AZ')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('BA')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('BB')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('BC')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('BD')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('BE')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('BF')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('BG')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('BH')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('BI')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('BJ')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('BK')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('BL')->setWidth(20);

mysqli_close($conn);

// 엑셀파일로 저장

header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename='.$STARTDATE.'___'.$ENDDATE.'_트래픽 추이.xls');
header('Cache-Control: max-age=0');
 
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');
 
exit;

   }
 
?>
