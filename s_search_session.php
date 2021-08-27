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
$sheetIndex->getStyle('B2:AO100')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$sheetIndex->getStyle('A1:A100')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$sheetIndex->getStyle('A1:AO1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$sheetIndex->setCellValue('A1','날짜');
$sheetIndex->setCellValue('B1','MME#1');
$sheetIndex->setCellValue('C1','MME#2');
$sheetIndex->setCellValue('D1','vMME#1');
$sheetIndex->setCellValue('E1','AMF#1');
$sheetIndex->setCellValue('F1','AMF#2');
$sheetIndex->setCellValue('G1','vSPGW#1');
$sheetIndex->setCellValue('H1','vSPGW#2');
$sheetIndex->setCellValue('I1','vSPGW#3');
$sheetIndex->setCellValue('J1','PGW#14');
$sheetIndex->setCellValue('K1','PGW#50');
$sheetIndex->setCellValue('L1','PGW#51');
$sheetIndex->setCellValue('M1','SMF#1');
$sheetIndex->setCellValue('N1','SMF#3');
$sheetIndex->setCellValue('O1','UPF#1');
$sheetIndex->setCellValue('P1','UPF#3');
$sheetIndex->setCellValue('Q1','UPF#4');

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

								$sheetIndex->setCellValue('A'.$i,$new_date);
								
								// MYSQL 에 접속해서 하루 씩 가입자 수 DATA 파싱
								
								$MME1_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_A-MME1_MS|EMSC` where DATETIME BETWEEN '".$new_date." 00:00:00' AND '".$new_date." 23:55:55' GROUP BY DATETIME) subquery1;";
								$MME2_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_A-MME2_MS|EMSC` where DATETIME BETWEEN '".$new_date." 00:00:00' AND '".$new_date." 23:55:55' GROUP BY DATETIME) subquery1;";
								$vMME1_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_vMME1_MS|EMSC` where DATETIME BETWEEN '".$new_date." 00:00:00' AND '".$new_date." 23:55:55' GROUP BY DATETIME) subquery1;";
								$AMF1_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `GR_AMF01_MS|EMSC` where DATETIME BETWEEN '".$new_date." 00:00:00' AND '".$new_date." 23:55:55' GROUP BY DATETIME) subquery1;";
								$AMF2_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_AMF02_MS|EMSC` where DATETIME BETWEEN '".$new_date." 00:00:00' AND '".$new_date." 23:55:55' GROUP BY DATETIME) subquery1;";

								$vSPGW1_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_vSPGW1_CALL|SESS` where DATETIME BETWEEN '".$new_date." 00:00:00' AND '".$new_date." 23:55:55' GROUP BY DATETIME) subquery1;";
								$vSPGW2_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_vSPGW2_CALL|SESS` where DATETIME BETWEEN '".$new_date." 00:00:00' AND '".$new_date." 23:55:55' GROUP BY DATETIME) subquery1;";
								$vSPGW3_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_vSPGW3_CALL|SESS` where DATETIME BETWEEN '".$new_date." 00:00:00' AND '".$new_date." 23:55:55' GROUP BY DATETIME) subquery1;";
								$SMF1_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_SMF01_SS_CALL|SESS` where DATETIME BETWEEN '".$new_date." 00:00:00' AND '".$new_date." 23:55:55' GROUP BY DATETIME) subquery1;";
								$SMF3_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_SMF03_SS_CALL|SESS` where DATETIME BETWEEN '".$new_date." 00:00:00' AND '".$new_date." 23:55:55' GROUP BY DATETIME) subquery1;";
								$UPF1_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_UPF01_CALL|SESS` where DATETIME BETWEEN '".$new_date." 00:00:00' AND '".$new_date." 23:55:55' GROUP BY DATETIME) subquery1;";
								$UPF3_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_UPF03_SS_CALL|SESS` where DATETIME BETWEEN '".$new_date." 00:00:00' AND '".$new_date." 23:55:55' GROUP BY DATETIME) subquery1;";
								$UPF4_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_UPF04_SS_CALL|SESS` where DATETIME BETWEEN '".$new_date." 00:00:00' AND '".$new_date." 23:55:55' GROUP BY DATETIME) subquery1;";
								$PGW50_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_PGW50_CALL|SESS` where DATETIME BETWEEN '".$new_date." 00:00:00' AND '".$new_date." 23:55:55' GROUP BY DATETIME) subquery1;";
								$PGW51_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_PGW51_CALL|SESS` where DATETIME BETWEEN '".$new_date." 00:00:00' AND '".$new_date." 23:55:55' GROUP BY DATETIME) subquery1;";
								$PGW14_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_PGW14_CALL|SESS` where DATETIME BETWEEN '".$new_date." 00:00:00' AND '".$new_date." 23:55:55' GROUP BY DATETIME) subquery1;";

								// 행을 한 줄씩 추가하면서, SQL 결과를 저장
								
                                $result2 = mysqli_query($conn, $MME1_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/2000000*100,2);

                                $sheetIndex->setCellValue('B'.$i,$SLOAD.' %');

                                $result2 = mysqli_query($conn, $MME2_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/2000000*100,2);

                                $sheetIndex->setCellValue('C'.$i,$SLOAD.' %');

                                $result2 = mysqli_query($conn, $vMME1_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/2000000*100,2);

                                $sheetIndex->setCellValue('D'.$i,$SLOAD.' %');

                                $result2 = mysqli_query($conn, $AMF1_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/2000000*100,2);

                                $sheetIndex->setCellValue('E'.$i,$SLOAD.' %');

                                $result2 = mysqli_query($conn, $AMF2_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/2000000*100,2);
								
                                $sheetIndex->setCellValue('F'.$i,$SLOAD.' %');

                                $result2 = mysqli_query($conn, $vSPGW1_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/2000000*100,2);

                                $sheetIndex->setCellValue('G'.$i,$SLOAD.' %');

                                $result2 = mysqli_query($conn, $vSPGW2_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/2000000*100,2);

                                $sheetIndex->setCellValue('H'.$i,$SLOAD.' %');

                                $result2 = mysqli_query($conn, $vSPGW3_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/2000000*100,2);

                                $sheetIndex->setCellValue('I'.$i,$SLOAD.' %');

                                $result2 = mysqli_query($conn, $PGW14_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/5000000*100,2);

                                $sheetIndex->setCellValue('J'.$i,$SLOAD.' %');

                                $result2 = mysqli_query($conn, $PGW50_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/150000*100,2);
								
                                $sheetIndex->setCellValue('K'.$i,$SLOAD.' %');
								
                                $result2 = mysqli_query($conn, $PGW51_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/220000*100,2);

                                $sheetIndex->setCellValue('L'.$i,$SLOAD.' %');

                                $result2 = mysqli_query($conn, $SMF1_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/8000000*100,2);

                                $sheetIndex->setCellValue('M'.$i,$SLOAD.' %');

                                $result2 = mysqli_query($conn, $SMF3_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/600000*100,2);

                                $sheetIndex->setCellValue('N'.$i,$SLOAD.' %');

                                $result2 = mysqli_query($conn, $UPF1_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/4000000*100,2);

                                $sheetIndex->setCellValue('O'.$i,$SLOAD.' %');           

                                $result2 = mysqli_query($conn, $UPF3_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/4000000*100,2);

                                $sheetIndex->setCellValue('P'.$i,$SLOAD.' %');

                                $result2 = mysqli_query($conn, $UPF4_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/4000000*100,2);

                                $sheetIndex->setCellValue('Q'.$i,$SLOAD.' %');


								if($new_date == $ENDDATE) break;
								$i=$i+1;


}

mysqli_close($conn);

// 엑셀파일로 저장

header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename='.$STARTDATE.'___'.$ENDDATE.'_세션 추이(삼성).xls');
header('Cache-Control: max-age=0');
 
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');
 
exit;

   }
 
?>
