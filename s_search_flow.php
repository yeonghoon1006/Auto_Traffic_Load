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
$sheetIndex->getStyle('B2:AO100')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$sheetIndex->getStyle('A1:A100')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$sheetIndex->getStyle('A1:AO1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$sheetIndex->setCellValue('A1','날짜');
$sheetIndex->setCellValue('B1','vSPGW#1');
$sheetIndex->setCellValue('C1','vSPGW#2');
$sheetIndex->setCellValue('D1','vSPGW#3');
$sheetIndex->setCellValue('E1','PGW#14');
$sheetIndex->setCellValue('F1','PGW#50');
$sheetIndex->setCellValue('G1','PGW#51');
$sheetIndex->setCellValue('H1','UPF#1');
$sheetIndex->setCellValue('I1','UPF#3');
$sheetIndex->setCellValue('J1','UPF#4');

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

                // MYSQL 에 접속해서 하루 씩 트래픽 DATA 파싱
  
								$vSPGW1_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.TRAFFIC) FROM (SELECT SYSTEM,SUM(UL_RX_PEAK_KBPS + DL_TX_PEAK_KBPS) AS TRAFFIC from `".$STAT."_vSPGW1_PUD|PKT` where DATETIME BETWEEN '".$new_date." 00:00:00' AND '".$new_date." 23:55:55' GROUP BY DATETIME) subquery1;";
								$vSPGW2_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.TRAFFIC) FROM (SELECT SYSTEM,SUM(UL_RX_PEAK_KBPS + DL_TX_PEAK_KBPS) AS TRAFFIC from `".$STAT."_vSPGW2_PUD|PKT` where DATETIME BETWEEN '".$new_date." 00:00:00' AND '".$new_date." 23:55:55' GROUP BY DATETIME) subquery1;";
								$vSPGW3_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.TRAFFIC) FROM (SELECT SYSTEM,SUM(UL_RX_PEAK_KBPS + DL_TX_PEAK_KBPS) AS TRAFFIC from `".$STAT."_vSPGW3_PUD|PKT` where DATETIME BETWEEN '".$new_date." 00:00:00' AND '".$new_date." 23:55:55' GROUP BY DATETIME) subquery1;";
								$UPF1_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.TRAFFIC) FROM (SELECT SYSTEM,SUM(UL_RX_PEAK_KBPS + DL_TX_PEAK_KBPS) AS TRAFFIC from `".$STAT."_UPF01_PUD|PKT` where DATETIME BETWEEN '".$new_date." 00:00:00' AND '".$new_date." 23:55:55' GROUP BY DATETIME) subquery1;";
								$UPF3_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.TRAFFIC) FROM (SELECT SYSTEM,SUM(UL_RX_PEAK_KBPS + DL_TX_PEAK_KBPS) AS TRAFFIC from `".$STAT."_UPF03_SS_PUD|PKT` where DATETIME BETWEEN '".$new_date." 00:00:00' AND '".$new_date." 23:55:55' GROUP BY DATETIME) subquery1;";
								$UPF4_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.TRAFFIC) FROM (SELECT SYSTEM,SUM(UL_RX_PEAK_KBPS + DL_TX_PEAK_KBPS) AS TRAFFIC from `".$STAT."_UPF04_SS_PUD|PKT` where DATETIME BETWEEN '".$new_date." 00:00:00' AND '".$new_date." 23:55:55' GROUP BY DATETIME) subquery1;";
								$PGW50_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.TRAFFIC) FROM (SELECT SYSTEM,SUM(UL_RX_PEAK_KBPS + DL_TX_PEAK_KBPS) AS TRAFFIC from `".$STAT."_PGW50_PUD|PKT` where DATETIME BETWEEN '".$new_date." 00:00:00' AND '".$new_date." 23:55:55' GROUP BY DATETIME) subquery1;";
								$PGW51_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.TRAFFIC) FROM (SELECT SYSTEM,SUM(UL_RX_PEAK_KBPS + DL_TX_PEAK_KBPS) AS TRAFFIC from `".$STAT."_PGW51_PUD|PKT` where DATETIME BETWEEN '".$new_date." 00:00:00' AND '".$new_date." 23:55:55' GROUP BY DATETIME) subquery1;";
								$PGW14_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.TRAFFIC) FROM (SELECT SYSTEM,SUM(UL_RX_PEAK_KBPS + DL_TX_PEAK_KBPS) AS TRAFFIC from `".$STAT."_PGW14_PUD|PKT` where DATETIME BETWEEN '".$new_date." 00:00:00' AND '".$new_date." 23:55:55' GROUP BY DATETIME) subquery1;";



                // 행을 한 줄씩 추가하면서, SQL 결과를 저장

								$sheetIndex->setCellValue('A'.$i,$new_date);

                                $result = mysqli_query($conn, $vSPGW1_TRAFFIC);
                                $row = mysqli_fetch_array($result);
                                $TEMP = $row['MAX(subquery1.TRAFFIC)']/1000000;
                                $LOAD = round($TEMP/50*100,2);

								$sheetIndex->setCellValue('B'.$i,$LOAD.' %');

                                $result = mysqli_query($conn, $vSPGW2_TRAFFIC);
                                $row = mysqli_fetch_array($result);
                                $TEMP = $row['MAX(subquery1.TRAFFIC)']/1000000;
                                $LOAD = round($TEMP/50*100,2);

								$sheetIndex->setCellValue('C'.$i,$LOAD.' %');

                                $result = mysqli_query($conn, $vSPGW3_TRAFFIC);
                                $row = mysqli_fetch_array($result);
                                $TEMP = $row['MAX(subquery1.TRAFFIC)']/1000000;
                                $LOAD = round($TEMP/50*100,2);

								$sheetIndex->setCellValue('D'.$i,$LOAD.' %');

                                $result = mysqli_query($conn, $PGW14_TRAFFIC);
                                $row = mysqli_fetch_array($result);
                                $TEMP = $row['MAX(subquery1.TRAFFIC)']/1000000;
                                $LOAD = round($TEMP/50*100,2);
								
								$sheetIndex->setCellValue('E'.$i,$LOAD.' %');
                                
								$result = mysqli_query($conn, $PGW50_TRAFFIC);
                                $row = mysqli_fetch_array($result);
                                $TEMP = $row['MAX(subquery1.TRAFFIC)']/1000000;
                                $LOAD = round($TEMP/10*100,2);

								$sheetIndex->setCellValue('F'.$i,$LOAD.' %');
								
                                $result = mysqli_query($conn, $PGW51_TRAFFIC);
                                $row = mysqli_fetch_array($result);
                                $TEMP = $row['MAX(subquery1.TRAFFIC)']/1000000;
                                $LOAD = round($TEMP/10*100,2);

								$sheetIndex->setCellValue('G'.$i,$LOAD.' %');                               

                                $result = mysqli_query($conn, $UPF1_TRAFFIC);
                                $row = mysqli_fetch_array($result);
                                $TEMP = $row['MAX(subquery1.TRAFFIC)']/1000000;
                                $LOAD = round($TEMP/80*100,2);

								$sheetIndex->setCellValue('H'.$i,$LOAD.' %'); 
								
                                $result = mysqli_query($conn, $UPF3_TRAFFIC);
                                $row = mysqli_fetch_array($result);
                                $TEMP = $row['MAX(subquery1.TRAFFIC)']/1000000;
                                $LOAD = round($TEMP/80*100,2);

								$sheetIndex->setCellValue('I'.$i,$LOAD.' %'); 

                                $result = mysqli_query($conn, $UPF4_TRAFFIC);
                                $row = mysqli_fetch_array($result);
                                $TEMP = $row['MAX(subquery1.TRAFFIC)']/1000000;
                                $LOAD = round($TEMP/40*100,2);
								
								$sheetIndex->setCellValue('J'.$i,$LOAD.' %'); 
                                
								if($new_date == $ENDDATE) break;
								$i=$i+1;


}

mysqli_close($conn);

  // 엑셀파일로 저장
  
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename='.$STARTDATE.'___'.$ENDDATE.'_트래픽 추이(삼성).xls');
header('Cache-Control: max-age=0');
 
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');
 
exit;

   }
 
?>
