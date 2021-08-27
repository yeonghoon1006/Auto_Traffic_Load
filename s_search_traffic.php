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

                $sheetIndex->setCellValue('A1','구로 호처리 트래픽 정보(삼성)');
                $sheetIndex->mergeCells('A1:D1');
                $sheetIndex->getStyle('A1')->getFont()->setSize(20)->setBold(true);
                $sheetIndex->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);


                $sheetIndex->setCellValue('A2','SYSTEM');
                $sheetIndex->setCellValue('B2','CAPACITY');
                $sheetIndex->setCellValue('C2','MAX Gbps');
                $sheetIndex->setCellValue('D2','MAX SESSION');
                $sheetIndex->setCellValue('E2','비고');

                $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(20);
                $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(35);
                $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(25);
                $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(30);
                $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(20);

                $sheetIndex->setCellValue('E3','재난망');
                $sheetIndex->setCellValue('E6','재난망,DATA');
                $sheetIndex->setCellValue('E7','DATA');
                $sheetIndex->setCellValue('E8','DATA');
                $sheetIndex->setCellValue('E9','VOLTE');
                $sheetIndex->setCellValue('E10','기업전용(전용회선)');
                $sheetIndex->setCellValue('E11','기업전용(인터넷)');
                $sheetIndex->setCellValue('E12','5G');
                $sheetIndex->setCellValue('E13','LTE');
                $sheetIndex->setCellValue('E14','5G');
                $sheetIndex->setCellValue('E15','5G');
                $sheetIndex->setCellValue('E16','LTE');

                $TD_COLOR = array(

                        //배경색 설정
                        'fill' => array(
                         'type' => PHPExcel_Style_Fill::FILL_SOLID,
                         'color' => array('rgb'=>'888888'),
                        ),

                        //글자색 설정
                        'font' => array(
                         'bold' => 'true',
                         'size' => '12',
                         'color' => array('rgb'=>'FFFFFF')
                        ),

                        //테두리 설정
                        'borders' => array(
                         'allborders' => array(
                          'style' => PHPExcel_Style_Border::BORDER_THIN,
                          'color' => array('argb'=>'000000')
                         )
                        ),

                );
				$TD_COLOR2 = array(

					//배경색 설정
					'fill' => array(
					 'type' => PHPExcel_Style_Fill::FILL_SOLID,
					 'color' => array('rgb'=>'FFFFFF'),
					),

					//글자색 설정
					'font' => array(
					 'bold' => 'true',
					 'size' => '12',
					 'color' => array('rgb'=>'000000')
					),

					//테두리 설정
					'borders' => array(
					 'allborders' => array(
					  'style' => PHPExcel_Style_Border::BORDER_THIN,
					  'color' => array('argb'=>'000000')
					 )
					),
				);


				$objPHPExcel->getActiveSheet()->getStyle("A3:E18")->applyFromArray($TD_COLOR2);
                $objPHPExcel->getActiveSheet()->getStyle("A2:E2")->applyFromArray($TD_COLOR);
                $objPHPExcel->getActiveSheet()->getStyle("A2:E18")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				
				// MYSQL 에 접속해서 시작날짜, 종료날짜 사이의 원하는 DATA 파싱(트래픽, 세션)
				
                $MME1_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_A-MME1_MS|EMSC` where DATETIME BETWEEN '".$STARTDATE."' AND '".$ENDDATE."' GROUP BY DATETIME) subquery1;";
                $MME2_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_A-MME2_MS|EMSC` where DATETIME BETWEEN '".$STARTDATE."' AND '".$ENDDATE."' GROUP BY DATETIME) subquery1;";
                $vMME1_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_vMME1_MS|EMSC` where DATETIME BETWEEN '".$STARTDATE."' AND '".$ENDDATE."' GROUP BY DATETIME) subquery1;";
                $AMF1_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `GR_AMF01_MS|EMSC` where DATETIME BETWEEN '".$STARTDATE."' AND '".$ENDDATE."' GROUP BY DATETIME) subquery1;";
                $AMF2_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_AMF02_MS|EMSC` where DATETIME BETWEEN '".$STARTDATE."' AND '".$ENDDATE."' GROUP BY DATETIME) subquery1;";


                $vSPGW1_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.TRAFFIC) FROM (SELECT SYSTEM,SUM(UL_RX_PEAK_KBPS + DL_TX_PEAK_KBPS) AS TRAFFIC from `".$STAT."_vSPGW1_PUD|PKT` where DATETIME BETWEEN '".$STARTDATE."' AND '".$ENDDATE."' GROUP BY DATETIME) subquery1;";
                $vSPGW2_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.TRAFFIC) FROM (SELECT SYSTEM,SUM(UL_RX_PEAK_KBPS + DL_TX_PEAK_KBPS) AS TRAFFIC from `".$STAT."_vSPGW2_PUD|PKT` where DATETIME BETWEEN '".$STARTDATE."' AND '".$ENDDATE."' GROUP BY DATETIME) subquery1;";
                $vSPGW3_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.TRAFFIC) FROM (SELECT SYSTEM,SUM(UL_RX_PEAK_KBPS + DL_TX_PEAK_KBPS) AS TRAFFIC from `".$STAT."_vSPGW3_PUD|PKT` where DATETIME BETWEEN '".$STARTDATE."' AND '".$ENDDATE."' GROUP BY DATETIME) subquery1;";
                $UPF1_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.TRAFFIC) FROM (SELECT SYSTEM,SUM(UL_RX_PEAK_KBPS + DL_TX_PEAK_KBPS) AS TRAFFIC from `".$STAT."_UPF01_PUD|PKT` where DATETIME BETWEEN '".$STARTDATE."' AND '".$ENDDATE."' GROUP BY DATETIME) subquery1;";
                $UPF3_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.TRAFFIC) FROM (SELECT SYSTEM,SUM(UL_RX_PEAK_KBPS + DL_TX_PEAK_KBPS) AS TRAFFIC from `".$STAT."_UPF03_SS_PUD|PKT` where DATETIME BETWEEN '".$STARTDATE."' AND '".$ENDDATE."' GROUP BY DATETIME) subquery1;";
                $UPF4_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.TRAFFIC) FROM (SELECT SYSTEM,SUM(UL_RX_PEAK_KBPS + DL_TX_PEAK_KBPS) AS TRAFFIC from `".$STAT."_UPF04_SS_PUD|PKT` where DATETIME BETWEEN '".$STARTDATE."' AND '".$ENDDATE."' GROUP BY DATETIME) subquery1;";
                $PGW50_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.TRAFFIC) FROM (SELECT SYSTEM,SUM(UL_RX_PEAK_KBPS + DL_TX_PEAK_KBPS) AS TRAFFIC from `".$STAT."_PGW50_PUD|PKT` where DATETIME BETWEEN '".$STARTDATE."' AND '".$ENDDATE."' GROUP BY DATETIME) subquery1;";
                $PGW51_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.TRAFFIC) FROM (SELECT SYSTEM,SUM(UL_RX_PEAK_KBPS + DL_TX_PEAK_KBPS) AS TRAFFIC from `".$STAT."_PGW51_PUD|PKT` where DATETIME BETWEEN '".$STARTDATE."' AND '".$ENDDATE."' GROUP BY DATETIME) subquery1;";
                                $PGW14_TRAFFIC = "SELECT SYSTEM,MAX(subquery1.TRAFFIC) FROM (SELECT SYSTEM,SUM(UL_RX_PEAK_KBPS + DL_TX_PEAK_KBPS) AS TRAFFIC from `".$STAT."_PGW14_PUD|PKT` where DATETIME BETWEEN '".$STARTDATE."' AND '".$ENDDATE."' GROUP BY DATETIME) subquery1;";

                $vSPGW1_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_vSPGW1_CALL|SESS` where DATETIME BETWEEN '".$STARTDATE."' AND '".$ENDDATE."' GROUP BY DATETIME) subquery1;";
                $vSPGW2_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_vSPGW2_CALL|SESS` where DATETIME BETWEEN '".$STARTDATE."' AND '".$ENDDATE."' GROUP BY DATETIME) subquery1;";
                $vSPGW3_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_vSPGW3_CALL|SESS` where DATETIME BETWEEN '".$STARTDATE."' AND '".$ENDDATE."' GROUP BY DATETIME) subquery1;";
                $SMF1_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_SMF01_SS_CALL|SESS` where DATETIME BETWEEN '".$STARTDATE."' AND '".$ENDDATE."' GROUP BY DATETIME) subquery1;";
                $SMF3_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_SMF03_SS_CALL|SESS` where DATETIME BETWEEN '".$STARTDATE."' AND '".$ENDDATE."' GROUP BY DATETIME) subquery1;";
                $UPF1_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_UPF01_CALL|SESS` where DATETIME BETWEEN '".$STARTDATE."' AND '".$ENDDATE."' GROUP BY DATETIME) subquery1;";
                $UPF3_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_UPF03_SS_CALL|SESS` where DATETIME BETWEEN '".$STARTDATE."' AND '".$ENDDATE."' GROUP BY DATETIME) subquery1;";
                $UPF4_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_UPF04_SS_CALL|SESS` where DATETIME BETWEEN '".$STARTDATE."' AND '".$ENDDATE."' GROUP BY DATETIME) subquery1;";
                $PGW50_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_PGW50_CALL|SESS` where DATETIME BETWEEN '".$STARTDATE."' AND '".$ENDDATE."' GROUP BY DATETIME) subquery1;";
                $PGW51_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_PGW51_CALL|SESS` where DATETIME BETWEEN '".$STARTDATE."' AND '".$ENDDATE."' GROUP BY DATETIME) subquery1;";
                $PGW14_SESSION = "SELECT SYSTEM,MAX(subquery1.SESS) FROM (SELECT SYSTEM,SUM(PEAK_TOTAL) AS SESS from `".$STAT."_PGW14_CALL|SESS` where DATETIME BETWEEN '".$STARTDATE."' AND '".$ENDDATE."' GROUP BY DATETIME) subquery1;";

				 // 같은 범주에 들어가는 장비 별 배열 선언
                $MME_LTE_SESSION=array();
                $vSPGW_LTE_SESSION=array();
                $vSPGW_LTE_TRAFFIC=array();
                $SMF_LTE_SESSION=array();
                $UPF_LTE_SESSION=array();
                $UPF_LTE_TRAFFIC=array();
                $PGW50_LTE_TRAFFIC=array();
                $PGW50_LTE_SESSION=array();
                $PGW51_LTE_TRAFFIC=array();
                $PGW51_LTE_SESSION=array();
                $PGW14_LTE_SESSION=array();
                $AMF_5G_SESSION=array();
                $SMF_5G_SESSION=array();
                $UPF_5G_TRAFFIC=array();
                $UPF_5G_SESSION=array();

				// 각 장비 별 트래픽, 세션 data를 파싱하여 최고 부하를 계산하고 엑셀 및 배열에 저장
				
                                $conn = mysqli_connect(
                                  '172.21.223.167',
                                  'guro',
                                  'Drcore12345!',
                                  'MACA_RAW');
                                  

                                $result2 = mysqli_query($conn, $MME1_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/2000000*100,2);

                                $sheetIndex ->setCellValue('A3', 'GURO_MME#1');
                                $sheetIndex ->setCellValue('B3', '2,000,000 SESSION ');
                                $sheetIndex ->setCellValue('D3', number_format($STEMP).' SESS ('.$SLOAD.'%)');
                                $MME_LTE_SESSION['MME#1'] = $SLOAD;

                                $result2 = mysqli_query($conn, $MME2_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/2000000*100,2);

                                $sheetIndex ->setCellValue('A4', 'GURO_MME#2');
                                $sheetIndex ->setCellValue('B4', '2,000,000 SESSION ');
                                $sheetIndex ->setCellValue('D4', number_format($STEMP).' SESS ('.$SLOAD.'%)');
                                $MME_LTE_SESSION['MME#2'] = $SLOAD;

                                $result2 = mysqli_query($conn, $vMME1_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/2000000*100,2);

                                $sheetIndex ->setCellValue('A5', 'GURO_vMME#1');
                                $sheetIndex ->setCellValue('B5', '2,000,000 SESSION ');
                                $sheetIndex ->setCellValue('D5', number_format($STEMP).' SESS ('.$SLOAD.'%)');
                                $MME_LTE_SESSION['vMME#1'] = $SLOAD;


                                $result = mysqli_query($conn, $vSPGW1_TRAFFIC);
                                $row = mysqli_fetch_array($result);
                                $TEMP = $row['MAX(subquery1.TRAFFIC)']/1000000;
                                $LOAD = round($TEMP/50*100,2);

                                $result2 = mysqli_query($conn, $vSPGW1_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/2000000*100,2);

                                $sheetIndex ->setCellValue('A6', 'vSPGW#1');
                                $sheetIndex ->setCellValue('B6', '50 Gbps / 2,000,000 SESSION ');
                                $sheetIndex ->setCellValue('C6', number_format($TEMP).' Gbps ('.$LOAD.'%)');
                                $sheetIndex ->setCellValue('D6', number_format($STEMP).' SESS ('.$SLOAD.'%)');
                                $vSPGW_LTE_SESSION['vSPGW#1'] = $LOAD;
                                $vSPGW_LTE_TRAFFIC['vSPGW#1'] = $SLOAD;


                                $result = mysqli_query($conn, $vSPGW2_TRAFFIC);
                                $row = mysqli_fetch_array($result);
                                $TEMP = $row['MAX(subquery1.TRAFFIC)']/1000000;
                                $LOAD = round($TEMP/50*100,2);

                                $result2 = mysqli_query($conn, $vSPGW2_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/2000000*100,2);

                                $sheetIndex ->setCellValue('A7', 'vSPGW#2');
                                $sheetIndex ->setCellValue('B7', '50 Gbps / 2,000,000 SESSION ');
                                $sheetIndex ->setCellValue('C7', number_format($TEMP).' Gbps ('.$LOAD.'%)');
                                $sheetIndex ->setCellValue('D7', number_format($STEMP).' SESS ('.$SLOAD.'%)');
                                $vSPGW_LTE_SESSION['vSPGW#2'] = $SLOAD;
                                $vSPGW_LTE_TRAFFIC['vSPGW#2'] = $LOAD;

                                $result = mysqli_query($conn, $vSPGW3_TRAFFIC);
                                $row = mysqli_fetch_array($result);
                                $TEMP = $row['MAX(subquery1.TRAFFIC)']/1000000;
                                $LOAD = round($TEMP/50*100,2);

                                $result2 = mysqli_query($conn, $vSPGW3_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/2000000*100,2);

                                $sheetIndex ->setCellValue('A8', 'vSPGW#3');
                                $sheetIndex ->setCellValue('B8', '50 Gbps / 2,000,000 SESSION ');
                                $sheetIndex ->setCellValue('C8', number_format($TEMP).' Gbps ('.$LOAD.'%)');
                                $sheetIndex ->setCellValue('D8', number_format($STEMP).' SESS ('.$SLOAD.'%)');
                                $vSPGW_LTE_SESSION['vSPGW#3'] = $SLOAD;
                                $vSPGW_LTE_TRAFFIC['vSPGW#3'] = $LOAD;

                                $result = mysqli_query($conn, $PGW14_TRAFFIC);
                                $row = mysqli_fetch_array($result);
                                $TEMP = $row['MAX(subquery1.TRAFFIC)']/1000000;
                                $LOAD = round($TEMP/50*100,2);

                                $result2 = mysqli_query($conn, $PGW14_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/5000000*100,2);

                                $sheetIndex ->setCellValue('A9', 'PGW#14');
                                $sheetIndex ->setCellValue('B9', '50 Gbps / 5,000,000 SESSION ');
                                $sheetIndex ->setCellValue('C9', number_format($TEMP).' Gbps ('.$LOAD.'%)');
                                $sheetIndex ->setCellValue('D9', number_format($STEMP).' SESS ('.$SLOAD.'%)');
                                $PGW14_LTE_SESSION['PGW#14'] = $SLOAD;


                                $result = mysqli_query($conn, $PGW50_TRAFFIC);
                                $row = mysqli_fetch_array($result);
                                $TEMP = $row['MAX(subquery1.TRAFFIC)']/1000000;
                                $LOAD = round($TEMP/10*100,2);

                                $result2 = mysqli_query($conn, $PGW50_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/150000*100,2);

                                $sheetIndex ->setCellValue('A10', 'PGW#50');
                                $sheetIndex ->setCellValue('B10', '10 Gbps / 150,000 SESSION ');
                                $sheetIndex ->setCellValue('C10', number_format($TEMP).' Gbps ('.$LOAD.'%)');
                                $sheetIndex ->setCellValue('D10', number_format($STEMP).' SESS ('.$SLOAD.'%)');
                                $PGW50_LTE_SESSION['PGW#50'] = $SLOAD;
                                $PGW50_LTE_TRAFFIC['PGW#50'] = $LOAD;

                                $result = mysqli_query($conn, $PGW51_TRAFFIC);
                                $row = mysqli_fetch_array($result);
                                $TEMP = $row['MAX(subquery1.TRAFFIC)']/1000000;
                                $LOAD = round($TEMP/10*100,2);

                                $result2 = mysqli_query($conn, $PGW51_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/220000*100,2);

                                $sheetIndex ->setCellValue('A11', 'PGW#51');
                                $sheetIndex ->setCellValue('B11', '10 Gbps / 220,000 SESSION ');
                                $sheetIndex ->setCellValue('C11', number_format($TEMP).' Gbps ('.$LOAD.'%)');
                                $sheetIndex ->setCellValue('D11', number_format($STEMP).' SESS ('.$SLOAD.'%)');
                                $PGW51_LTE_SESSION['PGW#51'] = $SLOAD;
                                $PGW51_LTE_TRAFFIC['PGW#51'] = $LOAD;


                                $result2 = mysqli_query($conn, $SMF1_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/8000000*100,2);

                                $sheetIndex ->setCellValue('A12', 'GURO_SMF#1');
                                $sheetIndex ->setCellValue('B12', '600,000 SESSION ');
                                $sheetIndex ->setCellValue('D12', number_format($STEMP).' SESS ('.$SLOAD.'%)');
                                $SMF_5G_SESSION['SMF#1'] = $SLOAD;

                                $result2 = mysqli_query($conn, $SMF3_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/600000*100,2);

                                $sheetIndex ->setCellValue('A13', 'GURO_SMF#3');
                                $sheetIndex ->setCellValue('B13', '600,000 SESSION ');
                                $sheetIndex ->setCellValue('D13', number_format($STEMP).' SESS ('.$SLOAD.'%)');
                                $SMF_LTE_SESSION['SMF#3'] = $SLOAD;

                                $result = mysqli_query($conn, $UPF1_TRAFFIC);
                                $row = mysqli_fetch_array($result);
                                $TEMP = $row['MAX(subquery1.TRAFFIC)']/1000000;
                                $LOAD = round($TEMP/80*100,2);

                                $result2 = mysqli_query($conn, $UPF1_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/4000000*100,2);

                                $sheetIndex ->setCellValue('A14', 'UPF#1');
                                $sheetIndex ->setCellValue('B14', '80 Gbps / 4,000,000 SESSION ');
                                $sheetIndex ->setCellValue('C14', number_format($TEMP).' Gbps ('.$LOAD.'%)');
                                $sheetIndex ->setCellValue('D14', number_format($STEMP).' SESS ('.$SLOAD.'%)');
                                $UPF_5G_SESSION['UPF#1'] = $SLOAD;
                                $UPF_5G_TRAFFIC['UPF#1'] = $LOAD;
/*								철거 장비
								
                                $result = mysqli_query($conn, $UPF3_TRAFFIC);
                                $row = mysqli_fetch_array($result);
                                $TEMP = $row['MAX(subquery1.TRAFFIC)']/1000000;
                                $LOAD = round($TEMP/80*100,2);

                                $result2 = mysqli_query($conn, $UPF3_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/4000000*100,2);

                                $sheetIndex ->setCellValue('A15', 'UPF#3');
                                $sheetIndex ->setCellValue('B15', '80 Gbps / 4,000,000 SESSION ');
                                $sheetIndex ->setCellValue('C15', number_format($TEMP).' Gbps ('.$LOAD.'%)');
                                $sheetIndex ->setCellValue('D15', number_format($STEMP).' SESS ('.$SLOAD.'%)');
                                $UPF_5G_SESSION['UPF#3'] = $SLOAD;
                                $UPF_5G_TRAFFIC['UPF#3'] = $LOAD;
*/

                                $result = mysqli_query($conn, $UPF4_TRAFFIC);
                                $row = mysqli_fetch_array($result);
                                $TEMP = $row['MAX(subquery1.TRAFFIC)']/1000000;
                                $LOAD = round($TEMP/90*100,2);

                                $result2 = mysqli_query($conn, $UPF4_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/4000000*100,2);

                                $sheetIndex ->setCellValue('A16', 'UPF#4');
                                $sheetIndex ->setCellValue('B16', '90 Gbps / 4,000,000 SESSION ');
                                $sheetIndex ->setCellValue('C16', number_format($TEMP).' Gbps ('.$LOAD.'%)');
                                $sheetIndex ->setCellValue('D16', number_format($STEMP).' SESS ('.$SLOAD.'%)');
                                $UPF_LTE_SESSION['UPF#4'] = $SLOAD;
                                $UPF_LTE_TRAFFIC['UPF#4'] = $LOAD;

                                $result2 = mysqli_query($conn, $AMF1_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/2000000*100,2);

                                $sheetIndex ->setCellValue('A17', 'GURO_AMF#1');
                                $sheetIndex ->setCellValue('B17', '2,000,000 SESSION ');
                                $sheetIndex ->setCellValue('D17', number_format($STEMP).' SESS ('.$SLOAD.'%)');
                                $AMF_5G_SESSION['AMF#1'] = $SLOAD;

                                $result2 = mysqli_query($conn, $AMF2_SESSION);
                                $row2 = mysqli_fetch_array($result2);
                                $STEMP = $row2['MAX(subquery1.SESS)'];
                                $SLOAD = round($STEMP/2000000*100,2);

                                $sheetIndex ->setCellValue('A18', 'GURO_AMF#2');
                                $sheetIndex ->setCellValue('B18', '2,000,000 SESSION ');
                                $sheetIndex ->setCellValue('D18', number_format($STEMP).' SESS ('.$SLOAD.'%)');
                                $AMF_5G_SESSION['AMF#2'] = $SLOAD;





				// 같은 범주에 있는 장비 중 가장 높은 부하율을 가지는 장비 요약 시트 생성
				
				$myWorkSheet = new PHPExcel_Worksheet($objPHPExcel, 'Summary');
				$objPHPExcel->addSheet($myWorkSheet, 0);

				$sheetIndex2 = $objPHPExcel->setActiveSheetIndex(0);
				$sheet = $objPHPExcel->getActiveSheet();

				$sheetIndex2->mergeCells('A1:G1');
				$sheetIndex2->mergeCells('A3:A13');
				$sheetIndex2->mergeCells('A14:A17');
				$sheetIndex2->mergeCells('B4:B5');
				$sheetIndex2->mergeCells('B7:B8');
				$sheetIndex2->mergeCells('B9:B10');
				$sheetIndex2->mergeCells('B11:B12');
				$sheetIndex2->mergeCells('B16:B17');
				$sheetIndex2->mergeCells('C4:C5');
				$sheetIndex2->mergeCells('C7:C8');
				$sheetIndex2->mergeCells('C9:C10');
				$sheetIndex2->mergeCells('C11:C12');
				$sheetIndex2->mergeCells('C16:C17');				
				
				$sheetIndex2->setCellValue('A2','구분');
				$sheetIndex2->setCellValue('B2','시스템');
				$sheetIndex2->setCellValue('C2','시설 수');
				$sheetIndex2->setCellValue('D2','부하기준');
				$sheetIndex2->setCellValue('E2','최고부하');
				$sheetIndex2->setCellValue('F2','부하율');
				$sheetIndex2->setCellValue('G2','비고');

				$sheetIndex2->setCellValue('A3','LTE');
				$sheetIndex2->setCellValue('A14','5G');

				$sheetIndex2->setCellValue('B3','MME/vMME');
				$sheetIndex2->setCellValue('B4','vSPGW');
				$sheetIndex2->setCellValue('B6','SMF#3(LTE)');
				$sheetIndex2->setCellValue('B7','UPF#4(LTE)');
				$sheetIndex2->setCellValue('B9','PGW#50(기업전용 전용회선)');
				$sheetIndex2->setCellValue('B11','PGW#51(기업전용 인터넷)');
				$sheetIndex2->setCellValue('B13','PGW#14(VOLTE)');
				$sheetIndex2->setCellValue('B14','AMF');
				$sheetIndex2->setCellValue('B15','SMF(5G)');
				$sheetIndex2->setCellValue('B16','UPF(5G)');


				$sheetIndex2->setCellValue('D3','세션수');
				$sheetIndex2->setCellValue('D4','트래픽');
				$sheetIndex2->setCellValue('D5','세션수');
				$sheetIndex2->setCellValue('D6','세션수');
				$sheetIndex2->setCellValue('D7','트래픽');
				$sheetIndex2->setCellValue('D8','세션수');
				$sheetIndex2->setCellValue('D9','트래픽');
				$sheetIndex2->setCellValue('D10','세션수');
				$sheetIndex2->setCellValue('D11','트래픽');
				$sheetIndex2->setCellValue('D12','세션수');
				$sheetIndex2->setCellValue('D13','세션수');
				$sheetIndex2->setCellValue('D14','세션수');
				$sheetIndex2->setCellValue('D15','세션수');
				$sheetIndex2->setCellValue('D16','트래픽');
				$sheetIndex2->setCellValue('D17','세션수');

				$sheetIndex2->setCellValue('A1','ㅇ 호처리(삼성)');


				// 배열을 이용하여 data 표기
				
				$sheetIndex2->setCellValue('C3',count($MME_LTE_SESSION));
				$sheetIndex2->setCellValue('C4',count($vSPGW_LTE_SESSION));
				$sheetIndex2->setCellValue('C6',count($SMF_LTE_SESSION));
				$sheetIndex2->setCellValue('C7',count($UPF_LTE_SESSION));
				$sheetIndex2->setCellValue('C9',count($PGW50_LTE_SESSION));
				$sheetIndex2->setCellValue('C11',count($PGW51_LTE_SESSION));
				$sheetIndex2->setCellValue('C13',count($PGW14_LTE_SESSION));
				$sheetIndex2->setCellValue('C14',count($AMF_5G_SESSION));
				$sheetIndex2->setCellValue('C15',count($SMF_5G_SESSION));
			    $sheetIndex2->setCellValue('C16',count($UPF_5G_SESSION));

				$sheetIndex2->setCellValue('E3',array_search(max($MME_LTE_SESSION),$MME_LTE_SESSION));
				$sheetIndex2->setCellValue('E4',array_search(max($vSPGW_LTE_TRAFFIC),$vSPGW_LTE_TRAFFIC));
				$sheetIndex2->setCellValue('E5',array_search(max($vSPGW_LTE_SESSION),$vSPGW_LTE_SESSION));
				$sheetIndex2->setCellValue('E6',array_search(max($SMF_LTE_SESSION),$SMF_LTE_SESSION));
				$sheetIndex2->setCellValue('E7',array_search(max($UPF_LTE_TRAFFIC),$UPF_LTE_TRAFFIC));
				$sheetIndex2->setCellValue('E8',array_search(max($UPF_LTE_SESSION),$UPF_LTE_SESSION));
				$sheetIndex2->setCellValue('E9',array_search(max($PGW50_LTE_TRAFFIC),$PGW50_LTE_TRAFFIC));
				$sheetIndex2->setCellValue('E10',array_search(max($PGW50_LTE_SESSION),$PGW50_LTE_SESSION));
				$sheetIndex2->setCellValue('E11',array_search(max($PGW51_LTE_TRAFFIC),$PGW51_LTE_TRAFFIC));
				$sheetIndex2->setCellValue('E12',array_search(max($PGW51_LTE_SESSION),$PGW51_LTE_SESSION));
				$sheetIndex2->setCellValue('E13',array_search(max($PGW14_LTE_SESSION),$PGW14_LTE_SESSION));
				$sheetIndex2->setCellValue('E14',array_search(max($AMF_5G_SESSION),$AMF_5G_SESSION));								
				$sheetIndex2->setCellValue('E15',array_search(max($SMF_5G_SESSION),$SMF_5G_SESSION));
				$sheetIndex2->setCellValue('E16',array_search(max($UPF_5G_TRAFFIC),$UPF_5G_TRAFFIC));								
				$sheetIndex2->setCellValue('E17',array_search(max($UPF_5G_SESSION),$UPF_5G_SESSION));


				$sheetIndex2->setCellValue('F3',max($MME_LTE_SESSION).'%');
				$sheetIndex2->setCellValue('F4',max($vSPGW_LTE_TRAFFIC).'%');
				$sheetIndex2->setCellValue('F5',max($vSPGW_LTE_SESSION).'%');
				$sheetIndex2->setCellValue('F6',max($SMF_LTE_SESSION).'%');
				$sheetIndex2->setCellValue('F7',max($UPF_LTE_TRAFFIC).'%');
				$sheetIndex2->setCellValue('F8',max($UPF_LTE_SESSION).'%');
				$sheetIndex2->setCellValue('F9',max($PGW50_LTE_TRAFFIC).'%');
				$sheetIndex2->setCellValue('F10',max($PGW50_LTE_SESSION).'%');
				$sheetIndex2->setCellValue('F11',max($PGW51_LTE_TRAFFIC).'%');
				$sheetIndex2->setCellValue('F12',max($PGW51_LTE_SESSION).'%');
				$sheetIndex2->setCellValue('F13',max($PGW14_LTE_SESSION).'%');
				$sheetIndex2->setCellValue('F14',max($AMF_5G_SESSION).'%');
				$sheetIndex2->setCellValue('F15',max($SMF_5G_SESSION).'%');
				$sheetIndex2->setCellValue('F16',max($UPF_5G_TRAFFIC).'%');
				$sheetIndex2->setCellValue('F17',max($UPF_5G_SESSION).'%');


				$sheetIndex2->setCellValue('G3','용량 : 200만');
				$sheetIndex2->setCellValue('G4','용량 : 200만');
				$sheetIndex2->setCellValue('G5','용량 : 50G');
				$sheetIndex2->setCellValue('G6','용량 : 60만');
				$sheetIndex2->setCellValue('G7','용량 : 90G');
				$sheetIndex2->setCellValue('G8','용량 : 400만');
				$sheetIndex2->setCellValue('G9','용량 : 10G');
				$sheetIndex2->setCellValue('G10','용량 : 15만');
				$sheetIndex2->setCellValue('G11','용량 : 10G');
				$sheetIndex2->setCellValue('G12','용량 : 22만');
				$sheetIndex2->setCellValue('G13','용량 : 500만');
				$sheetIndex2->setCellValue('G14','용량 : 200만');
				$sheetIndex2->setCellValue('G15','용량 : 800만');
				$sheetIndex2->setCellValue('G16','용량 : 80G');
				$sheetIndex2->setCellValue('G17','용량 : 400만');



				$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(10);
				$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(30);
				$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(10);
				$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(20);
				$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(20);
				$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(20);
				$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(35);

				$sheetIndex2->getStyle('A1')->getFont()->setSize(15)->setBold(true);

				$objPHPExcel->getActiveSheet()->getStyle("A2:G17")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$objPHPExcel->getActiveSheet()->getStyle("A2:G17")->getAlignment ()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);


				$TD_COLOR3 = array(

					//배경색 설정
					'fill' => array(
					 'type' => PHPExcel_Style_Fill::FILL_SOLID,
					 'color' => array('rgb'=>'FFFFFF'),
					),

					//글자색 설정
					'font' => array(
					 'size' => '12',
					 'color' => array('rgb'=>'000000')
					),

					//테두리 설정
					'borders' => array(
					 'allborders' => array(
					  'style' => PHPExcel_Style_Border::BORDER_THIN,
					  'color' => array('argb'=>'000000')
					 )
					),
				);

				$TD_COLOR4 = array(

					//배경색 설정
					'fill' => array(
					 'type' => PHPExcel_Style_Fill::FILL_SOLID,
					 'color' => array('rgb'=>'FFFFFF'),
					),

					//글자색 설정
					'font' => array(
					 'bold' => 'true',
					 'size' => '12',
					 'color' => array('rgb'=>'FF0000')
					),

					//테두리 설정
					'borders' => array(
					 'allborders' => array(
					  'style' => PHPExcel_Style_Border::BORDER_THIN,
					  'color' => array('argb'=>'000000')
					 )
					),
				);

				$objPHPExcel->getActiveSheet()->getStyle("A2:G2")->applyFromArray($TD_COLOR);
				$objPHPExcel->getActiveSheet()->getStyle("A3:G17")->applyFromArray($TD_COLOR3);		
				
				 // 부하가 80% 이상일 경우 색상 표시
				 
				if(max($MME_LTE_SESSION)>=80)
				{
						$objPHPExcel->getActiveSheet()->getStyle("F3")->applyFromArray($TD_COLOR4);
				}
				if(max($vSPGW_LTE_SESSION)>=80)
				{
						$objPHPExcel->getActiveSheet()->getStyle("F4")->applyFromArray($TD_COLOR4);
				}
				if(max($vSPGW_LTE_SESSION)>=80)
				{
						$objPHPExcel->getActiveSheet()->getStyle("F5")->applyFromArray($TD_COLOR4);
				}
				if(max($SMF_LTE_SESSION)>=80)
				{
						$objPHPExcel->getActiveSheet()->getStyle("F6")->applyFromArray($TD_COLOR4);
				}
				if(max($UPF_LTE_TRAFFIC)>=80)
				{
						$objPHPExcel->getActiveSheet()->getStyle("F7")->applyFromArray($TD_COLOR4);
				}
				if(max($UPF_LTE_SESSION)>=80)
				{
						$objPHPExcel->getActiveSheet()->getStyle("F8")->applyFromArray($TD_COLOR4);
				}				
				if(max($PGW50_LTE_TRAFFIC)>=80)
				{
						$objPHPExcel->getActiveSheet()->getStyle("F9")->applyFromArray($TD_COLOR4);
				}
				if(max($PGW50_LTE_SESSION)>=80)
				{
						$objPHPExcel->getActiveSheet()->getStyle("F10")->applyFromArray($TD_COLOR4);
				}
				if(max($PGW51_LTE_TRAFFIC)>=80)
				{
						$objPHPExcel->getActiveSheet()->getStyle("F11")->applyFromArray($TD_COLOR4);
				}					
				if(max($PGW51_LTE_SESSION)>=80)
				{
						$objPHPExcel->getActiveSheet()->getStyle("F12")->applyFromArray($TD_COLOR4);
				}
				if(max($PGW14_LTE_SESSION)>=80)
				{
						$objPHPExcel->getActiveSheet()->getStyle("F13")->applyFromArray($TD_COLOR4);
				}
				if(max($AMF_5G_SESSION)>=80)
				{
						$objPHPExcel->getActiveSheet()->getStyle("F14")->applyFromArray($TD_COLOR4);
				}				
				if(max($SMF_5G_SESSION)>=80)
				{
						$objPHPExcel->getActiveSheet()->getStyle("F15")->applyFromArray($TD_COLOR4);
				}
				if(max($UPF_5G_TRAFFIC)>=80)
				{
						$objPHPExcel->getActiveSheet()->getStyle("F16")->applyFromArray($TD_COLOR4);
				}
				if(max($UPF_5G_SESSION)>=80)
				{
						$objPHPExcel->getActiveSheet()->getStyle("F17")->applyFromArray($TD_COLOR4);
				}	


				
                mysqli_close($conn);
				
				// 엑셀로 저장
				
                header('Content-Type: application/vnd.ms-excel');
                header('Content-Disposition: attachment;filename='.$STARTDATE.'___'.$ENDDATE.'___'.$STAT.'(삼성).xls');
                header('Cache-Control: max-age=0');
 
                $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
                $objWriter->save('php://output');
 
                exit;
        }

?>
