<?php
	
	ini_set('memory_limit','500M');
	set_time_limit(0);
	header("Content-type: text/html; charset=gb2312");
	require_once dirname(__FILE__) . '/Classes/PHPExcel.php';

	$singleUvList = include('oct_uv_s.php');  //单个UV
	$totalUvList  = include('oct_uv_t.php');  //总UV
	$imgList      = include('oct_img.php');   //图片
	$oct_keyword  = include('oct_faxian.php');//关键词
	$oct_date     = include('oct_faxian_date.php');  // 时间

	$objPHPExcel = new PHPExcel();
	$cols = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];

	$j=1;
	for ($i='A'; $i <= 'Z' ; $i++) { 
		if($j > 31) break;
		$date[$i] = $j;
		$j++;
	}

	for ($i=0; $i <count($oct_keyword) ; $i++) { 
		doCreatExcel('2016-10',$oct_keyword[$i],$cols[$i]);
	}
	header('Content-Type: application/vnd.ms-excel');
	header('Content-Disposition: attachment;filename="testtt点击率.xls"');
	header('Cache-Control: max-age=0');
	header('Cache-Control: max-age=1');

	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
	$objWriter->save('php://output');

	function getRate($month,$day,$num){
		global $totalUvList;
		global $singleUvList;

		$idx = $month.'-'.$day; 
		if(strlen($day) == 1){
			$idx = $month.'-0'.$day; 
		}
		$uv = $totalUvList[$idx];
		$suv= $singleUvList[$idx][$num][1];

		return round(($suv/$uv)*100,2).'%';
	}

	function doCreatExcel($month,$data,$column){
		global $objPHPExcel;
		global $oct_date;
		global $date;
		global $imgList;
		
  		$objActSheet = $objPHPExcel->getActiveSheet();
		$objActSheet->getColumnDimension($column)->setWidth(50);

		foreach ($data as $key => $value) {
			$objPHPExcel->setActiveSheetIndex(0)
			 		->setCellValue($column.'1', $key);

			$ii = 2;
			for ($i=0; $i < count($value); $i++) { 
				$pos =  $oct_date[$value[$i]];
				$day = $date[substr($pos, -1)];
				$num = substr($pos, 0, 1); //第几张banner

				$click_rate = getRate($month,$day,$num); //总的uv

				$objPHPExcel->setActiveSheetIndex(0)
			 		->setCellValue($column.$ii, $click_rate.' (10.'.$day.')');

			 	$style_color = array(
			        'font' => array(
			        	'bold' => true,
			            'color' => array('rgb'=>'FC031C')
			        )
			    );
		        $objActSheet->getStyle($column.$ii)->applyFromArray($style_color); 

		     	if(isset($imgList['10'.$day][$num])){
		     		//如果图片文件错误，则不导出图片
					$path = __DIR__.DIRECTORY_SEPARATOR ."images".DIRECTORY_SEPARATOR.'10'.$day.DIRECTORY_SEPARATOR.$imgList['10'.$day][$num];

					$objDrawing = new PHPExcel_Worksheet_Drawing();
		        	$objDrawing->setPath($path);
			
					$objDrawing->setCoordinates($column.($ii+1));
			        $objDrawing->setHeight(100);
			        $objDrawing->setWidth(335);
			        $objDrawing->setWorksheet($objPHPExcel->getActiveSheet());   
				}	

		        $ii += 10;     
			}
		}
		
	}


?>