<?php
	ini_set('memory_limit','500M');
	include './Classes/PHPExcel/IOFactory.php';

	header("Content-type: text/html; charset=utf-8");
	
	function outputUv($file,$type){
		$reader = PHPExcel_IOFactory::createReader('Excel5'); //设置以Excel5格式(Excel97-2003工作簿)
		$PHPExcel = $reader->load($file); // 载入excel文件
		$sheet = $PHPExcel->getSheet(0); // 读取第一個工作表

		$highestRow = $sheet->getHighestRow(); // 取得总行数
		//$highestColumm = $sheet->getHighestColumn(); // 取得总列数
		 
		/** 循环读取每个单元格的数据 */
		$rows = array();
		$dataset = array();
		
		for ($row = 2; $row <= $highestRow; $row++){//行数是以第1行开始
	    	$tmp = $sheet->getCell('C'.$row)->getValue();
	    	if($tmp == $type){
	        	array_push($rows, $row);
	    	}
		}
		
		for ($i = 0; $i < count($rows); $i++){
			$time = $sheet->getCell('A'.$rows[$i])->getValue();
			$uv   = $sheet->getCell('G'.$rows[$i])->getValue();
			$dataset[$time] = $uv;
		}
		file_put_contents('oct_uv_t.php', "<?php\nreturn " .var_export($dataset,true). ";");
		echo "done";
	}

	function outputSingleUv($file){
		$reader = PHPExcel_IOFactory::createReader('Excel5'); //设置以Excel5格式(Excel97-2003工作簿)
		$PHPExcel = $reader->load($file); // 载入excel文件
		$sheet = $PHPExcel->getSheet(0); // 读取第一個工作表

		$highestRow = $sheet->getHighestRow(); // 取得总行数

		$columns = array('G','L');

		$dataset = array();
		for ($row=2; $row <= $highestRow ; $row++) { 
			$time = $sheet->getCell('A'.$row)->getValue();
			$tmp  = array();
			for ($i=0; $i < count($columns); $i++) { 
				$tmp[] = $sheet->getCell( $columns[$i].$row )->getValue();
			}
			$dataset[$time][] = $tmp;		
		}
		file_put_contents('oct_uv_s.php', "<?php\nreturn " .var_export($dataset,true). ";");
		echo "done";
	}

	function outputImg(){
		header("Content-type: text/html; charset=gb2312");
		$dir = __DIR__.DIRECTORY_SEPARATOR ."images";
		$files = array();
		if (is_dir($dir)) {
			if ($dh = opendir($dir)) {
				while (($file = readdir($dh)) !== false) {
					if($file != '.' && $file != '..' ){
						$key = preg_replace('/([\x80-\xff]*)/i','',$file);//去掉中文
						rename($dir.DIRECTORY_SEPARATOR.$file, $dir.DIRECTORY_SEPARATOR.$key);
						$files[$key] = $key; 
					}	
				}	
			}	
		}

		$imgList = array();
		foreach ($files as $key => $value) {
			$path = $dir .DIRECTORY_SEPARATOR. $value;
			if($dh = opendir($path)){
				while(($file = readdir($dh)) !== false){
					if(substr($file, -3) == 'jpg'){
						preg_match('/banner(\d)/', $file,$matches);
						$idx = intval($matches[1]);
						$imgList[$key][$idx] = $file;
					}
				}
			}
		}
		//print_r($imgList);
		file_put_contents('oct_img.php', "<?php\nreturn " .var_export($imgList,true). ";");
		echo "done";
	}

	function outputSchedule($file){
		$reader = PHPExcel_IOFactory::createReader('Excel5'); //设置以Excel5格式(Excel97-2003工作簿)
		$PHPExcel = $reader->load($file); // 载入excel文件
		$sheet = $PHPExcel->getSheet(0); // 读取第一個工作表

		$highestRow = $sheet->getHighestRow(); // 取得总行数
		$highestColumm = $sheet->getHighestColumn(); //取得总列数

		$keys = array();
		$date = array();
		for ($row=1; $row <= $highestRow ; $row++) { 
			for ($i='A'; $i <= $highestColumm; $i++) {
				$tmp = preg_replace('/\s/', '', $sheet->getCell( $i.$row )->getValue());

				if(!empty($tmp)){
					array_push($keys,$tmp);
					array_push($date,$row.$i);
				}
			}
		}
		
		$tmp = array();
		$counts = array_count_values($keys);
		foreach($counts as $k => $value){
			if($value >= 2){
				$tmp[][$k] = array_keys($keys,$k);
			}
		}
		file_put_contents('oct_faxian.php', "<?php\nreturn " .var_export($tmp,true). ";");
		file_put_contents('oct_faxian_date.php', "<?php\nreturn " .var_export($date,true). ";");
		echo "done";
	}


	$config = array(
		'total_Uv'=>array(
				'file'  => 'oct_2016.xls'
				,'column'=> '微信-发现入口'
			)
		,'single_Uv'=>'RD-oct.xls'
		,'schedule'=>'Oct_schedule.xls'
	);


	outputUv( $config['total_Uv']['file'], $config['total_Uv']['column'] );  // 输出所有uv

	outputSingleUv( $config['single_Uv'] ); //输出单个uv
	
	outputSchedule( $config['schedule'] ); //输出排期
	
	outputImg();  //输出图片
?>