<!DOCTYPE html>
<html lang="en" >
<head>
  <meta charset="UTF-8">
  <title>MKPort</title>
  <link rel='stylesheet' href='https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.0.0-alpha.6/css/bootstrap.min.css'>
<link rel='stylesheet' href='https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css'>
<link rel='stylesheet' href='https://fonts.googleapis.com/css?family=Montserrat:300,400,700"rel="stylesheet'><link rel="stylesheet" href="./style.css">
</head>

<body>
    <div class="popup2">
    <img src="./images/pop_tit.png">

<?php
    require_once "../PHPExcel-1.8/Classes/PHPExcel.php"; // PHPExcel.php을 불러옴.
	
    $objPHPExcel = new PHPExcel();

    $filename = '../202311_new.xlsx'; // 읽어들일 엑셀 파일의 경로와 파일명을 지정한다.
	
    require_once "../PHPExcel-1.8/Classes/PHPExcel/IOFactory.php"; // IOFactory.php을 불러옴.
    
    $company_id = $_POST['id'];
    $company_id = trim($company_id, "-");
    $company_id = str_replace("-", "", $company_id);

    if ( empty($company_id) ) {
      echo "<br>검색할 내용이 없습니다.";

      return;

    }

    try {		// 업로드 된 엑셀 형식에 맞는 Reader객체를 만든다.		
    
        $objReader = PHPExcel_IOFactory::createReaderForFile($filename);	

        // 읽기전용으로 설정		
        $objReader->setReadDataOnly(true);		// 엑셀파일을 읽는다	
        $objExcel = $objReader->load($filename);		// 첫번째 시트를 선택	
        $objExcel->setActiveSheetIndex(0);	
        $objWorksheet = $objExcel->getActiveSheet();	
        $rowIterator = $objWorksheet->getRowIterator();		
    
        foreach ($rowIterator as $row) { // 모든 행에 대해서		
            $cellIterator = $row->getCellIterator();		
            $cellIterator->setIterateOnlyExistingCells(false); 	
        }

        $maxRow = $objWorksheet->getHighestRow();

        for ($i = 1; $i <= $maxRow; $i++) {		
            $dataA = $objWorksheet->getCell('A' . $i)->getValue(); // A열		
	    $dataB = $objWorksheet->getCell('B' . $i)->getValue(); // B열		
            $dataC = $objWorksheet->getCell('C' . $i)->getValue(); // C열		
            $dataD = $objWorksheet->getCell('D' . $i)->getValue(); // D열		
            $dataE = $objWorksheet->getCell('E' . $i)->getValue(); // E열		
            $dataF = $objWorksheet->getCell('F' . $i)->getValue(); // F열		

	    if ($dataB==$company_id) {
		$dataB = substr($dataB, 0, 3)."-".substr($dataB, 3, 2)."-*****";
		$addr = explode(" ", $dataC);
		$addr2= $addr[0]." ".$addr[1]." "."***";

		echo "<table class='table_normal'>";
		echo "<tr><td>한글업체명</td><td class='list'>$dataA</td></tr>";
		echo "<tr><td>사업자번호</td><td class='list'>$dataB</td></tr>";
		//echo "<tr><td>사업장 주소</td><td class='list'>$dataC</td></tr>";
		echo "<tr><td>사업장 주소</td><td class='list'>$addr2</td></tr>";
		echo "<tr><td>품목</td><td class='list'>$dataD</td></tr>";
		echo "<tr><td>copy</td><td class='list'>$dataE</td></tr>";

		if ( gettype($dataF) == 'double' ) {
		  $st = intval((($dataF - 25569) * 86400));
		  $st2 = date("Y-m-d", $st);

		  echo "<tr><td>계약일자</td><td class='list'>$st2</td></tr>";

		}
		else {
		  echo "<tr><td>계약일자</td><td class='list'>$dataF</td></tr>";

  		}

		echo "</table>";

                return;

	    }

        }
       
        echo "한컴 제품 구매이력이 없는 사업장 입니다.";

    } 
    catch (exception $e) {	
        echo '엑셀파일을 읽는도중 오류가 발생하였습니다.<br/>';	
    }	
?>
    </div>
</body>
</html>
