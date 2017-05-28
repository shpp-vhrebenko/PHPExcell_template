<?php

require_once("classes/PHPExcel.php");
require_once("classes/PHPExcel/IOFactory.php");

set_time_limit (300);
ini_set('error_reporting', E_ALL);
ini_set('display_errors', 1);
ini_set('display_startup_errors', 1);

date_default_timezone_set('europe/kiev');

//=====================================================================================================================
//														//							Header 2
//						Header 1						//=============================================================
//														//		//	block   //				//			block	    //
//														//		//		2	//				//				3		//
//=====================================================================================================================
//                      Header 3                        // 							Header 4
//======================================================//==============================================================
//	h1	//			h2		//		h3	//	h4	//	h5	// // // // // // //		block 1	// // // // // // // // //
//=====================================================================================================================
//				content 1								//							content 2							

header("Content-type: text/html; charset=utf-8;");

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
		if(isset($_POST["project"])) {
			$project = $_POST["project"]; 
	}
}

$phpexcel = new PHPExcel();// Создаём объект PHPExcel

$page = $phpexcel->setActiveSheetIndex(0);// Каждый раз делаем активной 1-ю страницу и получаем её, потом записываем в неё данные

$page->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);//Ориентация страницы и  размер листа

$phpexcel->getDefaultStyle()->getAlignment()->setWrapText(true); // перенос строки

// Стили таблицы ===========================================================
$style_wryte = array(
	'borders' => array(
		'allborders' => array(
			'style' => PHPExcel_Style_Border::BORDER_THIN,
			'color' => array(
				'rgb' => '000000'
			)
		)
	)
);

$style_fill_cell = array(
    'fill'  => array(
        'type' => PHPExcel_STYLE_FILL::FILL_SOLID,
        'color' => array(
            'rgb' => 'FEFE00'
        )
));

$arHeadStyleFont = array(
    'font'  => array(
        'italic' => false,
        'size'  => 15,
        'name'  => 'Verdana'
    ));
//=========================================================================

$countHeaders = 5;
$headerHight = 7;
$defaultCalendarHeaderLength = 60;
$totalCountDuration = $project['totalDuration'] > $defaultCalendarHeaderLength ? $project['totalDuration'] : $defaultCalendarHeaderLength;
$nameCellTotalCountDuration = numberToColumnName($totalCountDuration + $countHeaders);
// Размер ячеек заголовка ==============================================================================================
// h1-h5
$page->getColumnDimension('A')->setWidth(5);
$page->getColumnDimension('B')->setWidth(50);
$page->getColumnDimension('C')->setWidth(10);
$page->getColumnDimension('D')->setWidth(10);
$page->getColumnDimension('E')->setWidth(10);
//$phpexcel->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);

//==============================================================================================

//Header 1==============================================================================================
$page->mergeCells("A1:E5");
$page->setCellValue('A1', "Графік виконання робіт");
$page->getStyle('A1:E1')->applyFromArray($arHeadStyleFont);
$page->getStyle("A1:E5")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

// Header 2==============================================================================================
$nextCell = numberToColumnName($defaultCalendarHeaderLength + $countHeaders);
$page->mergeCells("F1:$nextCell"."1");
$page->setCellValue('F1', $project['name']);
$page->getStyle("F1:$nextCell"."1")->applyFromArray($arHeadStyleFont);
$page->getStyle("F1:$nextCell"."1")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT)->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

// block 1==============================================================================================
$nn = 1;
$startCol = 5;
$startIter = $startCol;
for ($j=0; $j < $totalCountDuration; $j++) { 
	$startIter++;
	$nameCell = numberToColumnName($startIter);
	$page->getColumnDimension($nameCell)->setWidth(4);
	$page->setCellValue($nameCell."7",$nn++);
}
$nameCellStart = numberToColumnName($startCol + 1);
$page->getStyle($nameCellStart."7:".$nameCellTotalCountDuration."7")->applyFromArray($style_wryte);
$page->getStyle($nameCellStart."7:".$nameCellTotalCountDuration."7")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT)->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$page->getStyle($nameCellStart."8:".$nameCellTotalCountDuration.(count($project['tasks']) + $headerHight))->applyFromArray($style_wryte);
// block 2==============================================================================================
$page->mergeCells("F2:P2");
$page->setCellValue('F2', "Затверджую");
$page->getStyle("F2:P2")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

$page->mergeCells("F3:P3");
$page->setCellValue('F3', "_____________________");
$page->getStyle("F3:P3")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$year_now = date("Y"); 
$page->mergeCells("F4:P4");
$page->setCellValue('F4', "'___' ____________ $year_now рік");
$page->getStyle("F4:P4")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

// block 3============================================================================================== 
$page->mergeCells("Q2:".$nameCellTotalCountDuration."2");
$page->setCellValue('Q2', "Примітка");
$page->getStyle("Q2:".$nameCellTotalCountDuration."2")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT)->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

$page->mergeCells("Q3:".$nameCellTotalCountDuration."3");
$page->setCellValue('Q3', "1. Графік платежів додається та є невід`ємною частиною грфіку виконання робіт");
$page->getStyle("Q3:".$nameCellTotalCountDuration."3")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT)->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

$page->mergeCells("Q4:".$nameCellTotalCountDuration."4");
$page->setCellValue('Q4', "2. Графік може коригуватись при зміні проектих рішень, затримки оплати та форс-мажорних обставин");
$page->getStyle("Q4:".$nameCellTotalCountDuration."4")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT)->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

// Header 3
$page->mergeCells("A6:E6");
$page->setCellValue('A6', "Фактична дата початку проведення робіт");
$page->getStyle('A6:E6')->applyFromArray($style_wryte);
$page->getStyle("A6:E6")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
// Header 4
$page->mergeCells("F6:$nameCellTotalCountDuration"."6");
$page->setCellValue('F6', "Робочі дні з урахування вихідних і святкових днів");
$page->getStyle("F6:$nameCellTotalCountDuration"."6")->applyFromArray($style_wryte);
$page->getStyle("F6:$nameCellTotalCountDuration"."6")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT)->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

//h1-h5
$page->setCellValue('A7', "№");
$page->setCellValue('B7', "Розділ");
$page->setCellValue('C7', "Вартість грн.");
$page->setCellValue('D7', "Прод.");
$page->setCellValue('E7', "Кількість працівників");
$page->getStyle("A7:E7")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$page->getStyle("A7:E7")->applyFromArray($style_wryte);

// content 1
$i = 8;
$nn = 1;
$startCell = 6;
foreach( $project['tasks'] as $row ) {
   		$page->setCellValue("A$i",$nn++);
        foreach( $row as $key=>$value ) {
       	
            switch ($key) {
                case 'name':
                    $page->setCellValue("B$i",$value);
                break;

                case 'price':
                    $page->setCellValue("C$i",$value);
                break;

                case 'duration':
                    $page->setCellValue("D$i",$value);
                    // content 2===================================================================================
                    if($value == 1) {
                    	$nameStartCell = numberToColumnName($startCell);
                    	$page->getStyle($nameStartCell."$i:$nameStartCell".$i)->applyFromArray($style_fill_cell);
                    	$startCell++;
                    } else if ($value > 1) {
                    	$nextCell = $startCell+$value;
                    	$nameStartCell = numberToColumnName($startCell);
                    	$nameNextCell = numberToColumnName($nextCell);
                    	$page->getStyle($nameStartCell."$i:$nameNextCell".$i)->applyFromArray($style_fill_cell);
                    	$startCell = $nextCell;
                    } else {

                    }
                    //==============================================================
                break;

                case 'mancount':
                    $page->setCellValue("E$i",$value);
                break; 
            }

    
        }
            $page->getStyle("A$i:E$i")->applyFromArray($style_wryte);
            $page->getStyle("A$i:E$i")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT)->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $i++;
}


$objWriter = PHPExcel_IOFactory::createWriter($phpexcel, 'Excel2007');

$objWriter->save("schedule_works.xlsx");

exit('<meta http-equiv="refresh" content="0; URL=schedule_works.xlsx">');

function numberToColumnName($number){
    $abc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    $abc_len = strlen($abc);

    $result_len = 1; // how much characters the column's name will have
    $pow = 0;
    while( ( $pow += pow($abc_len, $result_len) ) < $number ){
        $result_len++;
    }

    $result = "";
    $next = false;
    // add each character to the result...
    for($i = 1; $i<=$result_len; $i++){
        $index = ($number % $abc_len) - 1; // calculate the module

        // sometimes the index should be decreased by 1
        if( $next || $next = false ){
            $index--;
        }

        // this is the point that will be calculated in the next iteration
        $number = floor($number / strlen($abc));

        // if the index is negative, convert it to positive
        if( $next = ($index < 0) ) {
            $index = $abc_len + $index;
        }

        $result = $abc[$index].$result; // concatenate the letter
    }
    return $result;
}
