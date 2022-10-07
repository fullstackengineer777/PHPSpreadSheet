<?php

	require_once(__DIR__ . '/vendor/autoload.php');
	use PhpOffice\PhpSpreadsheet\Spreadsheet; 
	use PhpOffice\PhpSpreadsheet\Writer\Xlsx; 

	$filename = "statistics/".date('M Y').".xlsx";
	if(file_exists($filename)){
		echo "the file exists.";
		$cur_date = date("d");

		// $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
		$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($filename);
		$cur_sheet = $spreadsheet->getSheetByName((string)$cur_date);
		$row = count($cur_sheet->toArray());

		$cur_sheet->setCellValueByColumnAndRow(1,$row+1,"ddd");
		echo count($row);
		//write it again to Filesystem with the same name (=replace)
		$writer = new Xlsx($spreadsheet);
		$writer->save($filename);

	}else{
		// If the file doesn't exist, a brand new file for current month is created
		$mySpreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();
		// delete the default active sheet
		$mySpreadsheet->removeSheetByIndex(0);
		//define header data
		$column_header=["01-contentstart","04-clientsnow","06-clientsafter","11-brandname","12-contentsell",
			"13-industry","15-sell-product-description","15-sell-product-title","16-cost-low-high","16-cost-low-high" , "17-company-sell","18_strengths","19-way","20-regions-txt","21_gender",
			"22_age","24-target","25-interests","31-email","32-payment"];
		// Create the sheets for everyday
		for($i = 0 ; $i < 31 ; $i++){		
			if($i < 9)
				$sheet_temp = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($mySpreadsheet, "0".(string)($i+1));
			else
				$sheet_temp = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($mySpreadsheet, (string)($i+1));
			
			$mySpreadsheet->addSheet($sheet_temp, $i);
			//add header titles for every sheet
			$j=1;
			foreach($column_header as $x_value) {
				$sheet_temp->setCellValueByColumnAndRow($j,1,$x_value);
	  			$j=$j+1;  		
			}
			foreach ($sheet_temp->getColumnIterator() as $column)
		    {
		        $sheet_temp->getColumnDimension($column->getColumnIndex())->setAutoSize(true);
		    }

		}
		

	// 01-contentstart Sales   
	// 02-specialty    month   
	// 04-clientsnow   50-100  
	// 06-clientsafter 25-50   
	// 11-brandname    fedex   
	// 12-contentsell  products    
	// 13-industry Design&Renovation,Sports&Fitness    
	// 15-sell-product-description receiver    
	// 15-sell-product-title   computer    
	// 16-cost-low-high    200 
	// 17-company-sell discounted-purcase,gift-with-purchase,install-payment   
	// 18_strengths    In-house-production,Fast-shipping   
	// 19-way  products,services   
	// 20-regions-txt  London  
	// 21_gender   Man,Woman   
	// 22_age  25-35,18-25 
	// 24-target   OK  
	// 25-interests    Sports-Fitness,Technology   
	// 31-email    coindevmentor9211@gmail.com

		// Change the widths of the columns to be appropriately large for the content in them.
		// https://stackoverflow.com/questions/62203260/php-spreadsheet-cant-find-the-function-to-auto-size-column-width
		// $worksheets = [$worksheet1, $worksheet2];

		// foreach ($worksheets as $worksheet)
		// {
		//     foreach ($worksheet->getColumnIterator() as $column)
		//     {
		//         $worksheet->getColumnDimension($column->getColumnIndex())->setAutoSize(true);
		//     }
		// }

		// Save to file.
		$writer = new PhpOffice\PhpSpreadsheet\Writer\Xlsx($mySpreadsheet);
		$writer->save($filename);

	}

?>