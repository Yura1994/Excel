<?php
require_once ('Classes/PHPExcel.php');
require_once ('Classes/PHPExcel/Writer/Excel5.php');

$xls = new PHPExcel();
$xls->setActiveSheetIndex(0);
$sheet = $xls->getActiveSheet();
$sheet->setTitle('Мой  test 2 gg');




$sheet->mergeCells("A1:A4");
$sheet->mergeCells("B1:D4");
$sheet->mergeCells("E1:E4");

//Ширина столбца (автомвтически)
$sheet->getColumnDimension("A") ->setAutoSize(true);
$sheet->getColumnDimension("B") ->setAutoSize(true);
$sheet->getColumnDimension("C") ->setAutoSize(true);
$sheet->getColumnDimension("D") ->setAutoSize(true);
$sheet->getColumnDimension("E") ->setAutoSize(true);


//картинка
$objDrawing = new PHPExcel_Worksheet_Drawing();
$objDrawing->setResizeProportional(false);
$objDrawing ->setName('gg');
$objDrawing ->setDescription('lkw');
$objDrawing  -> setPath(__DIR__.'/img/mm.jpg');
$objDrawing ->setCoordinates('B1');
$objDrawing ->setOffsetX(10);
$objDrawing ->setOffsetY(10);
$objDrawing ->setWidth(163);
$objDrawing ->setHeight(50);
$objDrawing ->setWorksheet($sheet);


//стили для ячейки с bg

$style = array(
    'font' =>array(
        'name' => 'Times New Roman',
        'size' => 16,
        'color' =>array('rgb' => 'FFFFFF'),
        'bolder' => true,
    )
);


//отчет о продаже
$sheet->mergeCells("B5:D5");
$bg1 = array(
    'fill' => array(
        'type' => PHPExcel_Style_Fill::FILL_SOLID,
        'color' => array('rgb' => '1E90FF')
    )
);
$sheet->getStyle('B5')->applyFromArray($bg1);
$sheet->getStyle("B5")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); // выравнивание в ячейки по центру
//$sheet->getStyle('B5')->applyFromArray($bg1);
$sheet->getStyle('B5')->applyFromArray($style);
$sheet->setCellValue("B5","Отчет по продажам");

//табличка
$border = array(
    'borders'=>array(
        'allborders'=>array(
            'style' => PHPExcel_Style_Border::BORDER_THIN,
            'color' => array('rgb' =>'000000')
        )
    )
);
$sheet->getStyle("A6:E10")->applyFromArray($border);

//№
$sheet->setCellValue("A6","№ п.п");
$bg2 = array(
    'fill' => array(
        'type' => PHPExcel_Style_Fill::FILL_SOLID,
        'color' => array('rgb' => '008000')
    )
);
$sheet->getStyle('A6')->applyFromArray($style);
$sheet->getStyle('A6')->applyFromArray($bg2);

// Название
$sheet->setCellValue("B6","Название");
$bg3 = array(
    'fill' => array(
        'type' => PHPExcel_Style_Fill::FILL_SOLID,
        'color' => array('rgb' => 'FFC0CB')
    )
);
$sheet->getStyle('B6')->applyFromArray($style);
$sheet->getStyle('B6')->applyFromArray($bg3);

// Цена
$sheet->setCellValue("C6","Цена");
$bg4 = array(
    'fill' => array(
        'type' => PHPExcel_Style_Fill::FILL_SOLID,
        'color' => array('rgb' => '999999')
    )
);
$sheet->getStyle('C6')->applyFromArray($style);
$sheet->getStyle('C6')->applyFromArray($bg4);

// Количество
$sheet->setCellValue("D6","Количество");
$bg5 = array(
    'fill' => array(
        'type' => PHPExcel_Style_Fill::FILL_SOLID,
        'color' => array('rgb' => 'D9614C')
    )
);
$sheet->getStyle('D6')->applyFromArray($style);
$sheet->getStyle('D6')->applyFromArray($bg5);

// Сумма
$sheet->setCellValue("E6","Сумма");
$bg6 = array(
    'fill' => array(
        'type' => PHPExcel_Style_Fill::FILL_SOLID,
        'color' => array('rgb' => '7DBE52')
    )
);
$sheet->getStyle('E6')->applyFromArray($style);
$sheet->getStyle('E6')->applyFromArray($bg6);



//заполнения таблички
$nam = ["name" =>["Масло", "Хлеб", "Сыр", "Молоко"],"price" =>[50,20,100,30], "quantity"=>[3,4,8,12]];

for ($i = 0; $i<count($nam["name"]); $i++)
{
    $col = $i+1;
    $yiach = $i+7;
    $ms = $i;
    $sheet->setCellValue("A{$yiach}","{$col}");
    $sheet->setCellValue("B{$yiach}","{$nam['name'][$ms]}");
    $sheet->setCellValue("C{$yiach}","{$nam['price'][$ms]}");
    $sheet->setCellValue("D{$yiach}","{$nam['quantity'][$ms]}");
    $sheet->setCellValue("E{$yiach}","=C{$yiach}*D{$yiach}");

}

// Конечный итог
$sheet->mergeCells("B12:D12");
$sheet->setCellValue("B12","Итог дооход: ");
$sheet->getStyle("B12")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$sty = array(
    'font' =>array(
        'name' => 'Times New Roman',
        'size' => 16,
        'color' =>array('rgb' => '00000'),
        'bolder' => true,
    )
);
$sheet->getStyle('B12')->applyFromArray($sty);
$sheet->setCellValue("E12","=SUM(E7:E10)");
$sheet->getStyle('E12')->applyFromArray($sty);


// сохранения и запись
$objWriter = new PHPExcel_Writer_Excel5($xls);
$filename = 'Доход.xls';
if (file_exists($filename))
{
    unlink($filename);
}
echo "Документ {$filename} сформованно в поточночній теці";
$objWriter->save($filename);