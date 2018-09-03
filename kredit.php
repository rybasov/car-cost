<!DOCTYPE HTML>
<html>
<head>
    <meta charset="utf-8">
</head>

<?php

/**
 *
 * @param $arg
 * @param $price
 * @return bool
 */
function check_price($arg, &$price)
{
    $price = (int)$arg;

    if ($price < 0) {
        return false;
    };

    return true;
}

/**
 * @param $arg
 * @param array $array
 * @param $month
 * @return bool
 */
function check_month($arg, $array, &$month)
{
    $month = (int)$arg;

    if ($month < 0) {
        return false;
    };

    if (array_key_exists($month, $array)) {
        return true;
    } else {
        return false;
    }
}

/****************************************************************************/
$arr_term = [12 => 25, 24 => 50, 36 => 75, 48 => 100, 60 => 130];

$n = 0;

foreach ($argv as $arg) {
    $n = $n + 1;
}

switch ($n) {
    case 1:
        exit("Price of car not specified !" . PHP_EOL);
        break;

    case 2:
        if (!check_price($argv [1], $price_car)) {
            exit("Price of car has bad specified !" . PHP_EOL);
        }

        $prepayment = 0;
        $all_month = true;
        break;
    case 3:
        if (!check_price($argv [1], $price_car)) {
            exit("Price of car has bad specified !" . PHP_EOL);
        }

        if (!check_price($argv [2], $prepayment)) {
            exit("Prepayment has bad specified !" . PHP_EOL);
        }

        $all_month = true;
        break;
    case 4:
        if (!check_price($argv [1], $price_car)) {
            exit("Price of car has bad specified !" . PHP_EOL);
        }

        if (!check_price($argv [2], $prepayment)) {
            exit("Prepayment has bad specified !" . PHP_EOL);
        }

        if (!check_month($argv [3], $arr_term, $month_get)) {
            exit("Monthly payment has bad specified !" . PHP_EOL);
        }

        $all_month = false;
        break;
};
//------------------------------------------------------------------------------
include_once('PHPExcel/IOFactory.php');

$objPHPExcelOutput = new PHPExcel();

$objPHPExcelOutput->getProperties()->setCreator("PHP")
    ->setLastModifiedBy("Konstantin")
    ->setTitle("Office 2007 XLSX")
    ->setSubject("Office 2007 XLSX")
    ->setDescription("File Office 2007 XLSX, by PHPExcel.")
    ->setKeywords("office 2007 openxml php")
    ->setCategory("Test");

$objPHPExcelOutput->getActiveSheet()->setTitle('Kredit');
//==============================================================================
$row = 0;

foreach ($arr_term as $month => $percent) {
    if ($all_month or ($month == $month_get)) {
        ++$row;

        $row_begin = $row;

        $cell = "B{$row}";

        $objPHPExcelOutput->setActiveSheetIndex(0)
            ->setCellValue($cell, "{$month} мес");

        $objPHPExcelOutput->setActiveSheetIndex(0)
            ->getStyle($cell)
            ->getFill()->applyFromArray([
                'type'       => PHPExcel_Style_Fill::FILL_SOLID,
                'startcolor' => ['rgb' => 'FFE6E6'],
                'endcolor'   => ['rgb' => 'FFE6E6'],
            ]);

        $objPHPExcelOutput->setActiveSheetIndex(0)
            ->getStyle($cell)
            ->getFont()->applyFromArray(['bold' => true]);


        $objPHPExcelOutput->setActiveSheetIndex(0)
            ->getStyle($cell)
            ->getAlignment()
            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        //--------------------------------------------
        ++$row;

        $row_cost = $row;

        $cell = "A{$row}";

        $objPHPExcelOutput->setActiveSheetIndex(0)
            ->setCellValue($cell, "Соимость автомобиля");

        $cell = "B{$row}";

        $objPHPExcelOutput->setActiveSheetIndex(0)
            ->setCellValue($cell, $price_car);

        $objPHPExcelOutput->setActiveSheetIndex(0)
            ->getStyle($cell)
            ->getFill()->applyFromArray([
                'type'       => PHPExcel_Style_Fill::FILL_SOLID,
                'startcolor' => ['rgb' => 'FFFF00'],
                'endcolor'   => ['rgb' => 'FFFF00'],
            ]);
        //--------------------------------------------
        ++$row;

        $row_prepayment = $row;

        $cell = "A{$row}";

        $objPHPExcelOutput->setActiveSheetIndex(0)
            ->setCellValue($cell, "Аванс");

        $cell = "B{$row}";

        $objPHPExcelOutput->setActiveSheetIndex(0)
            ->setCellValue($cell, $prepayment);
        //--------------------------------------------
        ++$row;

        $row_real = $row;

        $cell = "A{$row}";

        $objPHPExcelOutput->setActiveSheetIndex(0)
            ->setCellValue($cell, "Стоимость с учетом аванса");

        $cell = "B{$row}";
        $value = "=B{$row_cost}-B{$row_prepayment}";

        $objPHPExcelOutput->setActiveSheetIndex(0)
            ->setCellValue($cell, $value);    //$price_car - $prepayment);
        //--------------------------------------------
        ++$row;

        $row_percent = $row;

        $percents = round(($price_car - $prepayment) * $arr_term[$month] / 100, 0);

        $cell = "A{$row}";

        $objPHPExcelOutput->setActiveSheetIndex(0)
            ->setCellValue($cell, "Проценты");

        $cell = "B{$row}";
        $value = "=ROUND(B{$row_real}*{$arr_term[$month]}/100,0)";

        $objPHPExcelOutput->setActiveSheetIndex(0)
            ->setCellValue($cell, $value);    // $percents);
        //--------------------------------------------
        ++$row;

        $row_total = $row;

        $total = $price_car - $prepayment + $percents;

        $cell = "A{$row}";

        $objPHPExcelOutput->setActiveSheetIndex(0)
            ->setCellValue($cell, "Всего");

        $cell = "B{$row}";
        $value = "=B{$row_cost}-B{$row_prepayment}+B{$row_percent}";

        $objPHPExcelOutput->setActiveSheetIndex(0)
            ->setCellValue($cell, $value);    // $total);
        //--------------------------------------------
        ++$row;

        $monthly_payment = round($total / $month, 2);

        $cell = "A{$row}";

        $objPHPExcelOutput->setActiveSheetIndex(0)
            ->setCellValue($cell, "Ежемесячный платеж");

        $cell = "B{$row}";
        $value = "=B{$row_total}/{$month}";

        $objPHPExcelOutput->setActiveSheetIndex(0)
            ->setCellValue($cell, $value);    // $monthly_payment);

        $objPHPExcelOutput->setActiveSheetIndex(0)
            ->getStyle($cell)
            ->getFill()->applyFromArray([
                'type'       => PHPExcel_Style_Fill::FILL_SOLID,
                'startcolor' => ['rgb' => '80FF80'],
                'endcolor'   => ['rgb' => '80FF80'],
            ]);
        //--------------------------------------------
        $cell = "A{$row_begin}:B{$row}";

        $objPHPExcelOutput->setActiveSheetIndex(0)
            ->getStyle($cell)
            ->getBorders()
            ->getAllBorders()
            ->applyFromArray(['style' => PHPExcel_Style_Border::BORDER_THIN]);
        //--------------------------------------------
        echo "мес - {$month}, percent - {$arr_term[$month]}, price car - {$price_car}, prepayment - {$prepayment}, percents - {$percents}, total - {$total}, monthly payment - {$monthly_payment}" . PHP_EOL;

        ++$row;
    }
}

$objPHPExcelOutput->setActiveSheetIndex(0)
    ->getColumnDimension("A")->setAutoSize(true);

$objPHPExcelOutput->setActiveSheetIndex(0)
    ->getColumnDimension("B")->setAutoSize(true);

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcelOutput, 'Excel2007');

$file = __DIR__ . DIRECTORY_SEPARATOR . "Kredit.xlsx";

$objWriter->save($file);
