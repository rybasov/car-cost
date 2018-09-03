<?php  
set_time_limit(NULL);

//------------------------------------------------------------------------------
function get_nom_column ( $nK )
{
    if ( $nK <= 26 )
    {
       return chr ( ord ( 'A' ) + $nK - 1 );
    }

    if ( $nK <= 26 + 26 )
    {
       return 'A'.chr ( ord ( 'A' ) + $nK - 1 - 26 );
    }
}
//==============================================================================
function output_excel_my ( $nN, $nK, $nIDTovar, $nKursEur, $aProdSotra, $aProdPE, $aProdTH, $objPHPExcelOutput, &$nPriceSrav )
{
    $nPriceSrav = 0;
    
    if ( array_key_exists ( $nIDTovar, $aProdSotra ) )
    {
        $cK = get_nom_column ( $nK );
        
        $cCell = $cK.$nN;
        
        $nPrice = $aProdSotra[$nIDTovar]['price'];
        
        if ( $nPrice > 0 )
        {
            $nPriceSrav = $nPrice * $nKursEur;
            
            $nPriceSrav = (int)$nPriceSrav;
            
            if ( $nPriceSrav > 0 )
            {
                $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                                  ->setCellValue ( $cCell, $nPriceSrav );
                
                if ( ! ( $aProdSotra[$nIDTovar]['promo'] == 0 ) )
                {
                    $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                                      ->getStyle($cCell)
                                      ->getFill()->applyFromArray ( array ( 'type'=>PHPExcel_Style_Fill::FILL_SOLID,
                                                                            'startcolor'=>array('rgb'=>'FFDB70'),
                                                                            'endcolor'  =>array('rgb'=>'FFDB70') ) );
                }
                
                if ( $aProdSotra[$nIDTovar]['kolvo'] === "0" )
                {
                    $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                                      ->getStyle($cCell)
                                      ->getFont()->applyFromArray ( array ( 'bold'=>true,
                                                                            'size'=>9,
                                                                            'color'=>array ( 'rgb'=>'CCCCCC' ) ) );
                }
            }
        }
    }
//------------------------------------------------------------------------------
    if ( array_key_exists ( $nIDTovar, $aProdTH ) )
    {
        $cK = get_nom_column ( $nK + 2 );
        
        $cCell = $cK.$nN;
        
        $nPrice = $aProdTH[$nIDTovar]['curprice'];
        
        if ( $nPrice > 0 )
        {
            if ( $nPriceSrav <= 0 )
            {
                $nPriceSrav = $nPrice;
            }
            
            $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                              ->setCellValue ( $cCell, $nPrice );
            
            if ( $aProdTH[$nIDTovar]['promo'] === "1" )
            {
                $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                                  ->getStyle($cCell)
                                  ->getFill()->applyFromArray ( array ( 'type'=>PHPExcel_Style_Fill::FILL_SOLID,
                                                                        'startcolor'=>array('rgb'=>'FFDB70'),
                                                                        'endcolor'  =>array('rgb'=>'FFDB70') ) );
            }
            
            if ( $aProdTH[$nIDTovar]['avail'] === "0" )
            {
                $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                                  ->getStyle($cCell)
                                  ->getFont()->applyFromArray ( array ( 'bold'=>true,
                                                                        'size'=>9,
                                                                        'color'=>array ( 'rgb'=>'CCCCCC' ) ) );
            }
            else 
            {
                if ( $nPriceSrav > 0 )
                {
                    if ( $nPrice > $nPriceSrav * 1.1 )
                    {
                        $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                                          ->getStyle($cCell)
                                          ->getFont()->applyFromArray ( array ( 'bold'=>true,
                                                                                'color'=>array ( 'rgb'=>'51FF3D' ) ) );
                    }
                    
                    if ( $nPrice < $nPriceSrav * 0.9 )
                    {
                        $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                                          ->getStyle($cCell)
                                          ->getFont()->applyFromArray ( array ( 'bold'=>true,
                                                                                'color'=>array ( 'rgb'=>'FF0000' ) ) );
                    }
                }
            }
        }
    }
//------------------------------------------------------------------------------
    if ( array_key_exists ( $nIDTovar, $aProdPE ) )
    {
        $cK = get_nom_column ( $nK + 1 );
        
        $cCell = $cK.$nN;
        
        $nPrice = $aProdPE[$nIDTovar]['curprice'];
        
        if ( $nPrice > 0 )
        {
            $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                              ->setCellValue ( $cCell, $nPrice );
            
            if ( $aProdPE[$nIDTovar]['promo'] === "1" )
            {
                $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                                  ->getStyle($cCell)
                                  ->getFill()->applyFromArray ( array ( 'type'=>PHPExcel_Style_Fill::FILL_SOLID,
                                                                        'startcolor'=>array('rgb'=>'FFDB70'),
                                                                        'endcolor'  =>array('rgb'=>'FFDB70') ) );
            }
            
            if ( $aProdPE[$nIDTovar]['avail'] === "0" )
            {
                $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                                  ->getStyle($cCell)
                                  ->getFont()->applyFromArray ( array ( 'bold'=>true,
                                                                        'size'=>9,
                                                                        'color'=>array ( 'rgb'=>'CCCCCC' ) ) );
            }
            else 
            {
                if ( $nPriceSrav > 0 )
                {
                    if ( $nPrice > $nPriceSrav * 1.1 )
                    {
                        $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                                          ->getStyle($cCell)
                                          ->getFont()->applyFromArray ( array ( 'bold'=>true,
                                                                                'color'=>array ( 'rgb'=>'51FF3D' ) ) );
                    }
                    
                    if ( $nPrice < $nPriceSrav * 0.9 )
                    {
                        $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                                          ->getStyle($cCell)
                                          ->getFont()->applyFromArray ( array ( 'bold'=>true,
                                                                                'color'=>array ( 'rgb'=>'FF0000' ) ) );
                    }
                }
            }
        }
    }
}
//==============================================================================                     
function output_excel_PE ( $nN, $nK, $nIDTovar, $nKursEur, $aProdSotra, $aProdPE, $aProdTH, $objPHPExcelProEco, &$nPriceSrav )
{
    $nPriceSrav = 0;
    
    if ( array_key_exists ( $nIDTovar, $aProdSotra ) )
    {
        $cK = get_nom_column ( $nK );
        
        $cCell = $cK.$nN;
        
        $nPrice = $aProdSotra[$nIDTovar]['price'];
        
        if ( $nPrice > 0 )
        {
            $nPriceSrav = $nPrice * $nKursEur;
            
            $nPriceSrav = (int)$nPriceSrav;
            
            if ( $nPriceSrav > 0 )
            {
                $objPHPExcelProEco->setActiveSheetIndex ( 0 )
                                  ->setCellValue ( $cCell, $nPriceSrav );
            }
        }
    }
//------------------------------------------------------------------------------
    if ( array_key_exists ( $nIDTovar, $aProdTH ) )
    {
        $cK = get_nom_column ( $nK + 1 );
        
        $cCell = $cK.$nN;
        
        $nPrice = $aProdTH[$nIDTovar]['curprice'];
        
        if ( $nPrice > 0 )
        {
            if ( $nPriceSrav <= 0 )
            {
                $nPriceSrav = $nPrice;
            }
            
            $objPHPExcelProEco->setActiveSheetIndex ( 0 )
                              ->setCellValue ( $cCell, $nPrice );
        }
    }
//------------------------------------------------------------------------------
//    if ( array_key_exists ( $nIDTovar, $aProdPE ) )
//    {
//        $cK = get_nom_column ( $nK + 1 );
//        
//        $cCell = $cK.$nN;
//        
//        $nPrice = $aProdPE[$nIDTovar]['curprice'];
//        
//        if ( $nPrice > 0 )
//        {
//            $objPHPExcelOutput->setActiveSheetIndex ( 0 )
//                              ->setCellValue ( $cCell, $nPrice );
//            
//            if ( $aProdPE[$nIDTovar]['promo'] === "1" )
//            {
//                $objPHPExcelOutput->setActiveSheetIndex ( 0 )
//                                  ->getStyle($cCell)
//                                  ->getFill()->applyFromArray ( array ( 'type'=>PHPExcel_Style_Fill::FILL_SOLID,
//                                                                        'startcolor'=>array('rgb'=>'FFDB70'),
//                                                                        'endcolor'  =>array('rgb'=>'FFDB70') ) );
//            }
//            
//            if ( $aProdPE[$nIDTovar]['avail'] === "0" )
//            {
//                $objPHPExcelOutput->setActiveSheetIndex ( 0 )
//                                  ->getStyle($cCell)
//                                  ->getFont()->applyFromArray ( array ( 'bold'=>true,
//                                                                        'size'=>9,
//                                                                        'color'=>array ( 'rgb'=>'CCCCCC' ) ) );
//            }
//            else 
//            {
//                if ( $nPriceSrav > 0 )
//                {
//                    if ( $nPrice > $nPriceSrav * 1.1 )
//                    {
//                        $objPHPExcelOutput->setActiveSheetIndex ( 0 )
//                                          ->getStyle($cCell)
//                                          ->getFont()->applyFromArray ( array ( 'bold'=>true,
//                                                                                'color'=>array ( 'rgb'=>'51FF3D' ) ) );
//                    }
//                    
//                    if ( $nPrice < $nPriceSrav * 0.9 )
//                    {
//                        $objPHPExcelOutput->setActiveSheetIndex ( 0 )
//                                          ->getStyle($cCell)
//                                          ->getFont()->applyFromArray ( array ( 'bold'=>true,
//                                                                                'color'=>array ( 'rgb'=>'FF0000' ) ) );
//                    }
//                }
//            }
//        }
//    }
}
//==============================================================================
function output_excel_cell ( $nN, $nK, $nPrice, $nKurs, $nPriceOld, $nAvail, $l, $cWWW, $objPHPExcelOutput, $objPHPExcelProEco, $nPriceSrav )
{           
    $cK = get_nom_column ( $nK + 3 );
    
    $cCell = $cK.$nN;
    
    if ( $l )
    {
        $nPriceGrn = $nPrice * $nKurs;
        
        $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                          ->setCellValue ( $cCell, (int)$nPriceGrn );
        
        if ( $nPriceOld != 'NULL' )
        {
            $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                              ->getStyle($cCell)
                              ->getFill()->applyFromArray(array('type'=>PHPExcel_Style_Fill::FILL_SOLID,
                                                                'startcolor'=>array('rgb'=>'CCFF66'),
                                                                'endcolor'  =>array('rgb'=>'CCFF66')
                                                                )
                                                          );
        }
        
        if ( $nPriceSrav > 0 )
        {
            if ( ( $nPriceGrn > $nPriceSrav * 1.02 ) and ( $nAvail > 0 ) )
            {
                $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                                  ->getStyle($cCell)
                                  ->getFont()->applyFromArray ( array ( 'bold'=>true,
                                                                        'color'=>array ( 'rgb'=>'00B050' ) ) );
            }
            
            if ( ( $nPriceGrn > $nPriceSrav * 1.02 ) and ( $nAvail <= 0 ) )
            {
                $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                                  ->getStyle($cCell)
                                  ->getFont()->applyFromArray ( array ( 'bold'=>true,
                                                                        'size'=>8,
                                                                        'color'=>array ( 'rgb'=>'66FF66' ) ) );
            }
            
            if ( ( $nPriceGrn < $nPriceSrav * 0.95 ) and ( $nAvail > 0 ) )
            {
                $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                                  ->getStyle($cCell)
                                  ->getFont()->applyFromArray ( array ( 'bold'=>true,
                                                                        'color'=>array ( 'rgb'=>'FF0000' ) ) );
            }
            
            if ( ( $nPriceGrn < $nPriceSrav * 0.95 ) and ( $nAvail <= 0 ) )
            {
                $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                                  ->getStyle($cCell)
                                  ->getFont()->applyFromArray ( array ( 'bold'=>true,
                                                                        'size'=>8,
                                                                        'color'=>array ( 'rgb'=>'FF8080' ) ) );
            }
        }
        else 
        {
            if ( $nAvail === 0 )
            {
                $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                                  ->getStyle($cCell)
                                  ->getFont()->applyFromArray(array('bold'=>true,
                                                                    'size'=>8 ));
            }
            
            if ( $nAvail === -1 )
            {
                $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                                  ->getStyle($cCell)
                                  ->getFont()->applyFromArray(array('bold'=>true,
                                                                    'italic'=>true,
                                                                    'size'=>8 ));
            }
        }
    }
    else
    {
        if ( strlen ( $cWWW ) !== 0 )
        {
            $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                              ->getStyle($cCell)
                              ->getFill()->applyFromArray(array('type'=>PHPExcel_Style_Fill::FILL_SOLID,
                                                                'startcolor'=>array('rgb'=>'DDDDDD'),
                                                                'endcolor'  =>array('rgb'=>'DDDDDD')
                                                                )
                                                          );
        }
    }
    //======================================================================
    //======================================================================
    if ( $nK === 4 )
    {
        $cK = get_nom_column ( $nK - 1 );
        
        $cCell = $cK.$nN;
        
        $nPriceSotra = $objPHPExcelProEco->getActiveSheet ()
                                         ->getCell  ( $cCell )
                                         ->getValue ();
        
        if ( is_null ( $nPriceSotra ) or empty ( $nPriceSotra ) )
        {
            $nPriceSotra = 0;
        }
        // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        $cK = get_nom_column ( $nK );
        
        $cCell = $cK.$nN;
        
        $nPriceTH = $objPHPExcelProEco->getActiveSheet ()
                                      ->getCell  ( $cCell )
                                      ->getValue ();
        
        if ( is_null ( $nPriceTH ) or empty ( $nPriceTH ) )
        {
            $nPriceTH = 0;
        }
        // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        $nPrice = max ( $nPriceSotra, $nPriceTH );
        // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        if ( $l )
        {
            $nPriceMult = (int)$nPriceGrn;
            
            $cK = get_nom_column ( $nK + 1 );
            
            $cCell = $cK.$nN;
            
            $objPHPExcelProEco->setActiveSheetIndex ( 0 )
                              ->setCellValue ( $cCell, $nPriceMult );
        }
        else
        {
            $nPriceMult = 0;
        }
        // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        $nPricePE = $nPrice;
        
        $lMult = FALSE;
        
        if ( $nPriceMult > 0 )
        {
            if ( $nPriceMult <= $nPrice * 0.95 )
            {
                $nPricePE = $nPrice * 0.95;
                // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                $cK = get_nom_column ( $nK + 1 );
                
                $cCell = $cK.$nN;
                
                $objPHPExcelProEco->setActiveSheetIndex ( 0 )
                                  ->getStyle($cCell)
                                  ->getFont()->applyFromArray ( array ( 'bold'=>true,
                                                                        'color'=>array ( 'rgb'=>'FF0000' ) ) );
            }
            else
            {
                if ( $nPriceMult < $nPrice )
                {
                    $nPricePE = $nPriceMult;
                    $lMult    = TRUE;
                }
            }
        }
        // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        $cK = get_nom_column ( $nK + 2 );
        
        $cCell = $cK.$nN;
        
        $objPHPExcelProEco->setActiveSheetIndex ( 0 )
                          ->setCellValue ( $cCell, (int)$nPricePE );
        if ( $lMult )
        {
            $objPHPExcelProEco->setActiveSheetIndex ( 0 )
                              ->getStyle($cCell)
                              ->getFill()->applyFromArray(array('type'=>PHPExcel_Style_Fill::FILL_SOLID,
                                                                'startcolor'=>array('rgb'=>'FFFFCC'),
                                                                'endcolor'  =>array('rgb'=>'FFFFCC')
                                                                )
                                                          );
        }
    }
    //==========================================================================
    //==========================================================================
}
//==============================================================================
//==============================================================================
function send_mail_all ( $cTimeBeg, $cTimeEnd, $cFile )
{
//                                                             по THULE (15%)        
    $cMessage =           'Система мониторинга цен конкурентов<br><br>';
    $cMessage = $cMessage.'товары с 5%-й разницой в цене в сторону уменьшения выделены красным шрифтом ( не заливкой )<br>';
    $cMessage = $cMessage.'товары с большей ценой, чем рекомендована, выделены зеленым шрифтом ( не заливкой )<br>';
    $cMessage = $cMessage.'<br>';
    $cMessage = $cMessage.'Begin - '.$cTimeBeg.'<br>';
    $cMessage = $cMessage.'End   - '.$cTimeEnd;
    
    $mail = new PHPMailer ( true ); // the true param means it will throw exceptions on errors, which we need to catch
    
    $mail->IsSMTP();                // telling the class to use SMTP
    
    try {
        $mail->CharSet='UTF-8';
        $mail->SMTPDebug  = 0;                      // enables SMTP debug information (for testing)
        $mail->SMTPAuth   = true;                   // enable SMTP authentication
        $mail->SMTPSecure = 'tls';                  // sets the prefix to the servier
        
        $mail->Host       = 'smtp.googlemail.com';  // sets the SMTP server
        $mail->Port       = 587;                    // sets the SMTP port for the GMAIL server     587    465
        $mail->Username   = 'baza@ampplus.com.ua';  // SMTP account username
        $mail->Password   = 'zz-D$BT7';             // SMTP account password
        
//      $mail->AddReplyTo('name@yourdomain.com', 'First Last');
        
        $mail->AddAddress ( 'rybasov@ampplus.com.ua', 'Константин');
        $mail->AddCC      ( 'Ruslan@skubenko.com.ua', 'Руслан');
        $mail->AddCC      ( 'gurzhiy@lightf.com.ua', 'Дмитрий Гуржий');
        $mail->AddCC      ( 'pds@sotra.com.ua', 'Дмитрий Потапчук');
        $mail->AddCC      ( 'bam@sotra.com.ua', 'Александр Березинец');
        $mail->AddCC      ( 'sh@bagazhnik.ua', 'Сергей Швец');
        $mail->AddCC      ( 'bodyle15@gmail.com', 'Богдан Колодий');
        $mail->AddCC      ( 'vns@ampplus.com.ua', 'Валерий Скубенко');
        
        $mail->SetFrom ( 'baza@ampplus.com.ua', 'Константин Рыбасов');
        
//      $mail->AddReplyTo('name@yourdomain.com', 'First Last');
        
//                                                               по THULE (15%)        
        $mail->Subject = 'Система мониторинга цен конкурентов';
        
//      $mail->AltBody = 'To view the message, please use an HTML compatible email viewer!'; // optional - MsgHTML will create an alternate automatically
//      $mail->AltBody =$message;
        
        $mail->MsgHTML ( $cMessage );
        $mail->AddAttachment ( $cFile );         // attachment
//      $mail->AddAttachment('images/phpmailer_mini.gif'); // attachment
        
        $mail->Send();
        
        //echo "Message Sent OK<p></p>\n";
    }
    catch ( phpmailerException $e )
    {
        echo $e->errorMessage(); //Pretty error messages from PHPMailer
    }
    catch ( Exception $e )
    {
        echo $e->getMessage(); //Boring error messages from anything else!
    }
}
//==============================================================================
//==============================================================================
function send_mail_pe ( $cTimeBeg, $cTimeEnd, $cFile )
{
//                                                             по THULE (15%)        
    $cMessage = 'Цены для ProEco на основе прайс-листов Sotra, ThuleShop, Multibox';
    
    $mail = new PHPMailer ( true ); // the true param means it will throw exceptions on errors, which we need to catch
    
    $mail->IsSMTP();                // telling the class to use SMTP
    
    try {
        $mail->CharSet='UTF-8';
        $mail->SMTPDebug  = 0;                      // enables SMTP debug information (for testing)
        $mail->SMTPAuth   = true;                   // enable SMTP authentication
        $mail->SMTPSecure = 'tls';                  // sets the prefix to the servier
        
        $mail->Host       = 'smtp.googlemail.com';  // sets the SMTP server
        $mail->Port       = 587;                    // sets the SMTP port for the GMAIL server     587    465
        $mail->Username   = 'baza@ampplus.com.ua';  // SMTP account username
        $mail->Password   = 'zz-D$BT7';             // SMTP account password
        
//      $mail->AddReplyTo('name@yourdomain.com', 'First Last');
        
        $mail->AddAddress ( 'rybasov@ampplus.com.ua', 'Константин');
        $mail->AddCC      ( 'bodyle15@gmail.com', 'Богдан Колодий');
        
        $mail->SetFrom ( 'baza@ampplus.com.ua', 'Константин Рыбасов');
        
//      $mail->AddReplyTo('name@yourdomain.com', 'First Last');
        
//                                                               по THULE (15%)        
        $mail->Subject = 'Цены для ProEco';
        
//      $mail->AltBody = 'To view the message, please use an HTML compatible email viewer!'; // optional - MsgHTML will create an alternate automatically
//      $mail->AltBody =$message;
        
        $mail->MsgHTML ( $cMessage );
        $mail->AddAttachment ( $cFile );         // attachment
//      $mail->AddAttachment('images/phpmailer_mini.gif'); // attachment
        
        $mail->Send();
        
        //echo "Message Sent OK<p></p>\n";
    }
    catch ( phpmailerException $e )
    {
        echo $e->errorMessage(); //Pretty error messages from PHPMailer
    }
    catch ( Exception $e )
    {
        echo $e->getMessage(); //Boring error messages from anything else!
    }
}
//==============================================================================
//==============================================================================
////////////////////////////////////////////////////////////////////////////////

$price_car     = $argv [1];
$first_payment = $argv [2];
$percent       = $argv [3];

$arr_term = array ( 1 => 12, 2 => 24, 3 => 36, 4 => 48, 5 => 60 );

echo 'price_car = '.$price_car;
echo "\r\n";
echo 'first_payment = '.$first_payment;
echo "\r\n";
echo 'percent = '.$percent;
echo "\r\n";
echo "\r\n";
var_dump ( $arr_term );
//------------------------------------------------------------------------------
//include_once ( 'parser_manual.php' );
include_once ( 'PHPExcel/IOFactory.php' );

$objPHPExcel = PHPExcel_IOFactory::load($input);
$aSheet = $objPHPExcel->getActiveSheet();
$aData  = $objPHPExcel->getActiveSheet()->toArray(null,true,true,true);
//------------------------------------------------------------------------------
$objPHPExcelOutput = new PHPExcel();
$objPHPExcelOutput->getProperties()->setCreator("PHP")
                  ->setLastModifiedBy("Konstantin")
                  ->setTitle("Office 2007 XLSX")
                  ->setSubject("Office 2007 XLSX")
                  ->setDescription("File Office 2007 XLSX, by PHPExcel.")
                  ->setKeywords("office 2007 openxml php")
                  ->setCategory("Test");
$objPHPExcelOutput->getActiveSheet()->setTitle('Price');
//------------------------------------------------------------------------------
$cTime = date ( 'Y-m-d H:i:s' );
$cDate = date ( "Y-m-d H:i" );

$nRow = 0;
$nCol = 0;

foreach($aSheet->getRowIterator() as $row)
{
    $cellIterator = $row->getCellIterator();
    
    $nRow = $nRow + 1;
    $nN = 0;
    
    foreach ( $cellIterator as $cell )
    {
        $nN = $nN + 1;
        $nCol = max ( $nCol, $nN);
    }
}
//==============================================================================
$cCell = "D2";

$objPHPExcelOutput->setActiveSheetIndex ( 0 )
                  ->setCellValue ( $cCell, 'Sotra' );

$objPHPExcelOutput->setActiveSheetIndex ( 0 )
                  ->getStyle($cCell)
                  ->getAlignment()
                  ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    
$objPHPExcelOutput->setActiveSheetIndex ( 0 )
                  ->getColumnDimension("D")->setAutoSize(true);
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - -
$cCell = "E2";

$objPHPExcelOutput->setActiveSheetIndex ( 0 )
                  ->setCellValue ( $cCell, 'ProEco' );

$objPHPExcelOutput->setActiveSheetIndex ( 0 )
                  ->getStyle($cCell)
                  ->getAlignment()
                  ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    
$objPHPExcelOutput->setActiveSheetIndex ( 0 )
                  ->getColumnDimension("E")->setAutoSize(true);
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - -
$cCell = "F2";

$objPHPExcelOutput->setActiveSheetIndex ( 0 )
                  ->setCellValue ( $cCell, 'Thule Shop' );

$objPHPExcelOutput->setActiveSheetIndex ( 0 )
                  ->getStyle($cCell)
                  ->getAlignment()
                  ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    
$objPHPExcelOutput->setActiveSheetIndex ( 0 )
                  ->getColumnDimension("F")->setAutoSize(true);
//------------------------------------------------------------------------------
$cCell = "C2";

$objPHPExcelProEco->setActiveSheetIndex ( 0 )
                  ->setCellValue ( $cCell, 'Sotra' );

$objPHPExcelProEco->setActiveSheetIndex ( 0 )
                  ->getStyle($cCell)
                  ->getAlignment()
                  ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    
$objPHPExcelProEco->setActiveSheetIndex ( 0 )
                  ->getColumnDimension("C")->setAutoSize(true);
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - -
$cCell = "D2";

$objPHPExcelProEco->setActiveSheetIndex ( 0 )
                  ->setCellValue ( $cCell, 'Thule Shop' );

$objPHPExcelProEco->setActiveSheetIndex ( 0 )
                  ->getStyle($cCell)
                  ->getAlignment()
                  ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    
$objPHPExcelProEco->setActiveSheetIndex ( 0 )
                  ->getColumnDimension("D")->setAutoSize(true);
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - -
$cCell = "E2";

$objPHPExcelProEco->setActiveSheetIndex ( 0 )
                  ->setCellValue ( $cCell, 'multibox' );

$objPHPExcelProEco->setActiveSheetIndex ( 0 )
                  ->getStyle($cCell)
                  ->getAlignment()
                  ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    
$objPHPExcelProEco->setActiveSheetIndex ( 0 )
                  ->getColumnDimension("E")->setAutoSize(true);
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - -
$cCell = "F2";

$objPHPExcelProEco->setActiveSheetIndex ( 0 )
                  ->setCellValue ( $cCell, 'ProEco' );

$objPHPExcelProEco->setActiveSheetIndex ( 0 )
                  ->getStyle($cCell)
                  ->getAlignment()
                  ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    
$objPHPExcelProEco->setActiveSheetIndex ( 0 )
                  ->getColumnDimension("F")->setAutoSize(true);
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - -
$cCell = "C1:F2";

$objPHPExcelProEco->setActiveSheetIndex ( 0 )
                  ->getStyle($cCell)
                  ->getBorders()
                  ->getAllBorders()
                  ->applyFromArray(array('style' =>PHPExcel_Style_Border::BORDER_THIN,'color'=>array('rgb' => '000000')));
//------------------------------------------------------------------------------
for ( $nK=4; $nK<=$nCol; $nK++ )
{
    $cK = get_nom_column ( $nK );
    
    $cNameSite = $aData[1][$cK];
    
    $nIDClient = $aData[2][$cK];
    
    if ( ! check_client ( $db, $nIDClient ) )
    {
        insert_client ( $db, $nIDClient );
    }
    
    if ( ! check_site ( $db, $nIDClient, $cNameSite ) )
    {
        insert_site ( $db, $nIDClient, $cNameSite );
    }
    //-------------------------------------------------------------------------
    $cK = get_nom_column ( $nK + 3 );
    
    $cCell = $cK."1";
    
    $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                      ->setCellValue ( $cCell, $nIDClient );
    
    $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                      ->getStyle($cCell)
                      ->getAlignment()
                      ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    $cCell = $cK."2";
    
    $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                      ->setCellValue ( $cCell, $cNameSite );
    
    $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                      ->getStyle($cCell)
                      ->getAlignment()
                      ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    
    $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                      ->getColumnDimension($cK)->setAutoSize(true);
    //-------------------------------------------------------------------------
    $cCell = "D1".":".$cK."2";
    
    $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                      ->getStyle($cCell)
                      ->getBorders()
                      ->getAllBorders()
                      ->applyFromArray(array('style' =>PHPExcel_Style_Border::BORDER_THIN,'color'=>array('rgb' => '000000')));
}
//============================================================================================================================
for ( $nN=3; $nN<=$nRow; $nN++ ) {
    
    $nIDTovar   = preg_replace('/^(0123456789)/', '', $aData[$nN]['A'] );
    $cNameTovar = $aData[$nN]['B'];
    $cAlter     = $aData[$nN]['C'];
    
    if ( ! check_tovar ( $db, $nIDTovar ) )
    {
        insert_tovar ( $db, $nIDTovar, $cNameTovar, $cAlter );
    }
    
    $cCell = "A".$nN;
    
    $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                      ->setCellValue ( $cCell, $nIDTovar );
    $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                      ->getColumnDimension("A")->setAutoSize(true);
    
    $cCell = "B".$nN;
    
    $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                      ->setCellValue ( $cCell, $cNameTovar );
    $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                      ->getColumnDimension("B")->setAutoSize(true);
    
    $cCell = "C".$nN;
    
    $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                      ->setCellValue ( $cCell, $cAlter );
    $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                      ->getColumnDimension("C")->setAutoSize(true);
    //--------------------------------------------------------------------------
    $cCell = "A3:C".$nRow;
    
    $objPHPExcelOutput->setActiveSheetIndex ( 0 )
                      ->getStyle($cCell)
                      ->getBorders()
                      ->getAllBorders()
                      ->applyFromArray(array('style' =>PHPExcel_Style_Border::BORDER_THIN,'color'=>array('rgb' => '000000')));
    //==========================================================================
    $cCell = "A".$nN;
    
    $objPHPExcelProEco->setActiveSheetIndex ( 0 )
                      ->setCellValue ( $cCell, $nIDTovar );
    $objPHPExcelProEco->setActiveSheetIndex ( 0 )
                      ->getColumnDimension("A")->setAutoSize(true);
    
    $cCell = "B".$nN;
    
    $objPHPExcelProEco->setActiveSheetIndex ( 0 )
                      ->setCellValue ( $cCell, $cAlter );
    $objPHPExcelProEco->setActiveSheetIndex ( 0 )
                      ->getColumnDimension("C")->setAutoSize(true);
    //--------------------------------------------------------------------------
    $cCell = "A3:B".$nRow;
    
    $objPHPExcelProEco->setActiveSheetIndex ( 0 )
                      ->getStyle($cCell)
                      ->getBorders()
                      ->getAllBorders()
                      ->applyFromArray(array('style' =>PHPExcel_Style_Border::BORDER_THIN,'color'=>array('rgb' => '000000')));
}
//============================================================================================================================
for ( $nN=3; $nN<=$nRow; $nN++ ) 
{
    //$nIDTovar = str_replace ( ' ', '', $aData[$nN]['A'] );
    
    $nPriceSrav = 0;
    
    $nIDTovar = preg_replace('/^(0123456789)/', '', $aData[$nN]['A'] );
    
    echo "\r\n".$nIDTovar;
    
    output_excel_my ( $nN, 4, $nIDTovar, $nKursEur, $aProdSotra, $aProdPE, $aProdTH, $objPHPExcelOutput, $nPriceSrav );
    
    output_excel_PE ( $nN, 3, $nIDTovar, $nKursEur, $aProdSotra, $aProdPE, $aProdTH, $objPHPExcelProEco, $nPriceSrav );
    
    for ( $nK=4; $nK<=$nCol; $nK++ )
    {
        $cK = get_nom_column ( $nK );
        
        $cNameSite = $aData[1][$cK];
        $nIDClient = $aData[2][$cK];
        
        $cWWW = $aData[$nN][$cK];
        
        $cK = get_nom_column ( $nK + 3 );
        
        $nPrice = 'NULL';
        $nPriceOld = 'NULL';
        
        $l = get_inform ( $db, $cWWW, $nIDClient, $cNameSite, $nIDSite, $nIDTovar, $nPrice, $nPriceOld, $nAvail, $nKursEur, $cValuta, $nKurs );
        
        if ( $l )
        {
           insert_work ( $db, $nIDClient, $nIDSite, $nIDTovar, $cWWW, $nPrice, $nPriceOld, $nAvail, $cValuta, $nKurs, $cTime );
        }
           
        output_excel_cell ( $nN, $nK, $nPrice, $nKurs, $nPriceOld, $nAvail, $l, $cWWW, $objPHPExcelOutput, $objPHPExcelProEco, $nPriceSrav );
    }
}

echo "\r\n";
//------------------------------------------------------------------------------
$cK = get_nom_column ( $nCol + 3 );

$cCell = "D3:".$cK.$nRow;

$objPHPExcelOutput->setActiveSheetIndex ( 0 )
                  ->getStyle($cCell)
                  ->getBorders()
                  ->getAllBorders()
                  ->applyFromArray(array('style' =>PHPExcel_Style_Border::BORDER_THIN,'color'=>array('rgb' => '000000')));

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcelOutput, 'Excel2007');

$cFile = $cDate;
$cFile = 'RESULT\\' . substr ( $cFile, 0, 4 ) . substr ( $cFile, 5, 2 ) . substr ( $cFile, 8, 2 ) . substr ( $cFile, 11, 2 ) . '.xlsx';

$objWriter->save ( $cFile );
//------------------------------------------------------------------------------
$cK = get_nom_column ( 6 );

$cCell = "C3:".$cK.$nRow;

$objPHPExcelProEco->setActiveSheetIndex ( 0 )
                  ->getStyle($cCell)
                  ->getBorders()
                  ->getAllBorders()
                  ->applyFromArray(array('style' =>PHPExcel_Style_Border::BORDER_THIN,'color'=>array('rgb' => '000000')));

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcelProEco, 'Excel2007');

$cFilePE = $cDate;
$cFilePE = 'ProEco\\PE_' . substr ( $cFilePE, 0, 4 ) . substr ( $cFilePE, 5, 2 ) . substr ( $cFilePE, 8, 2 ) . substr ( $cFilePE, 11, 2 ) . '.xlsx';

$objWriter->save ( $cFilePE );
//------------------------------------------------------------------------------
echo "\r\n";
echo "\r\n";
//==============================================================================
?>