<?php
require('class/phpexcel18/PHPExcel.php');
require('class/phpexcel18/PHPExcel/Writer/Excel2007.php');

if (isset($_REQUEST['download_excel'])) {
    $alfa = array(
        '1' => 'A', '2' => 'B', '3' => 'C', '4' => 'D', '5' => 'E',
        '6' => 'F', '7' => 'G', '8' => 'H', '9' => 'I', '10' => 'J',
        '11' => 'K', '12' => 'L', '13' => 'M', '14' => 'N', '15' => 'O',
        '16' => 'P', '17' => 'Q', '18' => 'R', '19' => 'S', '20' => 'T',
        '21' => 'U', '22' => 'V', '23' => 'W', '24' => 'X', '25' => 'Y',
        '26' => 'Z'
    );
    $content = utf8_encode(file_get_contents('tourism-sweden.xml'));
    $xml = simplexml_load_string($content);

    $style_header = array(
        'alignment' => array(
            'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
        ),
        'font' => array(
            'bold' => true,
        ),
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'color' => array('rgb' => 'DCDCDC')
        )
    );

    $objPHPExcel = new PHPExcel();

    $objWorkSheet = $objPHPExcel->createSheet(0);
    $objPHPExcel->setActiveSheetIndex(0);

    $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(15);
    $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(15);
    $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(35);
    $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(35);
    $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(35);
    $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(35);
    $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(35);
    $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(35);
    $objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(35);
    $objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(35);
    $objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(35);
    $objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth(35);
    $objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth(35);
    $objPHPExcel->getActiveSheet()->getColumnDimension('N')->setWidth(35);
    $objPHPExcel->getActiveSheet()->getColumnDimension('O')->setWidth(35);
    $objPHPExcel->getActiveSheet()->getColumnDimension('P')->setWidth(35);
    $objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setWidth(35);
    $objPHPExcel->getActiveSheet()->getColumnDimension('R')->setWidth(35);
    $objPHPExcel->getActiveSheet()->getColumnDimension('S')->setWidth(35);
    $objPHPExcel->getActiveSheet()->getColumnDimension('T')->setWidth(35);
    $objPHPExcel->getActiveSheet()->getColumnDimension('U')->setWidth(35);
    $objPHPExcel->getActiveSheet()->getColumnDimension('V')->setWidth(35);
    $objPHPExcel->getActiveSheet()->getColumnDimension('W')->setWidth(35);
    $objPHPExcel->getActiveSheet()->getColumnDimension('X')->setWidth(35);
    $objPHPExcel->getActiveSheet()->getColumnDimension('Y')->setWidth(35);
    $objPHPExcel->getActiveSheet()->getColumnDimension('Z')->setWidth(35);

    $row = 1;
    $col = 0;

    $ar_tourism = array(
        '0' => 'hotel' //get hotel in xml where tags is tourism
    );

    foreach ($ar_tourism as $indx => $vals) {
        $objWorkSheet = $objPHPExcel->createSheet($indx);
        $objWorkSheet->setTitle($vals);
        $objPHPExcel->setActiveSheetIndex($indx);

        $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(15);
        $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(15);
        $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(35);
        $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(35);
        $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(35);
        $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(35);
        $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(35);
        $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(35);
        $objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(35);
        $objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(35);
        $objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(35);
        $objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth(35);
        $objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth(35);
        $objPHPExcel->getActiveSheet()->getColumnDimension('N')->setWidth(35);
        $objPHPExcel->getActiveSheet()->getColumnDimension('O')->setWidth(35);
        $objPHPExcel->getActiveSheet()->getColumnDimension('P')->setWidth(35);
        $objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setWidth(35);
        $objPHPExcel->getActiveSheet()->getColumnDimension('R')->setWidth(35);
        $objPHPExcel->getActiveSheet()->getColumnDimension('S')->setWidth(35);
        $objPHPExcel->getActiveSheet()->getColumnDimension('T')->setWidth(35);
        $objPHPExcel->getActiveSheet()->getColumnDimension('U')->setWidth(35);
        $objPHPExcel->getActiveSheet()->getColumnDimension('V')->setWidth(35);
        $objPHPExcel->getActiveSheet()->getColumnDimension('W')->setWidth(35);
        $objPHPExcel->getActiveSheet()->getColumnDimension('X')->setWidth(35);
        $objPHPExcel->getActiveSheet()->getColumnDimension('Y')->setWidth(35);
        $objPHPExcel->getActiveSheet()->getColumnDimension('Z')->setWidth(35);

        $row = 1;
        $col = 0;

        foreach ($xml->xpath("//node") as $way) {
            if (!isset($way->xpath("tag[@k='tourism']/@v")[0])) continue;
            if ($way->xpath("tag[@k='tourism']/@v")[0] != $vals) continue;

            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($col, $row, 'Lat');
            $col++;
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($col, $row, 'Lon');
            $col++;

            $out = array();
            $nomor = 2;
            foreach ($way->tag as $tag) {
                foreach ((array) $tag as $index => $node) {
                    $out[$index] = (is_object($node)) ? xml2array($node) : $node;
                }

                foreach ($out as $key => $val) {
                    $no = 0;
                    foreach ($val as $valu => $value) {
                        if ($value == 'tourism') {
                            $no++;
                            continue;
                        }
                        if ($no % 2 == 0) {
                            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($col, $row, utf8_decode($value));
                            $col++;
                            $nomor++;
                        }
                        $no++;
                    }
                }
            }
            if ($nomor > 0) {
                if ($nomor > 26) $nomor = 26;
                $objPHPExcel->getActiveSheet()->getStyle('A' . $row . ':' . $alfa[$nomor] . $row)->applyFromArray($style_header);
            }
            $row++;
            $col = 0;

            $lat = $way['lat'];
            $lon = $way['lon'];
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($col, $row, $lat);
            $col++;
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($col, $row, $lon);
            $col++;

            $out = array();
            foreach ($way->tag as $tag) {
                foreach ((array) $tag as $index => $node) {
                    $out[$index] = (is_object($node)) ? xml2array($node) : $node;
                }

                foreach ($out as $key => $val) {
                    $no = 0;
                    foreach ($val as $valu => $value) {
                        if ($value == 'tourism') {
                            $no++;
                            continue;
                        }
                        if ($no % 2 == 0) {
                            $var = isset($way->xpath("tag[@k='" . $value . "']/@v")[0]) ? $way->xpath("tag[@k='" . $value . "']/@v")[0] : '';
                            //echo $var;
                            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($col, $row, utf8_decode($var));
                            $col++;
                        }
                        $no++;
                    }
                }
            }
            $row++;
            $row++;
            $col = 0;
        }
    }

    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="tourism-sweden.xlsx"');
    header('Cache-Control: max-age=0');

    $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
    $objWriter->save('php://output');

    exit;
}
?>
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Export XML to Excel</title>
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous">
</head>

<body>
    <form action="" method="get">
        <div align="center">
            <h1>
                Export to Excel
            </h1>
            <br>
            <input type="submit" class="btn btn-primary col-md-6" name="download_excel" value="Export to Excel">
        </div>
    </form>
</body>

</html>