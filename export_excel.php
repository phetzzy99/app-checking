<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if (isset($_POST['data'])) {
    $data = json_decode($_POST['data'], true);

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setTitle('ID Card Data');

    // Write headers
    $sheet->setCellValue('A1', 'รหัสสมาชิก');
    $sheet->setCellValue('B1', 'ชื่อ-สกุล');
    $sheet->setCellValue('C1', 'Barcode');
    $sheet->setCellValue('D1', 'ชื่อเรื่อง');
    $sheet->setCellValue('E1', 'Call No.');
    $sheet->setCellValue('F1', 'วันกำหนดส่ง');
    $sheet->setCellValue('G1', 'วันที่คืน');
    $sheet->setCellValue('H1', 'ค่าปรับ');

    $row = 2;
    foreach ($data as $item) {
        $sheet->setCellValue('A' . $row, $item['id_card']);
        $sheet->setCellValue('B' . $row, $item['name_card']);
        $sheet->setCellValue('C' . $row, $item['BARCODE']);
        $sheet->setCellValue('D' . $row, $item['TITLE']);
        $sheet->setCellValue('E' . $row, $item['CALLNO']);
        $sheet->setCellValue('F' . $row, $item['DUEDATE']);
        $sheet->setCellValue('G' . $row, $item['RETURNDATE']);
        $sheet->setCellValue('H' . $row, $item['TOTALMONEY']);
        $row++;
    }

    $writer = new Xlsx($spreadsheet);
    $filename = 'id_card_data_' . date('Y-m-d_His') . '.xlsx';

    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="' . $filename . '"');
    header('Cache-Control: max-age=0');

    $file_path = 'exported_files/' . $filename;
    $writer->save($file_path);
    
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="' . $filename . '"');
    header('Cache-Control: max-age=0');
    
    readfile($file_path);
    unlink($file_path);
    exit;
}
?>