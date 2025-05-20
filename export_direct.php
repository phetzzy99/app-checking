<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Color;

session_start();

// ตรวจสอบว่ามีข้อมูลสำหรับ export หรือไม่
if (!isset($_SESSION['export_data']) || empty($_SESSION['export_data'])) {
    echo "<script>
        alert('ไม่พบข้อมูลสำหรับการ Export โปรดอัปโหลดไฟล์ Excel ก่อน');
        window.location.href = 'fine-check.php';
    </script>";
    exit;
}

$results = $_SESSION['export_data'];
$filter = isset($_GET['filter']) ? $_GET['filter'] : 'all';

// กรองข้อมูลตามเงื่อนไข
if ($filter === 'mismatch') {
    $filtered_results = [];
    foreach ($results as $item) {
        if ($item['has_debt']) {
            $filtered_results[] = $item;
        }
    }
    $results = $filtered_results;
} elseif ($filter === 'no_mismatch') {
    $filtered_results = [];
    foreach ($results as $item) {
        if (!$item['has_debt']) {
            $filtered_results[] = $item;
        }
    }
    $results = $filtered_results;
}

// สร้าง spreadsheet ใหม่
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('รายงานค่าปรับหนังสือ');

// กำหนดความกว้างของคอลัมน์
$sheet->getColumnDimension('A')->setWidth(10);
$sheet->getColumnDimension('B')->setWidth(20);
$sheet->getColumnDimension('C')->setWidth(30);

// ตรวจสอบว่าเป็นการ export เฉพาะรายการที่ไม่มีพันธะหรือไม่
if ($filter === 'no_mismatch') {
    $sheet->getColumnDimension('D')->setWidth(20);
    
    // ตั้งค่าส่วนหัวตาราง
    $sheet->setCellValue('A1', 'ลำดับ');
    $sheet->setCellValue('B1', 'รหัสสมาชิก');
    $sheet->setCellValue('C1', 'ชื่อ-สกุล');
    $sheet->setCellValue('D1', 'สถานะ');
    
    // จัดรูปแบบส่วนหัวตาราง
    $headerStyle = [
        'font' => [
            'bold' => true,
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical' => Alignment::VERTICAL_CENTER,
        ],
        'fill' => [
            'fillType' => Fill::FILL_SOLID,
            'startColor' => [
                'rgb' => 'CCCCCC',
            ],
        ],
        'borders' => [
            'allBorders' => [
                'borderStyle' => Border::BORDER_THIN,
            ],
        ],
    ];
    
    $sheet->getStyle('A1:D1')->applyFromArray($headerStyle);
    
    // เพิ่มข้อมูลในตาราง
    $row = 2;
    $index = 1;
    
    foreach ($results as $item) {
        $id_card = $item['id_card'];
        $name_card = $item['name_card'];
        
        $sheet->setCellValue('A' . $row, $index);
        $sheet->setCellValue('B' . $row, $id_card);
        $sheet->setCellValue('C' . $row, $name_card);
        $sheet->setCellValue('D' . $row, 'ไม่มีพันธะ');
        
        // กำหนดสีพื้นหลังเป็นสีเขียว
        $sheet->getStyle('A' . $row . ':D' . $row)->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()->setRGB('90EE90');
        
        // จัดตำแหน่งข้อความในเซลล์
        $sheet->getStyle('A' . $row)->getAlignment()
            ->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('D' . $row)->getAlignment()
            ->setHorizontal(Alignment::HORIZONTAL_CENTER);
        
        // เพิ่มเส้นขอบตาราง
        $sheet->getStyle('A' . $row . ':D' . $row)->getBorders()
            ->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
        
        $row++;
        $index++;
    }
    
    // เพิ่มแถวสรุปจำนวนรายการ
    $sheet->setCellValue('A' . $row, 'สรุป');
    $sheet->setCellValue('B' . $row, 'จำนวนรายการที่ไม่มีพันธะ:');
    $sheet->setCellValue('C' . $row, count($results));
    $sheet->getStyle('A' . $row . ':C' . $row)->getFont()->setBold(true);
    $sheet->getStyle('A' . $row . ':D' . $row)->getFill()
        ->setFillType(Fill::FILL_SOLID)
        ->getStartColor()->setRGB('EFEFEF');
    $sheet->getStyle('A' . $row . ':D' . $row)->getBorders()
        ->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
} else {
    // กรณี export ทั้งหมดหรือเฉพาะรายการที่มีพันธะ
    $sheet->getColumnDimension('D')->setWidth(60);
    $sheet->getColumnDimension('E')->setWidth(20);
    
    // ตั้งค่าส่วนหัวตาราง
    $sheet->setCellValue('A1', 'ลำดับ');
    $sheet->setCellValue('B1', 'รหัสสมาชิก');
    $sheet->setCellValue('C1', 'ชื่อ-สกุล');
    $sheet->setCellValue('D1', 'รายการหนังสือที่เกินกำหนด');
    $sheet->setCellValue('E1', 'สถานะ');
    
    // จัดรูปแบบส่วนหัวตาราง
    $headerStyle = [
        'font' => [
            'bold' => true,
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical' => Alignment::VERTICAL_CENTER,
        ],
        'fill' => [
            'fillType' => Fill::FILL_SOLID,
            'startColor' => [
                'rgb' => 'CCCCCC',
            ],
        ],
        'borders' => [
            'allBorders' => [
                'borderStyle' => Border::BORDER_THIN,
            ],
        ],
    ];
    
    $sheet->getStyle('A1:E1')->applyFromArray($headerStyle);
    
    // เพิ่มข้อมูลในตาราง
    $row = 2;
    $index = 1;
    
    foreach ($results as $item) {
        $id_card = $item['id_card'];
        $name_card = $item['name_card'];
        $result = $item['items'];
        $has_debt = $item['has_debt'];
        
        $sheet->setCellValue('A' . $row, $index);
        $sheet->setCellValue('B' . $row, $id_card);
        $sheet->setCellValue('C' . $row, $name_card);
        
        // เพิ่มข้อมูลรายการหนังสือ
        $bookDetails = '';
        if (!empty($result)) {
            foreach ($result as $subitem) {
                $bookDetails .= "Barcode: " . $subitem['BARCODE'] . "\n";
                $bookDetails .= "ชื่อเรื่อง: " . $subitem['TITLE'] . "\n";
                $bookDetails .= "Call No.: " . $subitem['CALLNO'] . "\n";
                $bookDetails .= "วันกำหนดส่ง: " . $subitem['DUEDATE'] . "\n";
                $bookDetails .= "วันที่คืน: " . $subitem['RETURNDATE'] . "\n";
                $bookDetails .= "ค่าปรับ: " . $subitem['TOTALMONEY'] . " บาท\n\n";
            }
        } else {
            $bookDetails = "ไม่มีรายการหนังสือที่เกินกำหนด";
        }
        
        $sheet->setCellValue('D' . $row, $bookDetails);
        $sheet->getStyle('D' . $row)->getAlignment()->setWrapText(true);
        
        // กำหนดสถานะและจัดรูปแบบตามเงื่อนไข
        if ($has_debt) {
            $sheet->setCellValue('E' . $row, 'ค้างค่าปรับ 1 รายการ');
            
            // กำหนดสีพื้นหลังเป็นสีแดง
            $sheet->getStyle('A' . $row . ':E' . $row)->getFill()
                ->setFillType(Fill::FILL_SOLID)
                ->getStartColor()->setRGB('FF0000');
                
            // ตั้งค่าตัวอักษรเป็นสีขาว
            $sheet->getStyle('A' . $row . ':E' . $row)->getFont()
                ->setColor(new Color(Color::COLOR_WHITE));
        } else {
            $sheet->setCellValue('E' . $row, 'ไม่มีพันธะ');
            
            // กำหนดสีพื้นหลังเป็นสีเขียว
            $sheet->getStyle('A' . $row . ':E' . $row)->getFill()
                ->setFillType(Fill::FILL_SOLID)
                ->getStartColor()->setRGB('90EE90');
        }
        
        // จัดตำแหน่งข้อความในเซลล์
        $sheet->getStyle('A' . $row)->getAlignment()
            ->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('E' . $row)->getAlignment()
            ->setHorizontal(Alignment::HORIZONTAL_CENTER);
        
        // เพิ่มเส้นขอบตาราง
        $sheet->getStyle('A' . $row . ':E' . $row)->getBorders()
            ->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
        
        $row++;
        $index++;
    }
    
    // เพิ่มแถวสรุปจำนวนรายการ
    $sheet->setCellValue('A' . $row, 'สรุป');
    
    if ($filter === 'mismatch') {
        $sheet->setCellValue('B' . $row, 'จำนวนรายการที่มีพันธะ:');
    } else {
        $sheet->setCellValue('B' . $row, 'จำนวนรายการทั้งหมด:');
    }
    
    $sheet->setCellValue('C' . $row, count($results));
    $sheet->getStyle('A' . $row . ':C' . $row)->getFont()->setBold(true);
    $sheet->getStyle('A' . $row . ':E' . $row)->getFill()
        ->setFillType(Fill::FILL_SOLID)
        ->getStartColor()->setRGB('EFEFEF');
    $sheet->getStyle('A' . $row . ':E' . $row)->getBorders()
        ->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
}

// กำหนด folder สำหรับบันทึกไฟล์ชั่วคราว
$temp_folder = 'temp_excel';
if (!is_dir($temp_folder)) {
    mkdir($temp_folder, 0777, true);
}

// สร้างชื่อไฟล์ที่มีความปลอดภัย
$filename = '';
if ($filter === 'mismatch') {
    $filename = 'report_with_debt_' . date('Y-m-d_H-i-s') . '.xlsx';
} elseif ($filter === 'no_mismatch') {
    $filename = 'report_without_debt_' . date('Y-m-d_H-i-s') . '.xlsx';
} else {
    $filename = 'report_all_' . date('Y-m-d_H-i-s') . '.xlsx';
}

$file_path = $temp_folder . '/' . $filename;

// บันทึกไฟล์ Excel
$writer = new Xlsx($spreadsheet);
$writer->save($file_path);

// ล้าง output buffer
if (ob_get_length()) {
    ob_clean();
}

// กำหนด header สำหรับการดาวน์โหลดไฟล์ Excel
header('Content-Description: File Transfer');
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="' . $filename . '"');
header('Content-Transfer-Encoding: binary');
header('Expires: 0');
header('Cache-Control: must-revalidate');
header('Pragma: public');
header('Content-Length: ' . filesize($file_path));

// อ่านไฟล์และส่งไปยังเบราว์เซอร์
readfile($file_path);

// ลบไฟล์หลังจากส่งเสร็จ
unlink($file_path);
exit;