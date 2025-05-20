<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

function read_excel($file_path) {
    $spreadsheet = IOFactory::load($file_path);
    $worksheet = $spreadsheet->getActiveSheet();
    $data = [];

    foreach ($worksheet->getRowIterator() as $row) {
        $row_data = [];
        foreach ($row->getCellIterator() as $cell) {
            $row_data[] = $cell->getValue();
        }
        $data[] = $row_data;
    }

    return $data;
}