<?php

// require 'vendor/autoload.php';
// require 'functions.php';
// $mahasiswa = query("SELECT * FROM mahasiswa");

// use PhpOffice\PhpSpreadsheet\Spreadsheet;
// use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// $spreadsheet = new Spreadsheet();
// $sheet = $spreadsheet->getActiveSheet();
// $sheet->setCellValue('A1', 'Hello World !');

// $writer = new Xlsx($spreadsheet);
// $writer->save('Ekspor Excel.xlsx');


//Menggabungkan dengan file koneksi yang telah kita buat
require 'functions.php';

// Load library phpspreadsheet
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\Reader\IReader;
use PhpOffice\PhpSpreadsheet\Writer\IWriter;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
// use PhpOffice\PhpSpreadsheet\Worksheet\HeaderFooterDrawing;
// use PhpOffice\PhpSpreadsheet\Calculation\DateTime;
// use PhpOffice\PhpSpreadsheet\Calculation\Logical;
// use PhpOffice\PhpSpreadsheet\Calculation\LookupRef::VLOOKUP;
// use PhpOffice\PhpSpreadsheet\Calculation\MathTrig;
// use PhpOffice\PhpSpreadsheet\Calculation\Statistical;
// use PhpOffice\PhpSpreadsheet\Calculation\TextData;
// use \PhpOffice\PhpSpreadsheet\Calculation\Web::WEBSERVICE;
// End load library phpspreadsheet

$spreadsheet = new Spreadsheet();

// Set document properties
$spreadsheet->getProperties()->setCreator('LAKIP.CO.ID')
  ->setLastModifiedBy('www.lakip.co.id')
  ->setCompany('Lembaga Administrasi Keuangan dan Ilmu Pemerintahan - LAKIP')
  ->setTitle('Report XLSX LAKIP.CO.ID')
  ->setSubject('Print By : LAKIP.CO.ID')
  ->setDescription('Document for Office 2007 XLSX, generated using PHP classes.')
  ->setKeywords('office 2007 openxml php')
  ->setCategory('Result file XLSX LAKIP.CO.ID');


// $drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\HeaderFooterDrawing();
// $drawing->setName('logo');
// $drawing->setPath('img/lakip.png');
// $drawing->setHeight(36);
// $spreadsheet->getActiveSheet()->getHeaderFooter()->addImage($drawing, \PhpOffice\PhpSpreadsheet\Worksheet\HeaderFooter::IMAGE_HEADER_LEFT);
// $spreadsheet->getActiveSheet()->mergeCells('A1:F1');



$spreadsheet->getActiveSheet()->mergeCells('A1:F1');
$spreadsheet->setActiveSheetIndex(0)->setCellValue('A1', 'Ekspor Laporan/Data dari Database MySQL ke dalam Excel (.xlsx)');

// logo
$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
$drawing->setName('Logo');
$drawing->setDescription('Logo');
$drawing->setPath('img/lakip.png');
$drawing->setCoordinates('A1');
$drawing->setHeight(75);
// $drawing->setWorksheet($spreadsheet->getActiveSheet('A2'));
$drawing->setWorksheet($spreadsheet->getActiveSheet('A2'));
$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);


//Font Color
$spreadsheet->getActiveSheet()->getStyle('A3:F3')
  ->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE);

// Background color
$spreadsheet->getActiveSheet()->getStyle('A3:F3')->getFill()
  ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
  ->getStartColor()->setARGB('FFFF0000');


// Header Tabel
$spreadsheet->setActiveSheetIndex(0)
  ->setCellValue('A3', 'NO')
  ->setCellValue('B3', 'NRP')
  ->setCellValue('C3', 'NAMA')
  ->setCellValue('D3', 'EMAIL')
  ->setCellValue('E3', 'JURUSAN')
  ->setCellValue('F3', 'GAMBAR');

$i = 4;
$no = 1;
// $query = "SELECT * FROM tbl_mahasiswa ORDER BY nama_mahasiswa ASC";
// $query = "SELECT * FROM mahasiswa ORDER BY nama ASC";
$query = "SELECT * FROM mahasiswa ORDER BY id ASC";
$conn = koneksi();
$lakip = $conn->prepare($query);
$lakip->execute();
// $drawing->execute();
$result1 = $lakip->get_result();
while ($row = $result1->fetch_assoc()) {
  $spreadsheet->setActiveSheetIndex(0)
    ->setCellValue('A' . $i, $no)
    ->setCellValue('B' . $i, $row['nrp'])
    ->setCellValue('C' . $i, $row['nama'])
    ->setCellValue('D' . $i, $row['email'])
    ->setCellValue('E' . $i, $row['jurusan'])
    ->setCellValue('F' . $i, $row['gambar']);
  $i++;
  $no++;
}


// Rename worksheet
$spreadsheet->getActiveSheet()->setTitle('Report LAKIP ' . date('d-m-Y Hi'));

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);

// Redirect output to a client’s web browser (Xlsx)
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="Report LAKIP XLSX.xlsx"');
header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
header('Pragma: public'); // HTTP/1.0

$spreadsheet->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 3);
$spreadsheet->getActiveSheet()->getPageSetup()->setPrintArea('A1:F10');

// $spreadsheet->getActiveSheet()->getStyle('A4:F7')
//   ->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED);
// $spreadsheet->getActiveSheet()->getStyle('A4:F7')
//   ->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);
// $spreadsheet->getActiveSheet()->getStyle('A4:F7')
//   ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK);
// $spreadsheet->getActiveSheet()->getStyle('A4:F7')
//   ->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK);
// $spreadsheet->getActiveSheet()->getStyle('A4:F7')
//   ->getBorders()->getLeft()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK);
// $spreadsheet->getActiveSheet()->getStyle('A4:F7')
//   ->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK);
// $spreadsheet->getActiveSheet()->getStyle('A4:F7')
//   ->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
// $spreadsheet->getActiveSheet()->getStyle('A4:F4')
//   ->getFill()->getStartColor()->setARGB('FFFF0000');



$spreadsheet->getActiveSheet()->getPageSetup()->setFitToWidth(1);
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToHeight(0);

$spreadsheet->getActiveSheet()->getPageMargins()->setTop(1);
$spreadsheet->getActiveSheet()->getPageMargins()->setRight(0.75);
$spreadsheet->getActiveSheet()->getPageMargins()->setLeft(0.75);
$spreadsheet->getActiveSheet()->getPageMargins()->setBottom(1);

$spreadsheet->getActiveSheet()->getPageSetup()->setHorizontalCentered(false);
$spreadsheet->getActiveSheet()->getPageSetup()->setVerticalCentered(false);

$spreadsheet->getActiveSheet()->getHeaderFooter()
  ->setOddHeader('&C&HPlease treat this document as confidential!');
$spreadsheet->getActiveSheet()->getHeaderFooter()
  ->setOddFooter('&L&B' . $spreadsheet->getProperties()->getTitle() . '&RPage &P of &N');

$writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
$writer->save('php://output');
