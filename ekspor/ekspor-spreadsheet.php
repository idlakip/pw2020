<?php
//Menggabungkan dengan file koneksi yang telah kita buat
require '../functions.php';

// Load library phpspreadsheet
require('../vendor/autoload.php');

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
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
$spreadsheet->getProperties()
	->setCreator('Masrianto')
	->setLastModifiedBy('www.lakip.co.id')
	->setCompany('www.lakip.co.id')
	->setTitle('Office XLSX LAKIP.CO.ID')
	->setSubject('Office Report XLSX LAKIP.CO.ID')
	->setDescription('Document for Office XLSX LAKIP.CO.ID.')
	->setKeywords('office openxml php Masrianto')
	->setCategory('Result Masrianto & LAKIP.CO.ID');

// Set protection
$sheet->getProtection()->setSheet(true);

// set Company
$spreadsheet->getProperties()->getCreated();
$spreadsheet->getProperties()->getLastModifiedBy();
// set Hyperlink
$cell->getHyperlink()->getUrl($url);
$drawing->getHyperlink()->getUrl();
$drawing->setHyperlink()->setUrl($url);


$spreadsheet->getActiveSheet()->mergeCells('A1:F1');
$spreadsheet->setActiveSheetIndex(0)->setCellValue('A1', 'Ekspor Laporan/Data dari Database MySQL ke dalam Excel (.xlsx)');


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
$spreadsheet->getActiveSheet()->setTitle('Report Excel ' . date('d-m-Y H'));

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);

// Redirect output to a client’s web browser (Xlsx)
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="Report LAKIP Excel.xlsx"');
header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
header('Pragma: public'); // HTTP/1.0

$writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
$writer->save('php://output');
