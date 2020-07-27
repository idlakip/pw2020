<?php
// PROTEKSI
// $protection = $spreadsheet->getActiveSheet()->getProtection();
// $allowed = $protection->verify('my password');

// if ($allowed) {
//   doSomething();
// } else {
//   throw new Exception('Incorrect password');
// }
// // PROTEKSI ALGORITHM_SHA_512
// $protection = $spreadsheet->getActiveSheet()->getProtection();
// $protection->setAlgorithm(Protection::ALGORITHM_SHA_512);
// $protection->setSpinCount(20000);
// $protection->setPassword('PhpSpreadsheet');

//Menggabungkan dengan file koneksi yang telah kita buat
require 'functions.php';

// Load library phpspreadsheet
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\Reader\IReader;
use PhpOffice\PhpSpreadsheet\Writer\IWriter;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


$spreadsheet = new Spreadsheet();

// Set document properties
$spreadsheet->getProperties()->setCreator('LAKIP.CO.ID')
  ->setLastModifiedBy('www.lakip.co.id')
  ->setCompany('Lembaga Administrasi Keuangan dan Ilmu Pemerintahan - LAKIP')
  ->setTitle('Report LAKIP.CO.ID')
  ->setSubject('Print Data Result By LAKIP.CO.ID')
  // ->setDescription('Document for Office 2007 XLSX, generated using PHP classes.')
  ->setDescription('*Dokumen ini telah di tandatangani secara elektronik menggunakan sertifikat elektronik yang diterbitkan oleh lakip.co.id , sehingga tidak diperlukan tandatangan dengan stempel basah.')
  ->setKeywords('office 2007 openxml php www.lakip.co.id')
  ->setCategory('Result file XLSX LAKIP.CO.ID');

// FONT
$spreadsheet->getDefaultStyle()->getFont()->setName('Gisha');
$spreadsheet->getDefaultStyle()->getFont()->setSize(11);

// AUTO WIDTH
$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(12);
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);


// COLAPSE
// $spreadsheet->getActiveSheet()->getColumnDimension('F')->setCollapsed(true);
// $spreadsheet->getActiveSheet()->getColumnDimension('F')->setVisible(false);

// HIDE/UNHIDE
// $spreadsheet->getActiveSheet()->getColumnDimension('D')->setVisible(true);
// $spreadsheet->getActiveSheet()->getColumnDimension('E')->setVisible(false);

// SET ROH HEIGHT
// $spreadsheet->getActiveSheet()->getRowDimension('10')->setRowHeight(100); //ROW No. 10

// Show/hide a row
// $spreadsheet->getActiveSheet()->getRowDimension('10')->setVisible(true); //ok

$spreadsheet->setActiveSheetIndex(0);
$spreadsheet->getActiveSheet()->setCellValue('A1', '');
$spreadsheet->getActiveSheet()->setCellValue('A2', '');
$spreadsheet->getActiveSheet()->setCellValue('A3', '');
$spreadsheet->getActiveSheet()->setCellValue('A4', '');
$spreadsheet->getActiveSheet()->setCellValue('A7', '');
$spreadsheet->getActiveSheet()->setCellValue('A8', '');


$spreadsheet->getActiveSheet()->mergeCells('C1:F1');
$spreadsheet->getActiveSheet()->setCellValue('C1', 'LEMBAGA ADMINISTRASI KEUANGAN DAN ILMU PEMERINTAHAN');

$spreadsheet->getActiveSheet()->mergeCells('C2:F2');
$spreadsheet->getActiveSheet()->setCellValue('C2', 'SKT DITJEN POLPUM KEMENDAGRI NOMOR : 001-00-00/034/I/2019');

$spreadsheet->getActiveSheet()->mergeCells('C3:F3');
$spreadsheet->getActiveSheet()->setCellValue('C3', 'Sekretariat : Jln. Serdang Baru Raya No. 4B, Kemayoran - Jakarta Pusat 10650');

$spreadsheet->getActiveSheet()->mergeCells('C4:F4');
$spreadsheet->getActiveSheet()->setCellValue('C4', 'Website : www.lakip.co.id  E-mail : admin@lakip.co.id Telp./Fax. 021-42885718');

// BORDER BOTTOM
$spreadsheet->getActiveSheet()->getStyle('A5:F5')
  ->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK);

$spreadsheet->getActiveSheet()->mergeCells('A7:F7');
$spreadsheet->getActiveSheet()->setCellValue('A7', 'KWITANSI');

$spreadsheet->getActiveSheet()->mergeCells('A8:F8');
$spreadsheet->getActiveSheet()->setCellValue('A8', 'NO. :');

// Define named ranges
$spreadsheet->addNamedRange(new \PhpOffice\PhpSpreadsheet\NamedRange('Lembaga', $spreadsheet->getActiveSheet(), 'A1'));
$spreadsheet->addNamedRange(new \PhpOffice\PhpSpreadsheet\NamedRange('SKT', $spreadsheet->getActiveSheet(), 'C2'));
$spreadsheet->addNamedRange(new \PhpOffice\PhpSpreadsheet\NamedRange('Alamat', $spreadsheet->getActiveSheet(), 'C3'));
$spreadsheet->addNamedRange(new \PhpOffice\PhpSpreadsheet\NamedRange('Kontak', $spreadsheet->getActiveSheet(), 'C4'));
$spreadsheet->addNamedRange(new \PhpOffice\PhpSpreadsheet\NamedRange('Kwitansi', $spreadsheet->getActiveSheet(), 'A7'));
$spreadsheet->addNamedRange(new \PhpOffice\PhpSpreadsheet\NamedRange('No', $spreadsheet->getActiveSheet(), 'A8'));


// PROTEKSI
$spreadsheet->getActiveSheet()->getProtection()->setSheet(true);

// PROTEKSI DOCUMENT
// $security = $spreadsheet->getSecurity();
// $security->setLockWindows(true);
// $security->setLockStructure(true);
// $security->setWorkbookPassword("PhpSpreadsheet");
// PROTEKSI DOCUMENT
// $protection = $spreadsheet->getActiveSheet()->getProtection();
// $protection->setPassword('PhpSpreadsheet');
// $protection->setSheet(true);
// $protection->setSort(true);
// $protection->setInsertRows(true);
// $protection->setFormatCells(true);
// // UNPROTEKSI CELL
// $spreadsheet->getActiveSheet()->getStyle('A7')
//   ->getProtection()
//   ->setLocked(\PhpOffice\PhpSpreadsheet\Style\Protection::PROTECTION_UNPROTECTED);
// logo on worksheet
$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing(); //ok
$drawing->setName('Logo'); //ok
$drawing->setDescription('Lakip.co.id'); //ok
$drawing->setPath('img/lakip.png'); //ok
$drawing->setCoordinates('B1'); //ok
$drawing->setHeight(85); //ok
$drawing->setOffsetX(5); //ok
// $drawing->setRotation(25); 
// $drawing->getShadow()->setVisible(false);
// $drawing->getShadow()->setDirection(45);
$drawing->setWorksheet($spreadsheet->getActiveSheet('A1')); //ok


//Font Color
$spreadsheet->getActiveSheet()->getStyle('A10:F10')
  ->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE);
// Background color
$spreadsheet->getActiveSheet()->getStyle('A10:F10')->getFill()
  ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
  ->getStartColor()->setARGB('FFFF0000');
// Default width dan row
$spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(12);
$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);

// Generate an image
// $gdImage = @imagecreatetruecolor(120, 20) or die('Cannot Initialize new GD image stream');
// $textColor = imagecolorallocate($gdImage, 255, 255, 255);
// imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
// $drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
// $drawing->setName('Sample image');
// $drawing->setDescription('Sample image');
// $drawing->setImageResource($gdImage);
// $drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
// $drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
// $drawing->setHeight(36);
// $drawing->setWorksheet($spreadsheet->getActiveSheet());

$spreadsheet->getActiveSheet()->getSheetView()->setZoomScale(75);
$worksheet1 = $spreadsheet->createSheet();
$worksheet1->setTitle('LAKIP sheet');
$worksheet1->getTabColor()->setRGB('FF0000');
// $worksheet->getTabColor()->setRGB('FF0000');
// Header Tabel
$spreadsheet->setActiveSheetIndex(0)
  ->setCellValue('A10', 'NO')
  ->setCellValue('B10', 'NRP')
  ->setCellValue('C10', 'NAMA')
  ->setCellValue('D10', 'EMAIL')
  ->setCellValue('E10', 'JURUSAN')
  ->setCellValue('F10', 'GAMBAR');

$i = 11;
$no = 1;
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
$spreadsheet->getActiveSheet()->setTitle('Report LAKIP- ' . date('dmY Hi'));

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);

// Redirect output to a client’s web browser (Xlsx)
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="Report LAKIP XLSX.xlsx"');
header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
header('Cache-Control: max-age=1');

// CONTOH
header('Date: ' . gmdate('D, d M Y H:i:s \G\M\T', time()));
header('Last-Modified: ' . gmdate('D, d M Y H:i:s \G\M\T', time()));
header('Expires: ' . gmdate('D, d M Y H:i:s \G\M\T', time() + 3600));
// ATAU
header("Date: " . gmdate("D, d M Y H:i:s", time()) . " GMT");
header("Last-Modified: " . gmdate("D, d M Y H:i:s", time()) . " GMT");
header("Expires: " . gmdate("D, d M Y H:i:s", time() + 3600) . " GMT");
// AKHIR CONTOH

// If you're serving to IE over SSL, then the following may be needed
// header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
// header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
header('Pragma: public'); // HTTP/1.0

// SET AUTO FILTER
$spreadsheet->getActiveSheet()->setAutoFilter('A10:F10'); //OK

// logo header
// $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('test.xlsx');
// $worksheet = $spreadsheet->getActiveSheet();

// $drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\HeaderFooterDrawing();
// $drawing->setName('logo');
// $drawing->setPath('img/lakip.png');
// $drawing->setHeight(36);
// $spreadsheet->getActiveSheet()->getHeaderFooter()->addImage($drawing, \PhpOffice\PhpSpreadsheet\Worksheet\HeaderFooter::IMAGE_HEADER_LEFT);

// SET REPEAT 
$spreadsheet->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 5);
// SET PRINT AREA
$spreadsheet->getActiveSheet()->getPageSetup()->setPrintArea('A1:F28');
// BREAK ROW
// $spreadsheet->getActiveSheet()->setBreak('A20', \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet::BREAK_ROW); //OK

// BREAK COLUMN
// $spreadsheet->getActiveSheet()->setBreak('D10', \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet::BREAK_COLUMN); //OK
// GRIDLINES
$spreadsheet->getActiveSheet()->setShowGridlines(true);
// TTD
$spreadsheet->getActiveSheet()->setCellValue('E18', '');
$spreadsheet->getActiveSheet()->setCellValue('E19', '');
$spreadsheet->getActiveSheet()->setCellValue('A25', '');
$spreadsheet->getActiveSheet()->setCellValue('A25', '');


$spreadsheet->getActiveSheet()->mergeCells('E18:F18');
$spreadsheet->getActiveSheet()->setCellValue('E18', 'Jakarta, ' . date('d M Y'));
$spreadsheet->addNamedRange(new \PhpOffice\PhpSpreadsheet\NamedRange('dikeluarkan', $spreadsheet->getActiveSheet(), 'E18'));

$spreadsheet->getActiveSheet()->mergeCells('E19:F19');
$spreadsheet->getActiveSheet()->getStyle('E19:F19')
  ->getAlignment()->setWrapText(true);
$spreadsheet->getActiveSheet()->getStyle('E19:F19')
  ->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP);
$spreadsheet->getActiveSheet()->setCellValue('E19', 'Lembaga Administrasi Keuangan');
$spreadsheet->addNamedRange(new \PhpOffice\PhpSpreadsheet\NamedRange('instansi1', $spreadsheet->getActiveSheet(), 'E19'));

$spreadsheet->getActiveSheet()->mergeCells('E20:F20');
$spreadsheet->getActiveSheet()->getStyle('E20:F20')
  ->getAlignment()->setWrapText(true);
$spreadsheet->getActiveSheet()->getStyle('E20:F20')
  ->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP);
$spreadsheet->getActiveSheet()->setCellValue('E20', 'dan Ilmu Pemerintahan');
$spreadsheet->addNamedRange(new \PhpOffice\PhpSpreadsheet\NamedRange('instansi2', $spreadsheet->getActiveSheet(), 'E20'));



$spreadsheet->getActiveSheet()->mergeCells('E23:F23');
$spreadsheet->getActiveSheet()->setCellValue('E23', 'MASRIANTO');
$spreadsheet->addNamedRange(new \PhpOffice\PhpSpreadsheet\NamedRange('penerima', $spreadsheet->getActiveSheet(), 'E23'));

$spreadsheet->getActiveSheet()->mergeCells('E24:F24');
$spreadsheet->getActiveSheet()->setCellValue('E24', 'Bendahara');
$spreadsheet->addNamedRange(new \PhpOffice\PhpSpreadsheet\NamedRange('jabatan', $spreadsheet->getActiveSheet(), 'E24'));



// SET BACKGROUND
// $spreadsheet->getActiveSheet()->getStyle('A8:F11')
// ->getFill()->getStartColor()->setARGB('FFFF0000');
$spreadsheet->getActiveSheet()->mergeCells('A26:F28');
$spreadsheet->getActiveSheet()->getStyle('A26:F28')
  ->getAlignment()->setWrapText(true);
$spreadsheet->getActiveSheet()->getStyle('A26:F28')
  ->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP);
$richText = new \PhpOffice\PhpSpreadsheet\RichText\RichText();
$richText->createText('* Sesuai dengan ketentuan peraturan perundang-undangan yang berlaku,  ');
$payable = $richText->createTextRun('dokumen ini telah di tandatangani secara elektronik menggunakan sertifikat elektronik yang diterbitkan oleh lakip.co.id ');
$payable->getFont()->setBold(true);
$payable->getFont()->setItalic(true);
$payable->getFont()->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_DARKGREEN));
$richText->createText(', sehingga tidak diperlukan tandatangan dengan stempel basah.');
$spreadsheet->getActiveSheet()->getCell('A26')->setValue($richText);

// SET PAGE PRINT
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToWidth(1);
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToHeight(0);

$spreadsheet->getActiveSheet()->getPageMargins()->setTop(0.3);
$spreadsheet->getActiveSheet()->getPageMargins()->setRight(0.5);
$spreadsheet->getActiveSheet()->getPageMargins()->setLeft(0.5);
$spreadsheet->getActiveSheet()->getPageMargins()->setBottom(0.3);

$spreadsheet->getActiveSheet()->getPageSetup()->setHorizontalCentered(true);
$spreadsheet->getActiveSheet()->getPageSetup()->setVerticalCentered(false);

// $spreadsheet->getActiveSheet()->getPageSetup()
// ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE);
$spreadsheet->getActiveSheet()->getPageSetup()
  ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);


// $spreadsheet->getActiveSheet()->getHeaderFooter()
//   ->setOddHeader('&C&HPlease treat this document as confidential!');
$spreadsheet->getActiveSheet()->getHeaderFooter()
  // ->setOddFooter('&L&B' . $spreadsheet->getProperties()->getTitle() . '&RPage &P of &N'); //ok
  ->setOddFooter('&L&B' . $spreadsheet->getProperties()->getDescription() . '&RPage &P of &N');

/* Here there will be some code where you create $spreadsheet */
$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
$writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
$writer->save('php://output');
