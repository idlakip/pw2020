<?php
require_once 'vendor/autoload.php';

// Creating the new document...
// $phpWord = new \PhpOffice\PhpWord\PhpWord();

/* Note: any element you append to a document must reside inside of a Section. */



// teks 1
// Adding an empty Section to the document...
// $section = $phpWord->addSection();
// // Adding Text element to the Section having font styled by default...
// $section->addText(
//   '"Learn from yesterday, live for today, hope for tomorrow. '
//     . 'The important thing is not to stop questioning." '
//     . '(Albert Einstein)'
// );

// /*
//  * Note: it's possible to customize font style of the Text element you add in three ways:
//  * - inline;
//  * - using named font style (new font style object will be implicitly created);
//  * - using explicitly created font style object.
//  */


// teks 2
// // Adding Text element with font customized inline...
// $section->addText(
//   '"Great achievement is usually born of great sacrifice, '
//     . 'and is never the result of selfishness." '
//     . '(Napoleon Hill)',
//   array('name' => 'Tahoma', 'size' => 10)
// );


// teks 3
// // Adding Text element with font customized using named font style...
// $fontStyleName = 'oneUserDefinedStyle';
// $phpWord->addFontStyle(
//   $fontStyleName,
//   array('name' => 'Tahoma', 'size' => 10, 'color' => '1B2232', 'bold' => true)
// );
// $section->addText(
//   '"The greatest accomplishment is not in never falling, '
//     . 'but in rising again after you fall." '
//     . '(Vince Lombardi)',
//   $fontStyleName
// );



// teks 4
// // Adding Text element with font customized using explicitly created font style object...
// $fontStyle = new \PhpOffice\PhpWord\Style\Font();
// $fontStyle->setBold(true);
// $fontStyle->setName('Tahoma');
// $fontStyle->setSize(13);
// $myTextElement = $section->addText('"Believe you can and you\'re halfway there." (Theodor Roosevelt)');
// $myTextElement->setFontStyle($fontStyle);





// $properties = $phpWord->getDocInfo();
// $properties->setCreator('My name');
// $properties->setCompany('My factory');
// $properties->setTitle('My title');
// $properties->setDescription('My description');
// $properties->setCategory('My category');
// $properties->setLastModifiedBy('My name');
// $properties->setCreated(mktime(0, 0, 0, 3, 12, 2014));
// $properties->setModified(mktime(0, 0, 0, 3, 14, 2014));
// $properties->setSubject('My subject');
// $properties->setKeywords('my, key, word');

// $phpWord->setDefaultFontName('Times New Roman');
// $phpWord->setDefaultFontSize(12);

// $phpWord->getSettings()->setZoom(75);
// $phpWord->getSettings()->setZoom(Zoom::BEST_FIT);

// $phpWord->getSettings()->setDecimalSymbol(',');

// $header = $section->addHeader();

// $phpWord = new \PhpOffice\PhpWord\PhpWord();

// // New portrait section
// $section = $phpWord->addSection();
// $textRun = $section->addTextRun();

// $text = $textRun->addText('Hello World! Time to ');

// $text = $textRun->addText('wake ', array('bold' => true));
// $text->setChangeInfo(TrackChange::INSERTED, 'Fred', time() - 1800);

// $text = $textRun->addText('up');
// $text->setTrackChange(new TrackChange(TrackChange::INSERTED, 'Fred'));

// $text = $textRun->addText('go to sleep');
// $text->setChangeInfo(TrackChange::DELETED, 'Barney', new \DateTime('@' . (time() - 3600)));

// $section->addPageBreak();
// $footer = $section->addFooter();
// $footer->addPreserveText('Page {PAGE} of {NUMPAGES}.');


// $phpWord = new \PhpOffice\PhpWord\PhpWord();
// $section = $phpWord->createSection();
// $section->addText('Hello World!');
// $file = 'kwitansi.docx';
// header("Content-Description: File Transfer");
// header('Content-Disposition: attachment; filename="' . $file . '"');
// header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
// header('Content-Transfer-Encoding: binary');
// header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
// header('Expires: 0');
// $xmlWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
// $xmlWriter->save("php://output");


$objReader = \PhpOffice\PhpWord\IOFactory::createReader('Word2007');
$phpWord = $objReader->load("helloWorld.doc");

// $rendedererName = \PhpOffice\PhpWord\Settings::PDF_RENDERER_DOMPDF;
// $rendedererName = \PhpOffice\PhpWord\Settings::PDF_RENDERER_MPDF;
$rendedererName = \PhpOffice\PhpWord\Settings::PDF_RENDERER_TCPDF;
$rendedererLibrary = 'mpdf'; // 'tcpdf'
$rendedererLibraryPath = '' . $rendedererLibrary;
if (!\PhpOffice\PhpWord\Settings::setPdfRenderer(
  $rendedererName,
  $rendedererLibraryPath
)) {
  die('NOTICE: Please set the $rendedererName and $rendedererLibraryPath values' .
    '<br />' .
    'at the top of this script as appropriate for your directory structur');
}
$rendedererLibraryPath = '' . $rendedererLibrary;





// Saving the document as OOXML file...
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'PDF');
$objWriter->save('kwitansi.pdf');


// $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
// $objWriter->save('kwitansi.docx');

// Saving the document as ODF file...
// $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'ODText');
// $objWriter->save('kwitansi.odt');

// Saving the document as HTML file...
// $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'HTML');
// $objWriter->save('kwitansi.html');

/* Note: we skip RTF, because it's not XML-based and requires a different example. */
/* Note: we skip PDF, because "HTML-to-PDF" approach is used to create PDF documents. */
