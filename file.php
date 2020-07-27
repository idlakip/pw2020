<?php
require_once 'vendor/autoload.php';
require 'functions.php';
// Creating the new document...
$phpWord = new \PhpOffice\PhpWord\PhpWord();

/* Note: any element you append to a document must reside inside of a Section. */

// Adding an empty Section to the document...
$section = $phpWord->addSection();
// HEADER
$header = $section->addHeader();
$header->addImage('img/bca.png');

// $textrun->addImage('https://upload.wikimedia.org/wikipedia/commons/e/eb/Intel-logo.jpg');
// $source = file_get_contents('img/lakip.png');
// $textrun->addImage($source);

$file = 'test.docx';
header("Content-Description: File Transfer");
header('Content-Disposition: attachment; filename="' . $file . '"');
header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
header('Content-Transfer-Encoding: binary');
header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
header('Expires: 0');


// PAGE NUMBER 
// Method 1
// $section = $phpWord->addSection(array('pageNumberingStart' => 1));

// FOOTER
$footer = $section->addFooter();
$footer->addPreserveText('Page {PAGE} of {NUMPAGES}.');
$xmlWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
$xmlWriter->save("php://output");
