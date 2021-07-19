<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class PhpSpreadSheetController extends Controller
{
    public function index() {

        // (1) Calculating value from excel 

        // $spreadsheet = IOFactory::load('assets/test.xlsx');

        // $somevalue = 50;

        // $spreadsheet->getActiveSheet()->setCellValue('B2', $somevalue)
        //                              ->setCellValue('B3', 20)
        //                              ->setCellValue('B4', 60);

        // echo '<strong>Excel to View:</strong>';
        // echo $spreadsheet->getActiveSheet()->getCell('B5')->getCalculatedValue();

        // (2) Read value from excel 

        // $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
        // $reader->setReadDataOnly(TRUE);
        // $spreadsheet = $reader->load("assets/test1.xlsx");

        // $worksheet = $spreadsheet->getActiveSheet();

        // echo '<table>' . PHP_EOL;
        // foreach ($worksheet->getRowIterator() as $row) {
        //     echo '<tr>' . PHP_EOL;
        //     $cellIterator = $row->getCellIterator();
        //     $cellIterator->setIterateOnlyExistingCells(FALSE);
        
        // foreach ($cellIterator as $cell) {
        //     echo '<td>' .
        //          $cell->getValue() .
        //          '</td>' . PHP_EOL;
        // };
        //     echo '</tr>' . PHP_EOL;
        // }
        // echo '</table>' . PHP_EOL;

        // (3) Write Hello World into Excel Sheet

        // $spreadsheet = new Spreadsheet();
        // $sheet = $spreadsheet->getActiveSheet();
        // $sheet->setCellValue('A1', 'Hello World !');

        // $writer = new Xlsx($spreadsheet);
        // $writer->save('hello-world.xlsx');

        // (4) Array Value into into Excel Sheet


            $spreadsheet = new Spreadsheet();
            $arrayData = $spreadsheet->getActiveSheet();
            $arrayData = [
                [NULL, 2010, 2011, 2012],
                ['Q1',   12,   15,   21],
                ['Q2',   56,   73,   86],
                ['Q3',   52,   61,   69],
                ['Q4',   30,   32,    0],
            ];

            $spreadsheet->getActiveSheet()
            ->fromArray(
                $arrayData,  // The data to set
                NULL,        // Array values with this value will not be set
                'C3'         // Top left coordinate of the worksheet range where
                             //    we want to set these values (default is A1)
            );
            $writer = new Xlsx($spreadsheet);
            $writer->save('testArray.xlsx');


    }
}
