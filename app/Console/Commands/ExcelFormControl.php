<?php

namespace App\Console\Commands;

use Illuminate\Console\Command;

use Maatwebsite\Excel\Facades\Excel;

class ExcelFormControl extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'z:excel-form-control';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'test to access form control in excel';

    /**
     * Create a new command instance.
     *
     * @return void
     */
    public function __construct()
    {
        parent::__construct();
    }

    /**
     * Execute the console command.
     *
     * @return mixed
     */
    public function handle()
    {
        // $inputFileName = './storage/excel/Book2.xlsx';
        $inputFileName = './storage/excel/Book2-2.xlsx';
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($inputFileName);

        $vmlDrawings = $spreadsheet->getUnparsedLoadedData()['sheets']['Sheet1']['vmlDrawings'];
        var_dump(array_values($vmlDrawings)[0]['content']);


        $ctrlProps = $spreadsheet->getUnparsedLoadedData()['sheets']['Sheet1']['ctrlProps'];
var_dump($ctrlProps);
    }
}

