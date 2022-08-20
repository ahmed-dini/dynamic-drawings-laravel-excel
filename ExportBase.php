<?php
// Made by Ahmed Al Dini -> https://github.com/ahmed-dini
namespace App\Exports;

use Maatwebsite\Excel\Concerns\FromQuery;
use Maatwebsite\Excel\Concerns\WithMapping;
use Maatwebsite\Excel\Concerns\WithDrawings;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Concerns\RegistersEventListeners;
use Maatwebsite\Excel\Events\BeforeExport;
use Maatwebsite\Excel\Events\BeforeWriting;
use Maatwebsite\Excel\Events\BeforeSheet;
use Maatwebsite\Excel\Events\AfterSheet;

class ExportBase implements FromQuery, WithMapping, WithDrawings, withEvents {
    
    use RegistersEventListeners;

    protected $extra_fields = [];
    protected $addHeading = true;
    protected $drawings = [];
    protected $columns = [];
    protected $row_index = 2;
    protected $row_values = [];
    protected $columns_count = 0;
    protected $row_height = 50;
    protected $column_width = 10;
    
    public function query() {}
    public function map($row): array {}

    public function get_heading_row() {
        return array_merge($this->columns, array_keys($this->extra_fields));
    }

    public function get_row_height() {
        return $this->row_height;   
    }
    
    public function set_row_height($height) {
        return $this->row_height = $height;   
    }
    
    public function get_column_width() {
        return $this->column_width;   
    }
    
    public function set_column_width($width) {
        return $this->column_width = $width;   
    }

    public function set_columns_count() {
        // latest column index is the columns count 
        $this->columns_count = $this->column_index;
    }
    
    public function get_columns_count() {
        return $this->columns_count;
    }

    public function set_column_index() {
        $this->column_index = count($this->columns);
    }

    public function add_column($key, $value) {
        $this->extra_fields[$key] = $value;
        $this->column_index++;
    }

    public function get_column_index() {
        return $this->column_index;
    }

    public function add_row() {
        $this->row_index++;
    }

    public function get_row_index() {
        return $this->row_index;
    }

    public function set_row_values($data = []) {
        $this->row_values = $data;
    }

    public function get_row_values($data = []) {
        $this->add_row();
        return array_merge($this->row_values, $this->extra_fields);
    }

    public function get_row() {
        // set columns count and reset the columns counter for each map loop
        $this->set_columns_count();
        $this->set_column_index();

        if ($this->addHeading) {
            $this->addHeading = false;

            // return two rows 1 = heading, 2 = row data
            return [ 
                $this->get_heading_row(),
                $this->get_row_values()
            ];
            
        } else {
            // return 1 row data 
            return $this->get_row_values();
        }
    }

    // returns the column letter based on column index 
    public function getLetterFromNumber($column_index) {
        $numeric = $column_index % 26;
        $letter = chr(65 + $numeric);
        $num2 = intval($column_index / 26);
        if ($num2 > 0) {
            return $this->getLetterFromNumber($num2 - 1) . $letter;
        } else {
            return $letter;
        }
    }

    public function drawings() {
        return $this->drawings;
    }

    public function draw($image_path, $coordinates = false) {
        if ($image_path) {
            $drawing = new Drawing();
            $drawing->setName(basename($image_path));
            $drawing->setDescription(basename($image_path));
            $drawing->setPath(public_path($image_path));
            $drawing->setHeight(50);

            $coordinates = $coordinates ? $coordinates : $this->get_current_cell_coordinates();
                
            $drawing->setCoordinates($coordinates);
    
            array_push($this->drawings, $drawing);
        }

        return '';
    }

    public function get_current_cell_coordinates() {
        return $this->getLetterFromNumber($this->get_column_index()).$this->get_row_index();
    }

    public function registerEvents(): array
    {
        return [
            BeforeSheet::class => function(BeforeSheet $event) {
                // set values before query starts 
                $this->set_column_index();
            },
            
            AfterSheet::class => function(AfterSheet $event) {
                
                // apply styling to rows or columns 
                for ($row_index = 1; $row_index < $this->get_row_index(); $row_index++) {
                    $event->sheet->getDelegate()->getRowDimension($row_index)->setRowHeight($this->get_row_height());
                }

                for ($column_index = 0; $column_index < $this->get_columns_count(); $column_index++) {
                    $event->sheet->getDelegate()->getColumnDimension($this->getLetterFromNumber($column_index))->setWidth($this->get_column_width());
                }
     
            },

            beforeWriting::class => function(beforeWriting $event) {
                // excecutes after mapping through all rows 

                $spread_sheet = $event->writer->getDelegate();
                
                // adding the prevously collected drawings to sheet 
                foreach ($event->getConcernable()->drawings as $key => $drawing) {
                    $drawing->setWorksheet($spread_sheet->getActiveSheet());
                }
                
                // to solve download unreadable file 
                ob_end_clean();
            },
        ];
    }
}
