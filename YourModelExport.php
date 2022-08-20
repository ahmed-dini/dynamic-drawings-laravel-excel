<?php
// Made by Ahmed Al Dini -> https://github.com/ahmed-dini
// example class 
namespace App\Exports;

use App\Models\YourModel;

class YourModelExport extends ExportBase {

    protected $row_height = 50;
    protected $column_width = 10;

    public function query() {
        // add all your query as you'll normally would 
        return YourModel::all();

        // example row output
        // return [
        //     [
        //         'id' => 1,
        //         'name' => 'Ahmed Al Dini',
        //         'age' => 27,
        //         'avatar' => 'your/avater/path.jpg',
        //     ]
        // ]);
    }

    public function map($row): array {

        $this->add_column('id', $row['id']);
        $this->add_column('name', $row['name']);
        $this->add_column('age', $row['age']);
        $this->add_column('avatar', $this->draw($row['avatar']));

        return $this->get_row();
    }
}
