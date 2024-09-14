# Spout

## About

Spout is a PHP library to read and write spreadsheet files (CSV, XLSX and ODS), in a fast and scalable way.
Unlike other file readers or writers, it is capable of processing very large files, while keeping the memory usage really low (less than 3MB).

Join the community and come discuss Spout: [![Gitter](https://badges.gitter.im/Join%20Chat.svg)](https://gitter.im/box/spout?utm_source=badge&utm_medium=badge&utm_campaign=pr-badge)


## Documentation

Full documentation can be found at [https://opensource.box.com/spout/](https://opensource.box.com/spout/).

## Requirements

* PHP version 7.2 or higher
* PHP extension `php_zip` enabled
* PHP extension `php_xmlreader` enabled

## Example


```php

$pathToFile = __DIR__.'/excel.xlsx';

$data = [
   ['Данные юр. лица', '', '', 'Купленный товар', '', '', ''],
   ['Наименование', 'ИНН', 'Адрес', 'Номенклатура', 'Количество', 'Цена, руб.', 'Дата оплаты', 'Сумма'],
   ['ИП Кошкин', '00111122311', 'Москва ул. Ленина д.5', 'Принтер', 2, 100, '01.01.2024', '=ROUND(E3*F3,2)'],
   ['ИП Собакин', '00111122421', 'Москва ул. Ленина д.6', 'Принтер', 3, 200, '02.01.2024', '=ROUND(E4*F4,2)'],
   ['ИП Конов', '00111122533', 'Москва ул. Ленина д.7', 'Принтер', 1, 300, '03.01.2024', '=ROUND(E5*F5,2)'],
   ['ИП Медведев', '00111122644', 'Москва ул. Ленина д.8', 'Принтер', 5, 400, '04.01.2024', '=ROUND(E6*F6,2)'],
   [null, null, null, null, null, null, null, '=SUM(H3:H6)'],
];

$border = (new BorderBuilder())
   ->setBorderBottom(Color::BLACK, Border::WIDTH_THIN, Border::STYLE_SOLID)
   ->setBorderTop(Color::BLACK, Border::WIDTH_THIN, Border::STYLE_SOLID)
   ->setBorderLeft(Color::BLACK, Border::WIDTH_THIN, Border::STYLE_SOLID)
   ->setBorderRight(Color::BLACK, Border::WIDTH_THIN, Border::STYLE_SOLID)
   ->build();

$defaultStyle = (new StyleBuilder())
   ->setFontSize(11)
   ->setFontName('Calibri')
   ->setShouldWrapText()
   ->setCellAlignment(CellAlignment::CENTER)
   ->setCellVerticalAlignment(CellVerticalAlignment::MIDDLE)
   ->setBorder($border)
   ->build();

$writer = WriterEntityFactory::createXLSXWriter();
$writer->setDefaultRowStyle($defaultStyle);
$writer->setFreezePane('A3');
$writer->openToFile($pathToFile);

$writer->getCurrentSheet()->setAutoFilter('A2:H6');

$row = 0;
foreach ($data as $item){
   $cells = []; $col=0;
   foreach ($item as $v) {

       if ($col < 3) $color = 'FBFCEB';
       elseif ($col < 6) $color = 'D4DFED';
       else $color = 'DDD7E7';

       $style =
           (new StyleBuilder())
           ->setBackgroundColor($color)
           ->build();

       if ($row==0){
           $style = null;
       }


       if ($style){
           if (in_array($col, [0,2,3]) && $row>1){
               $style->setCellAlignment(CellAlignment::LEFT);
           }

           if (in_array($col, [5,7])) {
               $style->setFormat('0.00');
           }
           elseif (in_array($col, [4])) {
               $style->setFormat('0');
           }
           elseif (in_array($col, [1,6])) {
               $style->setFormat('@');
           }

           if ($row==6) {
               $style->setBackgroundColor(null);
           }

       }

       $cells[] = WriterEntityFactory::createCell($v, $style);
       $col++;
   }

   $singleRow = WriterEntityFactory::createRow($cells);
   $writer->addRow($singleRow);
   $row++;
}


$writer->getCurrentSheet()->mergeCells('A1:C1');
$writer->getCurrentSheet()->mergeCells('D1:F1');

$writer->getCurrentSheet()->addColumnDimension(new ColumnDimension('A', 18));
$writer->getCurrentSheet()->addColumnDimension(new ColumnDimension('B', 18));
$writer->getCurrentSheet()->addColumnDimension(new ColumnDimension('C', 25));
$writer->getCurrentSheet()->addColumnDimension(new ColumnDimension('D', 18));
$writer->getCurrentSheet()->addColumnDimension(new ColumnDimension('E', 18));
$writer->getCurrentSheet()->addColumnDimension(new ColumnDimension('F', 18));
$writer->getCurrentSheet()->addColumnDimension(new ColumnDimension('G', 18));
$writer->getCurrentSheet()->addColumnDimension(new ColumnDimension('H', 18));

$writer->close();
```

## Copyright and License

Copyright 2022 Box, Inc. All rights reserved.

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

   http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
