<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Conditional;

$products = [
    [
        'category' => 'Умные кольца',
        'article' => 'UK-001',
        'name' => 'Кольцо-трекер пульса IWant SoReally',
        'unit' => 'Штука',
        'stock' => 45,
        'delivery_date' => '2025-01-14'
    ],
    [
        'category' => 'Умные кольца',
        'article' => 'UK-002',
        'name' => 'NFC кольцо ToWork',
        'unit' => 'Штука',
        'stock' => 3,
        'delivery_date' => '2025-02-21'
    ],
    [
        'category' => 'Мини-проекторы',
        'article' => 'MP-101',
        'name' => 'Мини-проектор WithYou',
        'unit' => 'Штука',
        'stock' => 18,
        'delivery_date' => '2025-03-11'
    ],
    [
        'category' => 'Мини-проекторы',
        'article' => 'MP-102',
        'name' => 'Мини-проектор для смартфона ItsMy Real',
        'unit' => 'Штука',
        'stock' => 150,
        'delivery_date' => '2025-04-01'
    ],
    [
        'category' => 'Умные вещи',
        'article' => 'UV-201',
        'name' => 'Умные часы Dream 3',
        'unit' => 'Штука',
        'stock' => 120,
        'delivery_date' => '2025-05-15'
    ],
    [
        'category' => 'Умные вещи',
        'article' => 'UV-202',
        'name' => 'Фитнес-браслет ThisIs 2',
        'unit' => 'Штука',
        'stock' => 85,
        'delivery_date' => '2025-06-25'
    ],
    [
        'category' => 'Гаджеты для здоровья',
        'article' => 'GH-301',
        'name' => 'Датчик шагомер NotA',
        'unit' => 'Штука',
        'stock' => 0,
        'delivery_date' => '2025-07-13'
    ],
    [
        'category' => 'Гаджеты для здоровья',
        'article' => 'GH-302',
        'name' => 'Умная зубная щетка Joke XD',
        'unit' => 'Штука',
        'stock' => 200,
        'delivery_date' => '2025-08-28'
    ],
];

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('Ассортимент');

$currentDateTime = date('d.m.Y в H:i:s');
$sheet->setCellValue('G1', "Дата формирования: {$currentDateTime}");
$sheet->getStyle('G1')->getFont()->setSize(11);
$sheet->getStyle('G1')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);

$sheet->setCellValue('A2', '');

$sheet->setCellValue('A3', 'Отчет по остаткам на складе');
$sheet->getStyle('A3')->getFont()->setBold(true)->setSize(14);
$sheet->getStyle('A3')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
$sheet->mergeCells('A3:G3');

$sheet->setCellValue('A4', 'Содержит ключевую информацию о номенклатуре для управления запасами');
$sheet->getStyle('A4')->getFont()->setSize(11);
$sheet->getStyle('A4')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
$sheet->mergeCells('A4:G4');

$sheet->setCellValue('A5', '');
$sheet->setCellValue('A6', '');

$headerRow = 7;
$headers = ['№ п/п', 'Категория', 'Артикул', 'Наименование материалов, изделий, конструкций и оборудования', 'Ед. изм.', 'Остаток', 'Срок поставки'];

foreach ($headers as $colIndex => $header) {
    $colLetter = chr(65 + $colIndex);
    $sheet->setCellValue($colLetter . $headerRow, $header);
}

$headerStyle = [
    'font' => ['bold' => true, 'size' => 11, 'color' => ['argb' => 'FF000000']],
    'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER, 'wrapText' => true],
    'borders' => [
        'allBorders' => ['borderStyle' => Border::BORDER_THIN]
    ],
    'fill' => [
        'fillType' => Fill::FILL_SOLID, 
        'startColor' => ['argb' => 'FF99CCFF']
    ]
];
$sheet->getStyle("A{$headerRow}:G{$headerRow}")->applyFromArray($headerStyle);

$dataRow = $headerRow + 1;
foreach ($products as $index => $product) {
    $row = $dataRow + $index;
    
    $sheet->setCellValue("A{$row}", $index + 1);
    $sheet->setCellValue("B{$row}", $product['category']);
    $sheet->setCellValue("C{$row}", $product['article']);
    $sheet->setCellValue("D{$row}", $product['name']);
    $sheet->setCellValue("E{$row}", $product['unit']);
    $sheet->setCellValue("F{$row}", $product['stock']);
    $sheet->setCellValue("G{$row}", $product['delivery_date']);
}

$conditionalRed = new Conditional();
$conditionalRed->setConditionType(Conditional::CONDITION_CELLIS);
$conditionalRed->setOperatorType(Conditional::OPERATOR_LESSTHAN);
$conditionalRed->addCondition(10);
$conditionalRed->getStyle()->getFill()->setFillType(Fill::FILL_SOLID)->getEndColor()->setARGB('FFFFCCCC');

$conditionalGreen = new Conditional();
$conditionalGreen->setConditionType(Conditional::CONDITION_CELLIS);
$conditionalGreen->setOperatorType(Conditional::OPERATOR_GREATERTHAN);
$conditionalGreen->addCondition(100);
$conditionalGreen->getStyle()->getFill()->setFillType(Fill::FILL_SOLID)->getEndColor()->setARGB('FFCCFFCC');

$lastDataRow = $dataRow + count($products) - 1;
$sheet->getStyle("F{$dataRow}:F{$lastDataRow}")->setConditionalStyles([$conditionalRed, $conditionalGreen]);

$sheet->setAutoFilter("A{$headerRow}:G{$lastDataRow}");

foreach (range('A', 'G') as $col) {
    $sheet->getColumnDimension($col)->setAutoSize(true);
}

$sheet->getColumnDimension('A')->setWidth(8);
$sheet->getColumnDimension('B')->setWidth(15);
$sheet->getColumnDimension('C')->setWidth(12);
$sheet->getColumnDimension('D')->setWidth(50);
$sheet->getColumnDimension('E')->setWidth(10);
$sheet->getColumnDimension('F')->setWidth(12);
$sheet->getColumnDimension('G')->setWidth(15);

$tableStyle = [
    'borders' => [
        'allBorders' => [
            'borderStyle' => Border::BORDER_THIN,
            'color' => ['argb' => 'FF000000'],
        ],
    ],
    'alignment' => [
        'vertical' => Alignment::VERTICAL_CENTER,
    ],
];

$sheet->getStyle("A{$dataRow}:G{$lastDataRow}")->applyFromArray($tableStyle);

$sheet->getStyle("A{$dataRow}:A{$lastDataRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
$sheet->getStyle("B{$dataRow}:B{$lastDataRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
$sheet->getStyle("C{$dataRow}:C{$lastDataRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
$sheet->getStyle("E{$dataRow}:E{$lastDataRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
$sheet->getStyle("F{$dataRow}:F{$lastDataRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
$sheet->getStyle("G{$dataRow}:G{$lastDataRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

$sheet->getStyle("A{$headerRow}:G{$lastDataRow}")->getAlignment()->setWrapText(true);

$sheet->getStyle("G{$dataRow}:G{$lastDataRow}")
    ->getNumberFormat()
    ->setFormatCode('DD.MM.YYYY');

$outerBorderStyle = [
    'borders' => [
        'outline' => [
            'borderStyle' => Border::BORDER_MEDIUM,
            'color' => ['argb' => 'FF000000'],
        ],
    ],
];
$sheet->getStyle("A{$headerRow}:G{$lastDataRow}")->applyFromArray($outerBorderStyle);

for ($row = $headerRow; $row <= $lastDataRow; $row++) {
    $sheet->getRowDimension($row)->setRowHeight(25);
}

$writer = new Xlsx($spreadsheet);
$filename = 'xlsxreportgenerated.xlsx';
$writer->save($filename);

echo "Файл создан: {$filename}";