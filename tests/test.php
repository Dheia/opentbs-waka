<?php 
require_once __DIR__ . '/../vendor/autoload.php'; // Autoload files using Composer autoload
use Waka\OpenTBS\MergePpt;

$test = new MergePpt();
$test->loadTemplate('tests/demo_ms_powerpoint3.pptx');
//echo $test->loadTemplate('tests/demo_ms_powerpoint3.pptx')->degubTemplate();
$baseData = [
    'yourname' => 'Charles',
    'sub' => [
        'test' => 'test',
        ],
    ];
$chartData = [
    1 => [
        'label' => 'CA',
        'datas' => [
            'Cat. A' => 0.7,
            'Cat. B' => 1.0,
            'Cat. C' => 3.2,
            'Cat. D' => 4.8,
        ],
    ],
    2 => [
        'label' => 'CA n-1',
        'datas' => [
            'Cat. A' => 0.5,
            'Cat. B' => 1.1,
            'Cat. C' => 3.0,
            'Cat. D' => 5,
        ],
    ],
    3 => [
        'label' => 'Autre',
        'datas' => [
            'Cat. A' => 5,
            'Cat. B' => 4,
            'Cat. C' => 5,
            'Cat. D' => 4,
        ],
    ],
];
$test->setVars($baseData);
$test->changePicture(2, '#merge_me#', 'tests/test.jpg');
$test->changeChart('chart1', $chartData);
$test->changeChart('chart2', $chartData, true);
$test->deleteSlide(3);
//echo $test->degubTemplate();
echo $test->savePpt("tests/salut.pptx");

