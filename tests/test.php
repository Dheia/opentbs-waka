<?php 
require_once __DIR__ . '/../vendor/autoload.php'; // Autoload files using Composer autoload
use Waka\OpenTBS\MergePpt;

$test = new MergePpt();
$test->loadTemplate('tests/demo_ms_powerpoint3.pptx');
//echo $test->degubTemplate();
//echo $test->loadTemplate('tests/demo_ms_powerpoint3.pptx')->degubTemplate();
$baseData = [
    'yourname' => 'Charles',
    'image' => 'tests/test.jpg',
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
            'Cat. E' => 5,
            'Cat. G' => 5,
        ],
    ],
    2 => [
        'label' => 'CA n-1',
        'datas' => [
            'Cat. A' => 0.5,
            'Cat. B' => 1.1,
            'Cat. C' => 3.0,
            'Cat. D' => 5,
            'Cat. E' => 5,
            'Cat. G' => 5,
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
$dataBlock = array();
$dataBlock[] = array('rank'=> 'A', 'image' => 'tests/test.jpg', 'firstname'=>'Sandra' , 'name'=>'Hill'      , 'number'=>'1523d', 'score'=>200, 'email_1'=>'sh@tbs.com',  'email_2'=>'sandra@tbs.com',  'email_3'=>'s.hill@tbs.com');
$dataBlock[] = array('rank'=> 'A', 'image' => 'tests/test.jpg', 'firstname'=>'Roger'  , 'name'=>'Smith'     , 'number'=>'1234f', 'score'=>800, 'email_1'=>'rs@tbs.com',  'email_2'=>'robert@tbs.com',  'email_3'=>'r.smith@tbs.com' );
$dataBlock[] = array('rank'=> 'B', 'image' => 'tests/test.jpg', 'firstname'=>'William', 'name'=>'Mac Dowell', 'number'=>'5491y', 'score'=>130, 'email_1'=>'wmc@tbs.com', 'email_2'=>'william@tbs.com', 'email_3'=>'w.m.dowell@tbs.com' );
//
$test->mergeField(1,$baseData);
//
$test->changePicture(3, '#merge_me#', 'tests/test.jpg');
//
$test->changeChart('chart1', $chartData);
$test->changeChart('chart2', $chartData, true);
// //
$test->MergeBlock(6,'a', $dataBlock);
$test->deleteSlide(2);
//echo $test->degubTemplate();
echo $test->savePpt("tests/salut.pptx");

