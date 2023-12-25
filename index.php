<?php

require_once __DIR__ . '/phpQuery/phpQuery.php';
require_once __DIR__ . '/PHPExcel/Classes/PHPExcel.php';
require_once __DIR__ . '/PHPExcel/Classes/PHPExcel/Writer/Excel2007.php';

// require_once dirname(__FILE__) . '/PHPExcel/Classes/PHPExcel.php';

function debug($item)
{
    echo '<pre>';
    echo print_r($item);
    echo '</pre>';
}

/* для указания кодировки utf-8 */
header('Content-type: text/html; charset=utf-8');
setlocale(LC_ALL, 'ru_RU.UTF-8');


/* для вывода ошибок */
ini_set('error_reporting', E_ALL);
ini_set('display_errors', 1);
ini_set('display_startup_errors', 1);
ini_set('max_execution_time', 900);

// Удаление строки
function deleteString($str)
{
    $pattern = [
        '/<a class="b-submit js-dialog" href="#dialog">Узнать цену<\/a>/',
        '/<a[^>]*class="back"[^>]*>.*?<\/a>/'
    ];
    
    $result = preg_replace($pattern, '', $str);

    return $result;
}


function parser($url)
{
    $ch = curl_init($url);
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
    curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, false);
    curl_setopt($ch, CURLOPT_HEADER, false);
    $result = curl_exec($ch);
    curl_close($ch);

    return $result;
}

$urls = [
    'https://atex-tools.ru/catalog/list/17'
];

$arrMainParamsq = [];
$count_id = 1;

foreach ($urls as $url) {
    $count_id++;
    $result = parser($url);
    $pq = phpQuery::newDocument($result);

    $namecat = $pq->find(".name-page h1")->text();

    $arrLinks = [];
    $listParamsProduct = $pq->find(".b-catalogue__item .b-catalogue__item__link");
    foreach ($listParamsProduct as $link) {
        $elemLink = pq($link);
        $arrLinks[] = "https://atex-tools.ru" . $elemLink->attr("href");
    }

    foreach ($arrLinks as $link) {
        $count_id++;
        $result = parser($link);
        $pq = phpQuery::newDocument($result);

        $arrMainParams = [
            "url" => $link,
            'name' => $pq->find("h1")->text(),
            'description' => trim($pq->find(".b-product__image-area__txt")->html()),
            'table' => deleteString(trim($pq->find(".b-product__table-area")->html())),
            'img' => 'https://atex-tools.ru/' . $pq->find(".b-product__image-wrapper img")->attr("src"),
        ];

        $arrMainParamsq[] = [
            'id' => $count_id,
            'cat' => "КАТЕГОРИЯ > $namecat",
            'listMainParams' => $arrMainParams,
        ];
    }
}

$xls = new PHPExcel();

$xls->setActiveSheetIndex(0);
$sheet = $xls->getActiveSheet();
$sheet->setTitle('lisproduct');

$sheet->setCellValue("A1", "ID");
$sheet->setCellValue("G1", "cat");
$sheet->setCellValue("B1", "name");
$sheet->setCellValue("C1", "description");
$sheet->setCellValue("D1", "table");
$sheet->setCellValue("E1", "img");
$sheet->setCellValue("F1", "url");

foreach($arrMainParamsq as $key=>$product) {
    $index = $key + 2;

    $productParam = $product['listMainParams'];

    $sheet->setCellValue("A".$index, $product['id']);
    $sheet->setCellValue("G".$index, $product["cat"]);
    $sheet->setCellValue("B".$index, $productParam["name"]);
    $sheet->setCellValue("C".$index, $productParam["description"]);
    $sheet->setCellValue("D".$index, $productParam["table"]);
    $sheet->setCellValue("E".$index, $productParam["img"]);
    $sheet->setCellValue("F".$index, $productParam["url"]);
}

$objWriter = new PHPExcel_Writer_Excel2007($xls);
$filePath = __DIR__ . "/file_catalog.xlsx";
$objWriter->save($filePath);