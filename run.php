<?php

require_once 'vendor/autoload.php';

use Fairy\MhtFileMaker;

// 1、保存为文件
MhtFileMaker::getInstance()
    ->addFile('index.html')
    ->completeImg('https://www.php.net', true)
    ->makeFile(__DIR__.'/a.doc');

// 2、浏览器下载
MhtFileMaker::getInstance()
    ->addFile('index.html')
    ->completeImg('https://www.php.net')
    ->download();