<?php

require_once 'vendor/autoload.php';

use Fairy\MhtFileMaker;

// 1、保存为文件
MhtFileMaker::getInstance()
    ->addFile('resource/tpl.html')
    ->eraseLink()
    ->fetchImg('http://php.test/html2word')
    ->makeFile('resource/a.doc');

// 2、浏览器下载
MhtFileMaker::getInstance()
    ->addFile('resource/tpl.html')
    ->fetchImg('http://php.test/html2word')
    ->download();