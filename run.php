<?php

require_once 'vendor/autoload.php';

use Fairy\Html2WordMaker;
use Fairy\MhtFileMaker;

// 1、保存为文件
Html2WordMaker::getInstance()
    ->addFile('resource/tpl.html')
    ->eraseLink()
    ->fetchImg('http://php.test/html2word')
    ->makeFile('a.doc');

 MhtFileMaker::getInstance()
     ->addFile('resource/tpl.html')
     ->eraseLink()
     ->fetchImg('http://php.test/html2word')
     ->makeFile('a.doc');

 //2、浏览器下载
 Html2WordMaker::getInstance()
     ->addFile('resource/tpl.html')
     ->fetchImg('http://php.test/html2word')
     ->download();

 MhtFileMaker::getInstance()
     ->addFile('resource/tpl.html')
     ->fetchImg('http://php.test/html2word')
     ->download();
