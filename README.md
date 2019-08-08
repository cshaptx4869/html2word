HTML2WORD
=======
Installation
------------

```bash
composer require cshaptx4869/html2word
```

Example
-------

###### index.html

```html
<!doctype html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport"
          content="width=device-width, user-scalable=no, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <title>Document</title>
    <style>
        .testclass {
            font-size: 30px;
            color: #ff0000;
        }
    </style>
</head>
<body>
    <img src="images/logos/php-logo.svg" width="50px">
    <h1>中文的标题，技术无止境，一直在路上</h1>
    <p>p是可以分段的. 使用PHP将html转word</p>
    <p>再分一段 使用PHP将html转word</p>
    <p>还分一段，下面加个图片</p>
    <a href="https://www.baidu.com">点我去百度</a><br>
    <img alt="" class="has" src="http://www.ibp.cas.cn/kyjz/zxdt/201901/W020190103493057285919.jpg">
    <div class="testclass">class样式样式是否可以</div>
    <div style="color:#999fff">测试行内样式</div>
</body>
</html>
```

###### run.php

```php
<?php
    
use Fairy\MhtFileMaker;

// 1、保存为文件
MhtFileMaker::getInstance()
    ->addFile('index.html')//html文件
    ->completeImg('https://www.php.net', true)//补全图片的相对地址为绝对地址并且去除a链接
    ->makeFile('/home/www/a.doc');

// 2、浏览器下载
MhtFileMaker::getInstance()
    ->addFile('index.html')
    ->completeImg('https://www.php.net')
    ->download();

?>
```