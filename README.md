# HTML转WORD

html文件转word格式，欢迎 Star (●'◡'●)

>推荐转`.doc` 格式
>如果为`.docx` 格式，本机测试 Microsoft Office 2019 打开会有问题，WPS没问题

安装
------------

```bash
composer require cshaptx4869/html2word
```

实例
-------

###### html模板文件
图片仅支持 PNG, GIF, JPEG, WEBP 格式，且其地址必须是相对于域名根目录或者直接为网络地址

```html
<!doctype html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport"
        content="width=device-width, user-scalable=no, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <title>html template</title>
    <style>
        .testclass {
            font-size: 30px;
            color: #ff0000;
        }
    </style>
</head>

<body>
    <h1>中文的标题，技术无止境，一直在路上</h1>
    <p>p是可以分段的. 使用PHP将html转word</p>
    <p>再分一段 使用PHP将html转word</p>
    <p>还分一段，下面加个图片</p>
    <img src="/resource/houtarou.jpg" width="50px" height="50px">
    <br />
    <img alt="" class="has" src="http://www.ibp.cas.cn/kyjz/zxdt/201901/W020190103493057285919.jpg" style="width: 500px;">
    <div class="testclass">class样式样式是否可以</div>
    <div style="color:#999fff">测试行内样式</div>
    <a href="https://www.baidu.com">是否去除链接</a>
</body>

</html>
```

###### run.php

> Html2WordMaker 和 MhtFileMaker 都可以实现。只是原理稍微不同，前者是直接html写入文件，后者是借助mht再写入文件。

```php
<?php

require_once 'vendor/autoload.php';

use Fairy\Html2WordMaker;
use Fairy\MhtFileMaker;

// 1、保存为文件
Html2WordMaker::getInstance()
    ->addFile('resource/tpl.html')
    ->eraseLink()
    ->fetchImg('http://php.test/html2word')
    ->makeFile('resource/a.doc');

MhtFileMaker::getInstance()
    ->addFile('resource/tpl.html')
    ->eraseLink()
    ->fetchImg('http://php.test/html2word')
    ->makeFile('resource/a.doc');

// 2、浏览器下载
Html2WordMaker::getInstance()
    ->addFile('resource/tpl.html')
    ->fetchImg('http://php.test/html2word')
    ->download();

MhtFileMaker::getInstance()
    ->addFile('resource/tpl.html')
    ->fetchImg('http://php.test/html2word')
    ->download();
```

说明：

addFile(): 添加html模板文件

eraseLink(): 去除a链接

fetchImg(): 填充图片，若没有图片资源则不必调用。如果html模板文件中含没有域名的相对路径，则需在第一个参数中传入对应的域名

makeFile(): 保存为文件

download(): 浏览器下载