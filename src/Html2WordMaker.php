<?php

namespace Fairy;

use Exception;
use Nette\Utils\Image;
use phpQuery;

/**
 * 快速html转word
 * 直接html写入word文档
 * Class Html2WordMaker
 * @package Fairy
 */
class Html2WordMaker
{
    static private $instance;
    private $filename;
    private $eraseLink = false;
    private $replace = [];

    /**
     * 单例
     * @return Html2WordMaker
     */
    static public function getInstance()
    {
        if (is_null(self::$instance)) {
            self::$instance = new self();
        }
        return self::$instance;
    }

    protected function __construct()
    {

    }

    /**
     * 添加html文件
     * @param $filename
     * @return $this
     * @throws Exception
     */
    public function addFile($filename)
    {
        if (!is_file($filename)) {
            throw new Exception('file:' . $filename . ' not exist!');
        }
        $this->filename = $filename;
        return $this;
    }

    /**
     * 该函数会分析img标签，提取src的属性值
     * 补全图片的相对地址为网络绝对地址
     * 分析文件内容并从远程下载页面中的图片资源并转化为base64格式
     * @param string $host
     * @return $this
     * @throws \Nette\Utils\UnknownImageFileException
     */
    public function fetchImg($host = "")
    {
        if ($this->filename) {
            $imgQuery = phpQuery::newDocumentFileHTML($this->filename)->find('img');
            if ($imgQuery->length) {
                foreach ($imgQuery as $img) {
                    $img = pq($img);
                    $src = $img->attr('src');
                    $width = $img->attr('width');
                    $height = $img->attr('height');
                    $style = $img->attr('style');
                    if (preg_match('/width\s*:\s*(\d+).*?;?/i', $style, $matches)) {
                        $width = $matches[1];
                    }
                    if (preg_match('/height\s*:\s*(\d+).*?;?/i', $style, $matches)) {
                        $height = $matches[1];
                    }
                    if (!preg_match('/^http[s]?:\/\//i', $src)) {
                        $imgUrl = $host . $src;
                    } else {
                        $imgUrl = $src;
                    }
                    $imgResource = $this->addImgFile($imgUrl, $width, $height);
                    $this->replace[] = [
                        'search' => $src,
                        'replace' => 'data:' . $this->getMimeType($imgUrl) . ';base64,' . base64_encode($imgResource)
                    ];
                }
            }
        }
        return $this;
    }

    /**
     * @param $imgUrl
     * @param null $width
     * @param null $height
     * @return string
     * @throws \Nette\Utils\UnknownImageFileException
     */
    protected function addImgFile($imgUrl, $width = null, $height = null)
    {
        $img = Image::fromFile($imgUrl);
        if ($width || $height) {
            $img->resize($width, $height);
        }
        return (string)$img;
    }

    /**
     * 获取mime
     * @param $filename
     * @return string
     */
    public function getMimeType($filename)
    {
        $ext = pathinfo($filename, PATHINFO_EXTENSION);
        switch (strtolower($ext)) {
            case 'htm':
                $mimetype = 'text/html';
                break;
            case 'html':
                $mimetype = 'text/html';
                break;
            case 'txt':
                $mimetype = 'text/plain';
                break;
            case 'cgi':
                $mimetype = 'text/plain';
                break;
            case 'php':
                $mimetype = 'text/plain';
                break;
            case 'css':
                $mimetype = 'text/css';
                break;
            case 'jpg':
                $mimetype = 'image/jpeg';
                break;
            case 'jpeg':
                $mimetype = 'image/jpeg';
                break;
            case 'jpe':
                $mimetype = 'image/jpeg';
                break;
            case 'gif':
                $mimetype = 'image/gif';
                break;
            case 'png':
                $mimetype = 'image/png';
                break;
            default:
                $mimetype = 'application/octet-stream';
                break;
        }
        return $mimetype;
    }

    /**
     * @param $type
     * @return mixed
     * @throws Exception
     */
    public function version($type)
    {
        $versionOpt = [
            2003 => [
                'mime' => 'application/msword',
                'ext' => '.doc',
            ],
            2007 => [
                'mime' => 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                'ext' => '.docx',
            ]
        ];

        if (!isset($versionOpt[$type])) {
            throw new Exception('error type');
        }
        return $versionOpt[$type];
    }

    /**
     * 去除链接
     * @return $this
     */
    public function eraseLink()
    {
        $this->eraseLink = true;
        return $this;
    }

    /**
     * 检查是否有文本资源信息
     * @return bool
     */
    public function checkFile()
    {
        return $this->filename ? true : false;
    }

    /**
     * 获取文件
     * @return false|string|string[]|null
     * @throws Exception
     */
    public function getFile()
    {
        if (!$this->checkFile()) {
            throw new Exception('No file was added.');
        }
        $content = file_get_contents($this->filename);
        if ($content) {
            if ($this->eraseLink) {
                $content = preg_replace('/<a.*?>(.*?)<\/a>/i', '$1', $content);
            }
            foreach ($this->replace as $row) {
                $search = addcslashes($row['search'], '/');
                $content = preg_replace("/{$search}/", $row['replace'], $content, 1);
            }
        }
        return $content;
    }

    /**
     * 生成文件
     * @param string $filename xxx.doc or xxx.docx
     * @return bool
     * @throws Exception
     */
    public function makeFile($filename)
    {
        $ext = pathinfo($filename, PATHINFO_EXTENSION);
        if (!in_array(strtolower($ext), ['doc', 'docx'])) {
            throw new Exception('not support file type ' . $ext);
        }
        $content = $this->getFile();
        $content = '<html 
            xmlns:o="urn:schemas-microsoft-com:office:office" 
            xmlns:w="urn:schemas-microsoft-com:office:word" 
            xmlns="http://www.w3.org/TR/REC-html40">
            <meta charset="UTF-8" />' . $content . '</html>';

        return file_put_contents($filename, $content) > 0;
    }

    /**
     * 浏览器下载
     * @param null $name
     * @param int $type
     * @throws Exception
     */
    public function download($name = null, $type = 2003)
    {
        $name = $name ? $name : date('YmdHis');
        $content = $this->getFile();

        header("Content-Disposition: attachment; filename=" . $name . $this->version($type)['ext']);
        header("Content-Type:" . $this->version($type)['mime']);
        header('Content-Transfer-Encoding: binary');
        header("Cache-Control: no-cache, must-revalidate, post-check=0, pre-check=0");
        header('Expires: 0');

        $html = '<html xmlns:v="urn:schemas-microsoft-com:vml"
         xmlns:o="urn:schemas-microsoft-com:office:office"
         xmlns:w="urn:schemas-microsoft-com:office:word" 
         xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" 
         xmlns="http://www.w3.org/TR/REC-html40">';
        $html .= '<head><meta http-equiv="Content-Type" content="text/html;charset="UTF-8" /></head>';

        echo $html . '<body>' . $content . '</body></html>';
    }

    protected function __clone()
    {

    }

    protected function __wakeup()
    {

    }
}