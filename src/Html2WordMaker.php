<?php

namespace Fairy;

use Exception;
use Nette\Utils\Image;
use DiDom\Document;

/**
 * 快速html转word
 * 直接html写入word文档
 */
class Html2WordMaker extends MakerAbstract
{
    private $filename;
    private $eraseLink = false;
    private $replace = [];
    /**@var Document */
    private $document;
    protected $loadedTpl = false;

    protected function init()
    {
        $this->document = new Document();
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
     * 去除链接
     * @return $this
     */
    public function eraseLink()
    {
        $this->eraseLink = true;

        return $this;
    }

    /**
     * 该函数会分析img标签，提取src的属性值
     * 补全图片的相对地址为网络绝对地址
     * 分析文件内容并从远程下载页面中的图片资源并转化为base64格式
     * @param string $host
     * @return $this
     * @throws \Nette\Utils\UnknownImageFileException
     * @throws \DiDom\Exceptions\InvalidSelectorException
     */
    public function fetchImg($host = "")
    {
        if ($this->filename) {
            $this->document->loadHtmlFile($this->filename);
            $this->loadedTpl = true;
            $elements = $this->document->find('img');
            foreach ($elements as $element) {
                $src = $element->getAttribute('src');
                $width = $element->getAttribute('width');
                $height = $element->getAttribute('height');
                $style = $element->getAttribute('style');
                if (preg_match('/width\s*:\s*(\d+).*?;?/i', $style, $matches)) {
                    $width = $matches[1];
                }
                if (preg_match('/height\s*:\s*(\d+).*?;?/i', $style, $matches)) {
                    $height = $matches[1];
                }
                $imgUrl = !preg_match('/^http[s]?:\/\//i', $src) ? $host . $src : $src;
                $imgResource = $this->addImgFile($imgUrl, $width, $height);
                $this->replace[] = [
                    'search' => $src,
                    'replace' => 'data:' . $this->getMimeType($imgUrl) . ';base64,' . base64_encode($imgResource)
                ];
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
            $img->resize((int)$width, (int)$height);
        }

        return (string)$img;
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

        return file_put_contents($filename, $this->getFile()) > 0;
    }

    /**
     * 获取文件内容
     * @return false|string|string[]|null
     * @throws Exception
     */
    protected function getFile()
    {
        if (empty($this->filename)) {
            throw new Exception('No file was added.');
        }
        $content = file_get_contents($this->filename);
        if ($content) {
            $this->loadedTpl === false && $this->document->loadHtmlFile($this->filename);
            $htmlElement = $this->document->find('html');
            $htmlElement && $content = $htmlElement[0]->setAttribute('xmlns:o', 'urn:schemas-microsoft-com:office:office')
                ->setAttribute('xmlns:w', 'urn:schemas-microsoft-com:office:word')
                ->setAttribute('xmlns', 'http://www.w3.org/TR/REC-html40')
                ->html();
            $this->eraseLink && $content = preg_replace('/<a.*?>(.*?)<\/a>/i', '$1', $content);
            foreach ($this->replace as $row) {
                $search = addcslashes($row['search'], '/');
                $content = preg_replace("/{$search}/", $row['replace'], $content, 1);
            }
        }

        return $content;
    }
}
