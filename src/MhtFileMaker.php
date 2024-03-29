<?php

namespace Fairy;

use Exception;
use Nette\Utils\Image;

/**
 * 由html导出word格式
 * word版本导出 2003(推荐.doc) 或者 2007(.docx Microsoft Office打开好像有点问题而wps没问题)
 *
 * 关于mht
 * Mht 会把网页中全部元素保存在一个文件里，不生成一个单独的文件夹，对于你文件的保存、管理会比较方便。
 */
class MhtFileMaker extends MakerAbstract
{
    protected $files = [];
    protected $boundary;
    protected $headers = [];
    protected $headersExists = [];
    protected $dirBase;
    protected $pageFirst;

    public function setHeader($header)
    {
        $this->headers[] = $header;
        $key = strtolower(substr($header, 0, strpos($header, ':')));
        $this->headersExists[$key] = true;
    }

    public function setDate($date = null, $istimestamp = false)
    {
        if ($date == null) {
            $date = time();
        }
        if ($istimestamp == true) {
            $date = date('D, d M Y H:i:s O', $date);
        }

        $this->setHeader("Date: $date");
    }

    public function setBoundary($boundary = null)
    {
        if ($boundary == null) {
            $this->boundary = '--' . strtoupper(md5(mt_rand())) . '_MULTIPART_MIXED';
        } else {
            $this->boundary = $boundary;
        }

        return $this;
    }

    public function setBaseDir($dir)
    {
        $this->dirBase = str_replace("\\", "/", realpath($dir));

        return $this;
    }

    public function setFirstPage($filename)
    {
        $this->pageFirst = str_replace("\\", "/", realpath("{$this->dirBase}/$filename"));

        return $this;
    }

    /**
     * @return $this
     * @throws Exception
     */
    public function autoAddFiles()
    {
        if (!isset($this->pageFirst)) {
            exit ('Not set the first page.');
        }
        $filepath = str_replace($this->dirBase, '', $this->pageFirst);
        $filepath = 'http://mhtfile' . $filepath;
        $this->addFile($this->pageFirst, $filepath);
        $this->addDir($this->dirBase);

        return $this;
    }

    /**
     * @param $dir
     * @return $this
     * @throws Exception
     */
    public function addDir($dir)
    {
        $handle = opendir($dir);
        while ($filename = readdir($handle)) {
            $path = $dir . DIRECTORY_SEPARATOR . $filename;
            if (($filename != '.') && ($filename != '..') && ($path != $this->pageFirst)) {
                if (is_dir($path)) {
                    $this->addDir($path);
                } elseif (is_file($path)) {
                    $filepath = str_replace($this->dirBase, '', $path);
                    $filepath = 'http://mhtfile' . $filepath;
                    $this->addFile($path, $filepath);
                }
            }
        }
        closedir($handle);

        return $this;
    }

    /**
     * 添加文件
     * @param $filename
     * @param null $filepath
     * @param null $encoding
     * @return $this
     * @throws Exception
     */
    public function addFile($filename, $filepath = null, $encoding = null)
    {
        if (!is_file($filename)) {
            throw new Exception('file:' . $filename . ' not exist!');
        }

        if ($filepath == null) {
            $filepath = $filename;
        }
        $mimetype = $this->getMimeType($filename);
        $filecont = file_get_contents($filename);
        $this->addContents($filepath, $mimetype, $filecont, $encoding);

        return $this;
    }

    /**
     * 添加文本信息
     * @param string $filepath 文件地址
     * @param $mimetype
     * @param $filecont
     * @param null $encoding
     */
    protected function addContents($filepath, $mimetype, $filecont, $encoding = null)
    {
        $this->files[] = [
            'filepath' => $filepath,
            'mimetype' => $mimetype,
            'filecont' => chunk_split(base64_encode($filecont)),
            'encoding' => $encoding ?: 'base64'
        ];
    }

    /**
     * 去除链接
     * @return $this
     */
    public function eraseLink()
    {
        foreach ($this->files as &$file) {
            $content = base64_decode(str_replace(PHP_EOL, '', $file['filecont']));
            $content = preg_replace('/<a.*?>(.*?)<\/a>/i', '$1', $content);
            $file['filecont'] = chunk_split(base64_encode($content));
        }

        return $this;
    }

    /**
     * 该函数会分析img标签，提取src的属性值
     * 补全图片的相对地址为网络绝对地址
     * 分析文件内容并从远程下载页面中的图片资源
     * @param string $host
     * @return $this
     * @throws \Nette\Utils\UnknownImageFileException
     */
    public function fetchImg($host = "")
    {
        $images = [];
        foreach ($this->files as &$file) {
            $content = base64_decode(str_replace(PHP_EOL, '', $file['filecont']));
            if ($host) {
                $content = preg_replace_callback('/<img.*?(?:>|\/>)/is', function ($imageTag) use ($host) {
                    $imageTag = preg_replace_callback('/(src\s*=\s*[\'\"]?)([^\'\"]*)([\'\"]?)/i', function ($src) use ($host) {
                        if (!preg_match('/^http[s]?:\/\//i', trim($src[2]))) {
                            $url = $host . $src[2];
                            $file['filepath'] = $url;

                            return $src[1] . $url . $src[3];
                        } else {
                            return $src[0];
                        }
                    }, $imageTag[0]);

                    return $imageTag;
                }, $content);
                $file['filecont'] = chunk_split(base64_encode($content));
            }

            // 匹配图片
            if (preg_match_all('/<img.*?(?:>|\/>)/is', $content, $imgTags)) {
                foreach ($imgTags[0] as $imgTag) {
                    if (preg_match('/src\s*=\s*[\'\"]?([^\'\"]*)[\'\"]?/i', $imgTag, $srcPaths)) {
                        $srcPath = trim($srcPaths[1]);
                        if ($srcPath != "") {
                            $width = $height = null;
                            if (preg_match('/width\s*[=:]\s*[\'\"]?(\d+)[\'\"]?/i', $imgTag, $sizes)) {
                                $width = $sizes[1];
                            }
                            if (preg_match('/height\s*[=:]\s*[\'\"]?(\d+).*?[\'\"]?/i', $imgTag, $sizes)) {
                                $height = $sizes[1];
                            }
                            $images[] = [
                                'path' => $srcPath,
                                'width' => $width,
                                'height' => $height
                            ];
                        }
                    }
                }
            }
        }
        $this->addImgFile($images);

        return $this;
    }

    /**
     * 下载图片资源
     * @param array $images
     * @throws \Nette\Utils\UnknownImageFileException
     */
    protected function addImgFile(array $images)
    {
        foreach ($images as $image) {
            $imgObj = Image::fromFile($image['path']);
            if ($image['width'] || $image['height']) {
                $imgObj->resize((int)$image['width'], (int)$image['height']);
            }
            $this->addContents($image['path'], $this->getMimeType($image['path']), (string)$imgObj);
        }
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

        return file_put_contents($filename, $content) > 0;
    }

    /**
     * 获取文件
     * @return string
     * @throws Exception
     */
    protected function getFile()
    {
        $this->checkHeaders();
        if (count($this->files) === 0) {
            throw new Exception('No file was added.');
        }
        $enter = PHP_EOL;
        $contents = implode($enter, $this->headers);
        $contents .= $enter;
        $contents .= "MIME-Version: 1.0 {$enter}";
        $contents .= "Content-Type: multipart/related;{$enter}";
        $contents .= "\tboundary=\"{$this->boundary}\";{$enter}";
        $contents .= "\ttype=\"" . $this->files[0]['mimetype'] . "\"{$enter}";
        $contents .= "X-MimeOLE: Produced By Mht File Maker v1.0 beta{$enter}";
        $contents .= "{$enter}";
        $contents .= "This is a multi-part message in MIME format.{$enter}";
        $contents .= "{$enter}";
        foreach ($this->files as $file) {
            $contents .= "--{$this->boundary}{$enter}";
            $contents .= "Content-Type: $file[mimetype]{$enter}";
            $contents .= "Content-Transfer-Encoding: $file[encoding]{$enter}";
            $contents .= "Content-Location: $file[filepath]{$enter}";
            $contents .= "{$enter}";
            $contents .= $file['filecont'];
            $contents .= "{$enter}";
        }
        $contents .= "--{$this->boundary}--{$enter}";

        return $contents;
    }

    /**
     * 检查生成完整文档必要的信息
     */
    public function checkHeaders()
    {
        if (!array_key_exists('date', $this->headersExists)) {
            $this->setDate(null, true);
        }
        if ($this->boundary == null) {
            $this->setBoundary();
        }
    }
}
