<?php

namespace Fairy;

use Exception;

abstract class MakerAbstract
{
    /**
     * @var static
     */
    static protected $instance;

    static public function getInstance()
    {
        if (is_null(static::$instance)) {
            static::$instance = new static();
        }

        return static::$instance;
    }

    protected function __construct()
    {
        $this->init();
    }

    protected function __clone()
    {
    }

    protected function __wakeup()
    {
    }

    protected function init()
    {
    }

    /**
     * 添加html文件
     * @param $filename
     * @return mixed
     */
    abstract public function addFile($filename);

    /**
     * 去除链接
     * @return mixed
     */
    abstract public function eraseLink();

    /**
     * 获取图片资源
     * @param string $host
     * @return mixed
     */
    abstract public function fetchImg($host = "");

    /**
     * 生成word文件
     * @param $filename
     * @return mixed
     */
    abstract public function makeFile($filename);

    /**
     * 获取文件内容
     * @return mixed
     */
    abstract protected function getFile();

    /**
     * 浏览器下载
     * @param null $name
     * @param int $type
     * @throws Exception
     */
    public function download($name = null, $type = 2003)
    {
        $name = $name ?: date('YmdHis');
        $content = $this->getFile();

        header("Content-Disposition: attachment; filename=" . $name . $this->version($type)['ext']);
        header("Content-Type:" . $this->version($type)['mime']);
        header('Content-Transfer-Encoding: binary');
        header("Cache-Control: no-cache, must-revalidate, post-check=0, pre-check=0");
        header('Expires: 0');
        echo $content;
    }

    /**
     * 根据文件名获取MIME
     * @param $filename
     * @return string
     */
    protected function getMimeType($filename)
    {
        $ext = pathinfo($filename, PATHINFO_EXTENSION);
        switch (strtolower($ext)) {
            case 'htm':
            case 'html':
                $mimetype = 'text/html';
                break;
            case 'txt':
            case 'cgi':
            case 'php':
                $mimetype = 'text/plain';
                break;
            case 'css':
                $mimetype = 'text/css';
                break;
            case 'jpg':
            case 'jpeg':
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
    protected function version($type)
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
}
