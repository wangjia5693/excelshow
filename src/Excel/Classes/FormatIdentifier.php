<?php namespace Excel\Classes;

use PHPExcel_IOFactory;
use Illuminate\Filesystem\Filesystem;
use Excel\Exceptions\LaravelExcelException;

class FormatIdentifier {

    /**
     * 格式化类型
     */
    protected $formats = array(
        'Excel2007',
        'Excel5',
        'Excel2003XML',
        'OOCalc',
        'SYLK',
        'Gnumeric',
        'CSV',
        'HTML',
        'PDF'
    );

    /**
     * 启动引入文件系统
     */
    public function __construct(Filesystem $filesystem)
    {
        $this->filesystem = $filesystem;
    }

    /**
     * Get the file format by file
     * @param $file
     * @throws LaravelExcelException
     * @return string $format
     */
    public function getFormatByFile($file)
    {
        // 获取文件扩展
        $ext = $this->getExtension($file);

        //确定真实的扩展Excel 2007，Excel 2005，html,pdf等
        $format = $this->getFormatByExtension($ext);

        // 判断文件是否可读
        if ($this->canRead($format, $file))
            return $format;

        // Do a last try to init the file with all available readers
        return $this->lastResort($file, $format, $ext);
    }

    /**
     * 确定真正扩展
     */
    public function getFormatByExtension($ext)
    {
        switch ($ext)
        {

            /*
            |--------------------------------------------------------------------------
            | Excel 2007
            |--------------------------------------------------------------------------
            */
            case 'xlsx':
            case 'xlsm':
            case 'xltx':
            case 'xltm':
                return 'Excel2007';
                break;

            /*
            |--------------------------------------------------------------------------
            | Excel5
            |--------------------------------------------------------------------------
            */
            case 'xls':
            case 'xlt':
                return 'Excel5';
                break;

            /*
            |--------------------------------------------------------------------------
            | OOCalc
            |--------------------------------------------------------------------------
            */
            case 'ods':
            case 'ots':
                return 'OOCalc';
                break;

            /*
            |--------------------------------------------------------------------------
            | SYLK
            |--------------------------------------------------------------------------
            */
            case 'slk':
                return 'SYLK';
                break;

            /*
            |--------------------------------------------------------------------------
            | Excel2003XML
            |--------------------------------------------------------------------------
            */
            case 'xml':
                return 'Excel2003XML';
                break;

            /*
            |--------------------------------------------------------------------------
            | Gnumeric
            |--------------------------------------------------------------------------
            */
            case 'gnumeric':
                return 'Gnumeric';
                break;

            /*
            |--------------------------------------------------------------------------
            | HTML
            |--------------------------------------------------------------------------
            */
            case 'htm':
            case 'html':
                return 'HTML';
                break;

            /*
            |--------------------------------------------------------------------------
            | CSV
            |--------------------------------------------------------------------------
            */
            case 'csv':
            case 'txt':
                return 'CSV';
                break;

            /*
            |--------------------------------------------------------------------------
            | PDF
            |--------------------------------------------------------------------------
            */
             case 'pdf':
                 return 'PDF';
                 break;
        }
    }

    /**
     * 根据扩展确定header content头
     */
    public function getContentTypeByFormat($format)
    {
        switch ($format)
        {

            /*
            |--------------------------------------------------------------------------
            | Excel 2007
            |--------------------------------------------------------------------------
            */
            case 'Excel2007':
                return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; charset=UTF-8';
                break;

            /*
            |--------------------------------------------------------------------------
            | Excel5
            |--------------------------------------------------------------------------
            */
            case 'Excel5':
                return 'application/vnd.ms-excel; charset=UTF-8';
                break;

            /*
            |--------------------------------------------------------------------------
            | HTML
            |--------------------------------------------------------------------------
            */
            case 'HTML':
                return 'HTML';
                break;

            /*
            |--------------------------------------------------------------------------
            | CSV
            |--------------------------------------------------------------------------
            */
            case 'CSV':
                return 'application/csv; charset=UTF-8';
                break;

            /*
            |--------------------------------------------------------------------------
            | PDF
            |--------------------------------------------------------------------------
            */
             case 'PDF':
                 return'application/pdf; charset=UTF-8';
                 break;
        }
    }

    /**
     * 对于无法判断类型的，将会使用现有支持的所有类型去匹配，最终返回类型
     */
    protected function lastResort($file, $wrongFormat = false, $ext = 'xls')
    {
        // Loop through all available formats
        foreach ($this->formats as $format)
        {
            // Check if the file could be read
            if ($wrongFormat != $format && $this->canRead($format, $file))
                return $format;
        }

        // 异常
        throw new LaravelExcelException('[ERROR] Reader could not identify file format for file [' . $file . '] with extension [' . $ext . ']');
    }

    /**
     * 判断文件是否可读，返回reader
     */
    protected function canRead($format, $file)
    {
        if ($format)
        {
            $reader = $this->initReader($format);

            return $reader && $reader->canRead($file);
        }

        return false;
    }

    /**
     * 根据格式化类型穿件读取
     */
    protected function initReader($format)
    {
        return PHPExcel_IOFactory::createReader($format);
    }

    /**
     * 获取文件的扩展
     */
    protected function getExtension($file)
    {
        return strtolower($this->filesystem->extension($file));
    }
}