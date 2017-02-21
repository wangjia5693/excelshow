<?php namespace Excel\Readers;

use Closure;
use Excel\Excel;
use Excel\Exceptions\LaravelExcelException;

/**
 *
 * LaravelExcel Batch Importer
 */
class Batch {

    /**
     * Excel 实例
     */
    protected $excel;

    /**
     * 批量导入的文件
     */
    public $files = array();

    /**
     * 允许的扩展
     */
    protected $allowedFileExtensions = array(
        'xls',
        'xlsx',
        'csv'
    );

    /**
     * 执行批量导入
     * @param  Excel   $excel
     * @param  array   $files
     * @param  Closure $callback
     * @return Excel
     */
    public function start(Excel $excel, $files, Closure $callback)
    {
        // Set excel object
        $this->excel = $excel;

        // 获取所有文件
        $this->_setFiles($files);

        // Do the callback
        if ($callback instanceof Closure)
        {
            foreach ($this->getFiles() as $file)
            {
                // Load the file
                $excel = $this->excel->load($file);

                // 执行回调函数
                call_user_func($callback, $excel, $file);
            }
        }

        //返回excel实体；
        return $this->excel;
    }

    /**
     *获取所有文件属性
     * @return array
     */
    public function getFiles()
    {
        return $this->files;
    }

    /**
     * 设置批量导入文件
     * @param array|string $files
     * @throws LaravelExcelException
     * @return void
     */
    protected function _setFiles($files)
    {
        // 如果是数组，则是批量导入的文件
        if (is_array($files))
        {
            $this->files = $this->_getFilesByArray($files);
        }

        // 获取一个文件夹中的所有文件
        elseif (is_string($files))
        {
            $this->files = $this->_getFilesByFolder($files);
        }

        // Check if files were found
        if (empty($this->files))
            throw new LaravelExcelException('[ERROR]: No files were found. Batch terminated.');
    }

    /**
     * 通过数组形式导入文件
     * @param  array $array
     * @return array
     */
    protected function _getFilesByArray($array)
    {
        $files = array();
        // 确定是一个真实的路径 real paths
        foreach ($array as $i => $file)
        {
            $files[$i] = realpath($file) ? $file : base_path($file);
        }

        return $files;
    }

    /**
     * 获取一个文件夹中的所有文件
     * @param  string $folder
     * @return array
     */
    protected function _getFilesByFolder($folder)
    {
        // 确定是一个真实路径
        if (!realpath($folder))
            $folder = base_path($folder);

        // 获取所有扩展的文件
        $glob = glob($folder . '/*.{' . implode(',', $this->allowedFileExtensions) . '}', GLOB_BRACE);

        // If no matches, return empty array
        if ($glob === false) return array();

        // Return files
        return array_filter($glob, function ($file)
        {
            return filetype($file) == 'file';
        });
    }
}