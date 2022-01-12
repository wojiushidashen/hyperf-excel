<?php

declare(strict_types=1);

if (! function_exists('app')) {
    function app()
    {
        return \Hyperf\Utils\ApplicationContext::getContainer();
    }
}

if (! function_exists('pathExists')) {
    /**
     * 判断文件夹是否存在，不存在则创建.
     *
     * @param $path
     */
    function pathExists($path)
    {
        mkdirs($path);
    }
}

if (! function_exists('mkdirs')) {
    /**
     * 创建文件夹.
     *
     * @param $dir
     * @param int $mode
     * @return bool
     */
    function mkdirs($dir, $mode = 0700)
    {
        if (is_dir($dir) || @mkdir($dir, $mode)) {
            return true;
        }

        if (! mkdir(dirname($dir), $mode)) {
            return false;
        }

        return @mkdir($dir, $mode);
    }
}
