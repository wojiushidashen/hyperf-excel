<?php

declare(strict_types=1);

if (! function_exists('app')) {
    function app()
    {
        return \Hyperf\Utils\ApplicationContext::getContainer();
    }
}

if (! function_exists('fastexcel')) {
    /**
     * Return app instance of FastExcel.
     *
     * @param null $data
     * @return bool
     */
    function fastexcel($data = null)
    {
        if ($data instanceof \Ezijing\HyperfExcel\SheetCollection) {
            return app()->make(\Ezijing\HyperfExcel\FastExcel::class)->data($data);
        }

        if (is_object($data) && method_exists($data, 'toArray')) {
            $data = $data->toArray();
        }

        return $data ?? app()->make(\Ezijing\HyperfExcel\FastExcel::class);
    }
}
