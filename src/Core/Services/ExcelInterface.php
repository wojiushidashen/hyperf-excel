<?php

declare(strict_types=1);

namespace Ezijing\HyperfExcel\Core\Services;

interface ExcelInterface
{
    /**
     * 导出单个sheet的excel.
     *
     * @param string $tableName 表格名称
     * @param array $data 参数
     * @return mixed
     */
    public function exportExcelForASingleSheet(string $tableName, array $data = []);

    /**
     * 导出多个sheet的excel.
     *
     * @param string $tableName 表格名称
     * @param array $data 参数
     * @return mixed
     */
    public function exportExcelWithMultipleSheets(string $tableName, array $data = []);
}
