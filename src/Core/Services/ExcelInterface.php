<?php

declare(strict_types=1);

namespace Ezijing\HyperfExcel\Core\Services;

interface ExcelInterface
{
    /**
     * 导出单个sheet的excel.
     *
     * @return mixed
     */
    public function exportExcelForASingleSheet(string $tableName, array $rows, array $data);

    /**
     * 导出多个sheet的excel.
     *
     * @return mixed
     */
    public function exportExcelWithMultipleSheets(string $tableName, array $sheets, array $rows, array $data);
}
