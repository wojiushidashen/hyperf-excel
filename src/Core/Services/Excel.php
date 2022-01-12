<?php

declare(strict_types=1);

namespace Ezijing\HyperfExcel\Core\Services;

use Ezijing\HyperfExcel\Core\Constants\ErrorCode;
use Ezijing\HyperfExcel\Core\Constants\ExcelConstant;
use Ezijing\HyperfExcel\Core\Exceptions\ExcelException;
use Hyperf\HttpMessage\Stream\SwooleStream;
use Hyperf\HttpServer\Response;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class Excel implements ExcelInterface
{
    /**
     * @var Spreadsheet
     */
    protected $spreadsheet;

    /**
     * @var Validator
     */
    protected $validator;

    private $_fileType = 'Xlsx';

    public function __construct()
    {
        $this->spreadsheet = new Spreadsheet();
        $this->validator = app()->get(Validator::class);
    }

    /**
     * 导出单个sheet的excel.
     *
     * @return mixed|void
     */
    public function exportExcelForASingleSheet(string $tableName, array $data = [])
    {
        $this->validator->verify($data, [
            'export_way' => 'required|int|in:' . implode(',', array_keys(ExcelConstant::getExportWayMap())),
            'titles' => 'required|array|distinct',
            'keys' => 'required|array|distinct',
            'data' => 'present|array',
            'data.*' => 'required|array',
        ], [
            'titles.required' => '未设置表头',
            'keys.required' => '未设置列标识',
            'titles.unique' => '表头不能重复',
            'keys.unique' => '列标识不能重复',
        ]);

        if (count($data['titles']) != count($data['keys'])) {
            throw new ExcelException(ErrorCode::PARAMETER_ERROR, '列标识个数与表头标识个数要保持一致');
        }

        $worksheet = $this->spreadsheet->getActiveSheet();

        // 设置工作表标题名称
        $worksheet->setTitle($tableName);

        // 表头 设置单元格内容
        foreach ($data['titles'] as $key => $value) {
            $worksheet->setCellValueExplicitByColumnAndRow($key + 1, 1, $value, 's');
        }

        // 从第二行开始,填充表格数据
        $row = 2;
        foreach ($data['data'] as $item) {
            // 从第一列设置并初始化数据
            foreach ($item as $i => $v) {
                $rowKey = array_search($i, $data['keys']);
                if ($rowKey === false) {
                    continue;
                }
                ++$rowKey;
                $worksheet->setCellValueExplicitByColumnAndRow($rowKey, $row, $v, 's');
            }
            ++$row;
        }

        $fileName = $this->getFileName($tableName);

        return $this->downloadDistributor($data['export_way'], $fileName);
    }

    /**
     * 导出多个sheet的excel.
     *
     * @return mixed|void
     */
    public function exportExcelWithMultipleSheets(string $tableName, array $sheets, array $rows, array $data)
    {
    }

    /**
     * 设置文件名.
     *
     * @return string
     */
    protected function getFileName(string $tableName)
    {
        return sprintf('%s_%s.xlsx', $tableName, date('Y_m_d_His'));
    }

    /**
     * 保存到服务器本地.
     *
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     * @return string[]
     */
    protected function saveToLocal(string $fileName)
    {
        $url = $this->getLocalUrl($fileName);
        $writer = IOFactory::createWriter($this->spreadsheet, $this->_fileType);
        $writer->save($url);
        $this->spreadsheet->disconnectWorksheets();
        unset($this->spreadsheet);

        return [
            'path' => $url,
            'filename' => $fileName,
        ];
    }

    /**
     * 下载分发器.
     *
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     * @throws \Psr\Container\ContainerExceptionInterface
     * @throws \Psr\Container\NotFoundExceptionInterface
     * @return \Psr\Http\Message\ResponseInterface|string[]|void
     */
    protected function downloadDistributor(int $exportWay, string $fileName)
    {
        switch ($exportWay) {
            case ExcelConstant::SAVE_TO_A_LOCAL_DIRECTORY:
                return $this->saveToLocal($fileName);
            case ExcelConstant::DOWNLOAD_TO_BROWSER:
                return $this->saveToBrowser($fileName);
            case ExcelConstant::DOWNLOAD_TO_BROWSER_BY_TMP:
                return $this->saveToBrowserByTmp($fileName);
            default:
                return $this->saveToBrowserByTmp($fileName);
        }
    }

    /**
     * 直接从浏览器下载到本地（不建议使用）.
     *
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    protected function saveToBrowser(string $fileName)
    {
        $writer = IOFactory::createWriter($this->spreadsheet, $this->_fileType);
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $fileName . '"');
        header('Cache-Control: max-age=0');

        return $writer->save('php://output');
    }

    /**
     * 保存到临时文件再从浏览器自动下载到本地.
     *
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     * @throws \Psr\Container\ContainerExceptionInterface
     * @throws \Psr\Container\NotFoundExceptionInterface
     * @return \Psr\Http\Message\ResponseInterface
     */
    protected function saveToBrowserByTmp(string $fileName)
    {
        $localFileName = $this->getLocalUrl($fileName);
        $writer = IOFactory::createWriter($this->spreadsheet, $this->_fileType);

        // 保存到临时文件下
        $writer->save($localFileName);

        // 将文件转为字符串
        $content = file_get_contents($localFileName);

        // 删除临时文件
        unlink($localFileName);

        $response = app()->get(Response::class);

        $contentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

        return $response->withHeader('content-description', 'File Transfer')
            ->withHeader('content-type', $contentType)
            ->withHeader('content-description', "attachment;filename={$fileName}")
            ->withHeader('content-transfer-encoding', 'binary')
            ->withHeader('pragma', 'public')
            ->withBody(new SwooleStream((string) $content));
    }

    protected function getLocalUrl(string $fileName)
    {
        return BASE_PATH . '/storage/' . $fileName;
    }
}
