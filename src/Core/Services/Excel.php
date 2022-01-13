<?php

declare(strict_types=1);

namespace Ezijing\HyperfExcel\Core\Services;

use Ezijing\HyperfExcel\Core\Constants\ErrorCode;
use Ezijing\HyperfExcel\Core\Constants\ExcelConstant;
use Ezijing\HyperfExcel\Core\Exceptions\ExcelException;
use Hyperf\Config\Annotation\Value;
use Hyperf\HttpMessage\Stream\SwooleStream;
use Hyperf\HttpServer\Response;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;

class Excel implements ExcelInterface
{
    /**
     * @var Spreadsheet
     */
    protected $spreadsheet;

    /**
     * 验证器.
     *
     * @var Validator
     */
    protected $validator;

    /**
     * 边框.
     *
     * @var array
     */
    protected $border;

    /**
     * 本地配置.
     *
     * @Value("excel_plugin")
     */
    protected $config;

    /**
     * 文件类型.
     *
     * @var string
     */
    private $_fileType = 'Xlsx';

    public function __construct()
    {
        $this->initSpreadsheet();
        $this->validator = app()->get(Validator::class);
        $this->initLocalFileDir();
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
        $cellMap = $this->cellMap();
        $maxCell = $cellMap[count($data['titles']) - 1];
        $worksheet->getStyle('A1:' . $maxCell . 1)->applyFromArray(array_merge($this->border, [
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER, // 水平居中对齐
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['argb' => 'fffbeeee'],
            ],
        ]));

        // 表头 设置单元格内容
        foreach ($data['titles'] as $key => $value) {
            $worksheet->setCellValueExplicitByColumnAndRow($key + 1, 1, $value, 's');
        }

        // 从第二行开始,填充表格数据
        $row = 2;
        foreach ($data['data'] as $item) {
            $worksheet->getStyle("A{$row}:" . $maxCell . $row)->applyFromArray($this->border);
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
    public function exportExcelWithMultipleSheets(string $tableName, array $data = [])
    {
        // 表格参数验证
        $this->validator->verify($data, [
            'export_way' => 'required|int|in:' . implode(',', array_keys(ExcelConstant::getExportWayMap())),
            'sheets_params' => 'required|array',
            'sheets_params.*.sheet_title' => 'required|string',
            'sheets_params.*.titles' => 'required|array|distinct',
            'sheets_params.*.keys' => 'required|array|distinct',
            'sheets_params.*.data' => 'present|array',
            'sheets_params.*.data.*' => 'required|array',
        ], [
            'sheets_params.required' => 'sheets参数未设置',
        ]);

        $firstSheet = true;
        foreach ($data['sheets_params'] as $sheetParamsValue) {
            if ($firstSheet) {
                $worksheet = $this->spreadsheet->getActiveSheet();
            } else {
                $worksheet = $this->spreadsheet->createSheet();
            }
            $firstSheet = false;

            // 设置工作表名称
            $worksheet->setTitle($sheetParamsValue['sheet_title']);
            $cellMap = $this->cellMap();
            $maxCell = $cellMap[count($sheetParamsValue['titles']) - 1];
            $worksheet->getStyle('A1:' . $maxCell . 1)->applyFromArray(array_merge($this->border, [
                'alignment' => [
                    'horizontal' => Alignment::HORIZONTAL_CENTER, // 水平居中对齐
                ],
                'fill' => [
                    'fillType' => Fill::FILL_SOLID,
                    'startColor' => ['argb' => 'fffbeeee'],
                ],
            ]));

            // 表头设置单元格内容
            foreach ($sheetParamsValue['titles'] as $titleKey => $titleValue) {
                $worksheet->setCellValueExplicitByColumnAndRow($titleKey + 1, 1, $titleValue, 's');
            }

            $row = 2;
            foreach ($sheetParamsValue['data'] as $item) {
                $worksheet->getStyle("A{$row}:" . $maxCell . $row)->applyFromArray($this->border);
                // 从第一列设置并初始化数据
                foreach ($item as $i => $v) {
                    $rowKey = array_search($i, $sheetParamsValue['keys']);
                    if ($rowKey === false) {
                        continue;
                    }
                    ++$rowKey;
                    $worksheet->setCellValueExplicitByColumnAndRow($rowKey, $row, $v, 's');
                }
                ++$row;
            }
        }
        // 默认打开第一个sheet
        $this->spreadsheet->setActiveSheetIndex(0);
        $fileName = $this->getFileName($tableName);

        return $this->downloadDistributor($data['export_way'], $fileName);
    }

    /**
     * 初始化spreadsheet.
     */
    protected function initSpreadsheet()
    {
        $this->spreadsheet = new Spreadsheet();
        $this->border = [
            'borders' => [
                //外边框
                'outline' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                ],
                //内边框
                'inside' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                ],
            ],
        ];
    }

    /**
     * 单元格.
     * @return string[]
     */
    protected function cellMap()
    {
        return ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
    }

    /**
     * 初始化保存到本地的文件夹.
     */
    protected function initLocalFileDir()
    {
        $this->config['local_file_address'] = $this->config['local_file_address'] ?? BASE_PATH . '/storage/excel';
        pathExists($this->config['local_file_address'] ?? BASE_PATH . '/storage/excel');
    }

    /**
     * 设置文件名.
     *
     * @return string
     */
    protected function getFileName(string $tableName)
    {
        return sprintf('%s_%s.xlsx', $tableName, date('Ymd_His'));
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
            case ExcelConstant::DOWNLOAD_TO_BROWSER_BY_TMP:
            default:
                return $this->saveToBrowserByTmp($fileName);
        }
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

    /**
     * 获取保存到本地的excel地址.
     *
     * @return string
     */
    protected function getLocalUrl(string $fileName)
    {
        return $this->config['local_file_address'] . DIRECTORY_SEPARATOR . $fileName;
    }
}
