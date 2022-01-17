<?php

declare(strict_types=1);

namespace Ezijing\HyperfExcel\Core\Services;

use Ezijing\HyperfExcel\Core\Constants\ErrorCode;
use Ezijing\HyperfExcel\Core\Constants\ExcelConstant;
use Ezijing\HyperfExcel\Core\Exceptions\ExcelException;
use Hyperf\Config\Annotation\Value;
use Hyperf\HttpMessage\Stream\SwooleStream;
use Hyperf\HttpServer\Request;
use Hyperf\HttpServer\Response;
use Hyperf\Utils\Arr;
use Hyperf\Utils\Coroutine;
use Hyperf\Utils\Exception\ParallelExecutionException;
use Hyperf\Utils\Parallel;
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
     * 值类型映射.
     *
     * @var string[]
     */
    protected $valueTypeMap = [
        'int', // int
        'string', // 字符串
        'date', // Y-H-d H:i:s
        'time', // 秒级时间戳
        'float', // 转为浮点型
        'function', // 函数
    ];

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

    public function import(array $data): array
    {
        $validation = [
            'import_way' => 'required|int|in:' . implode(',', array_keys(ExcelConstant::getImportWayMap())),
            'file_path' => ['required_if:import_way,' . ExcelConstant::THE_LOCAL_IMPORT, 'string', 'regex:/(?:[x|X][l|L][s|S][x|X])$/'],
        ];

        $noticeMesasage = [
            'import_way.in' => '导入方式错误',
            'file_path.required_if' => '本地上传file_path字段不能为空',
            'file_path.regx' => '上传文件只支持.Xlsx格式',
        ];

        $this->validator->verify($data, $validation, $noticeMesasage);

        switch ($data['import_way']) {
            case ExcelConstant::THE_LOCAL_IMPORT:
                $filePath = $data['file_path'];
                break;
            default:
                $file = $this->importFileRequetVerify();
                $filePath = $file['tmp_file'];
        }

        $spreadsheet = IOFactory::load($filePath);

        $sheets = $spreadsheet->getAllSheets();

        $data = [];

        foreach ($sheets as $sheet) {
            $data[] = $sheet->toArray();
        }

        return $data;
    }

    /**
     * 导入单个sheet的excel入口.
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \Psr\Container\ContainerExceptionInterface
     * @throws \Psr\Container\NotFoundExceptionInterface
     */
    public function importExcelForASingleSheet(array $data = []): array
    {
        $validation = [
            'import_way' => 'required|int|in:' . implode(',', array_keys(ExcelConstant::getImportWayMap())),
            // 验证文件
            'file_path' => ['required_if:import_way,' . ExcelConstant::THE_LOCAL_IMPORT, 'string', 'regex:/(?:[x|X][l|L][s|S][x|X])$/'],
            'titles' => 'required|array',
            'titles.*' => 'required|distinct',
            'keys' => 'required|array',
            'keys.*' => 'required|distinct',
            'value_type' => 'array',
            'value_type.*.key' => 'required',
            'value_type.*.type' => 'required',
            'value_type.*.func' => 'required_if:value_type.*.type,function',
        ];
        $noticeMesasage = [
            'import_way.in' => '导入方式错误',
            'file_path.required_if' => '本地上传file_path字段不能为空',
            'file_path.regx' => '上传文件只支持.Xlsx格式',
        ];
        $this->validator->verify($data, $validation, $noticeMesasage);

        // 强制titles和keys的数量保持一致
        if (count($data['titles']) != count($data['keys'])) {
            throw new ExcelException(ErrorCode::PARAMETER_ERROR, 'titles和keys要保持对应');
        }

        return $this->importDistributor($data);
    }

    /**
     * 导入分发器.
     *
     * @param array $data 数据
     * @throws \Psr\Container\ContainerExceptionInterface
     * @throws \Psr\Container\NotFoundExceptionInterface
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @return array
     */
    protected function importDistributor(array $data)
    {
        switch ($data['import_way']) {
            case ExcelConstant::THE_LOCAL_IMPORT:
                return $this->importExcelForASingleSheetLocal($data);
            default:
                return $this->importExcelForASingleSheetBrowser($data);
        }
    }

    /**
     * 从本地导入多个sheet的excel.
     *
     * @param array $data 数据
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @return array
     */
    protected function importExcelForASingleSheetLocal(array $data)
    {
        $filePath = $data['file_path'];

        return $this->readExcelForASingleSheet($filePath, $data);
    }

    /**
     * 从浏览器导入单个sheet的excel.
     *
     * @param array $data 数据
     * @throws \Psr\Container\ContainerExceptionInterface
     * @throws \Psr\Container\NotFoundExceptionInterface
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @return array
     */
    protected function importExcelForASingleSheetBrowser(array $data)
    {
        $file = $this->importFileRequetVerify();

        return $this->readExcelForASingleSheet($file['tmp_file'], $data);
    }

    /**
     * 通过接口请求获取file.
     *
     * @throws \Psr\Container\ContainerExceptionInterface
     * @throws \Psr\Container\NotFoundExceptionInterface
     * @return array
     */
    protected function importFileRequetVerify()
    {
        $request = app()->get(Request::class);

        if ($request->getMethod() != 'POST') {
            throw new ExcelException(ErrorCode::FOR_EXAMPLE_IMPORT_DATA, '只接收POST请求');
        }

        $file = $request->file('file');
        if (! $file || ! $file->isValid()) {
            throw new ExcelException(ErrorCode::FAILED_TO_IMPORT_FILES_PROCEDURE);
        }
        $file = $file->toArray();

        // 获取文件上传的临时文件
        if (! isset($file['tmp_file'])) {
            throw new ExcelException(ErrorCode::FAILED_TO_IMPORT_FILES_PROCEDURE);
        }

        return $file;
    }

    /**
     * 格式化单个sheet的excel.
     *
     * @param string $filePath 文件路径
     * @param array $data 数据
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @return array
     */
    protected function readExcelForASingleSheet(string $filePath, array $data)
    {
        $objRead = IOFactory::createReader($this->_fileType);

        $inputFileType = IOFactory::identify($filePath);
        if (strtolower($inputFileType) != 'xlsx') {
            throw new ExcelException(ErrorCode::PARAMETER_ERROR, '只支持导入Xlsx文件');
        }

        if (! $objRead->canRead($filePath)) {
            throw new ExcelException(ErrorCode::FAILED_TO_IMPORT_FILES_PROCEDURE, '只支持导入Excel文件');
        }

        $objRead->setReadDataOnly(true);
        $objRead->setReadEmptyCells(false);

        $spreadsheet = @$objRead->load($filePath);

        if (empty($list = $spreadsheet->getSheet(0)->toArray())) {
            return [];
        }

        // 强制第一行必须是表头，无导入的数据
        if (count($list) <= 1) {
            return [];
        }

        // 获取表头、非正常格式不作处理
        if (empty($headers = $list[0])) {
            throw new ExcelException(ErrorCode::FOR_EXAMPLE_IMPORT_DATA);
        }

        // 匹配自定义的titles和keys
        $titleMap = array_combine($data['titles'], $data['keys']);

        // 获取指定key的映射结果
        $keyMap = [];
        foreach ($headers as $headerIndex => $header) {
            if ($header && isset($titleMap[trim($header)])) {
                $keyMap[$headerIndex] = $titleMap[trim($header)];
            }
        }
        if (empty($keyMap)) {
            return [];
        }

        // 去除表头
        array_shift($list);

        // 初始化格式化的数据
        $formatData = [];

        $valueTypes = $data['value_type'] ?? [];

        // 开启携程，格式化数据
        $parallel = new Parallel(5);
        foreach ($list as $index => &$item) {
            $parallel->add(function () use (&$item, &$formatData, $index, $keyMap, $valueTypes) {
                foreach ($keyMap as $keyIndex => $key) {
                    $value = $item[$keyIndex];
                    // 格式化值类型
                    if ($valueTypes) {
                        $keys = Arr::pluck($valueTypes, 'key');
                        $valueTypes = array_combine($keys, $valueTypes);
                        if (isset($valueTypes[$key])) {
                            $value = $this->formatValue($key, $valueTypes[$key]['type'], $value, $valueTypes[$key]['func'] ?? null);
                        }
                    }
                    $formatData[$index][$key] = $value;
                }
                if (empty(implode('', $formatData[$index]))) {
                    unset($formatData[$index]);
                }
                return Coroutine::id();
            });
        }
        try {
            $parallel->wait();
        } catch (ParallelExecutionException $e) {
            throw new ExcelException(ErrorCode::ERROR, $e->getMessage());
        }

        return $formatData;
    }

    /**
     * 格式化值的类型.
     *
     * @param $key 键
     * @param $type 值类型
     * @param $value 值
     * @param null $func 回调函数
     * @return false|int|mixed|string
     */
    protected function formatValue($key, $type, $value, $func = null)
    {
        $typeMap = array_flip($this->valueTypeMap);
        if (! isset($typeMap[strtolower($type)])) {
            throw new ExcelException(ErrorCode::PARAMETER_ERROR, 'value_type.*.type error');
        }

        switch (strtolower($type)) {
            case 'int':
                $value = (int) $value;
                break;
            case 'string':
                $value = (string) $value;
                break;
            case 'date':
                $value = date('Y-m-d H:i:s', (int) $value);
                break;
            case 'time':
                $value = strtotime((string) $value);
                break;
            case 'function':
                if ($func) {
                    $value = $func($value);
                }
                break;
            default:
        }

        return $value;
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
        $writer = @IOFactory::createWriter($this->spreadsheet, $this->_fileType);
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

        $contentType = 'text/xlsx';

        return $response->withHeader('content-description', 'File Transfer')
            ->withHeader('content-type', $contentType)
            ->withHeader('content-description', "attachment;filename={$fileName}")
            ->withHeader('content-transfer-encoding', 'binary')
            ->withHeader('pragma', 'public')
            ->withHeader('file_name', urlencode($fileName))
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
