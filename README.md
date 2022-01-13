hyperf-excel
=====================================

安装准备
-------------------------------------
### 1、确保在项目中安装了hyperf验证器
```shell
> composer require hyperf/validation -vvv
> php bin/hyperf.php vendor:publish hyperf/translation # 发布 Translation 组件的文件
> php bin/hyperf.php vendor:publish hyperf/validation # 发布验证器组件的文件：
```
### 2、添加验证中间件 `config/autoload/middlewares.php`
```php
<?php
return [
    // 下面的 http 字符串对应 config/autoload/server.php 内每个 server 的 name 属性对应的值，意味着对应的中间件配置仅应用在该 Server 中
    'http' => [
        // 数组内配置您的全局中间件，顺序根据该数组的顺序
        \Hyperf\Validation\Middleware\ValidationMiddleware::class
        // 这里隐藏了其它中间件
    ],
];
```
### 3、添加异常处理器 `config/autoload/exceptions.php`
```php
<?php
return [
    'handler' => [
        // 这里对应您当前的 Server 名称
        'http' => [
            \Hyperf\Validation\ValidationExceptionHandler::class,
        ],
    ],
];
```

安装
----------------------------
### 1、在项目根目录下执行
```shell
> composer require ezijing/hyperf-excel -vvv
```

### 2、发布配置文件
```shell
>  php bin/hyperf.php vendor:publish ezijing/hyperf-excel
```

### 3、配置文件 `config/autoload/excel_plugin.php`
```php
<?php

declare(strict_types=1);

return [
    // 保存到本地的地址
    'local_file_address' => BASE_PATH . '/storage/excel',
];
```

### 4、配置异常处理 `app/Exception/Handler/AppExceptionHandler.php`
```php
public function handle(Throwable $throwable, ResponseInterface $response)
{
    switch (true) {
        case $throwable instanceof ExcelException:
            return $response
                ->withHeader('Sever', 'test')
                ->withStatus(200)
                ->withBody(new SwooleStream(Json::encode([
                    'code' => $throwable->getCode(),
                    'message' => $throwable->getMessage(),
                ])));

        default:
            $this->logger->error(sprintf('%s[%s] in %s', $throwable->getMessage(), $throwable->getLine(), $throwable->getFile()));
            $this->logger->error($throwable->getTraceAsString());
            return $response->withHeader('Server', 'Hyperf')->withStatus(500)->withBody(new SwooleStream('Internal Server Error.'));
    }
}
```

使用
---------------------------
### 1、导出单个sheet的excel到本地
```php
$tableName = 'test';
$data = [
    'export_way' => ExcelConstant::SAVE_TO_A_LOCAL_DIRECTORY,
    'titles' => ['ID', '用户名', '部门', '职位'],
    'keys' => ['id', 'username', 'department', 'position'],
    'data' => [
        ['id' => '1', 'username' => '小明', 'department' => '运营部', 'position' => '产品运营'],
        ['id' => '2', 'username' => '小王', 'department' => '技术部', 'position' => 'PHP'],
    ],
];

$res = (new Excel())->exportExcelForASingleSheet($tableName, $data);
```

### 2、从浏览器导出单个sheet的excel
```php
<?php

declare(strict_types=1);

namespace App\Controller;

use Ezijing\HyperfExcel\Core\Constants\ExcelConstant;
use Ezijing\HyperfExcel\Core\Services\Excel;
use Hyperf\HttpServer\Annotation\AutoController;

/**
 * @AutoController
 */
class ExcelController extends AbstractController
{
    /**
     * @var Excel
     */
    protected $excel;

    public function __construct()
    {
        $this->excel = make(Excel::class);
    }

    public function download()
    {
        $tableName = 'test';
        $data = [
            'export_way' => ExcelConstant::DOWNLOAD_TO_BROWSER_BY_TMP,
            'titles' => ['ID', '用户名', '部门', '职位'],
            'keys' => ['id', 'username', 'department', 'position'],
            'data' => [
                ['id' => '1', 'username' => '小明', 'department' => '运营部', 'position' => '产品运营'],
                ['id' => '2', 'username' => '小王', 'department' => '技术部', 'position' => 'PHP'],
            ],
        ];

        return $this->excel->exportExcelForASingleSheet($tableName, $data);
    }
}
```

### 3、导出多个sheet的excel到本地
```php
$tableName = 'sheets';
$data = [
    'export_way' => ExcelConstant::SAVE_TO_A_LOCAL_DIRECTORY,
    'sheets_params' => [
        [
            'sheet_title' => '企业1',
            'titles' => ['ID', '用户名', '部门', '职位'],
            'keys' => ['id', 'username', 'department', 'position'],
            'data' => [
                ['id' => '1', 'username' => '小明', 'department' => '运营部', 'position' => '产品运营'],
                ['id' => '2', 'username' => '小王', 'department' => '技术部', 'position' => 'PHP'],
            ],
        ],
        [
            'sheet_title' => '企业2',
            'titles' => ['ID', '用户名', '部门', '职位'],
            'keys' => ['id', 'username', 'department', 'position'],
            'data' => [
                ['id' => '3', 'username' => '小李', 'department' => '运营部', 'position' => '产品运营'],
                ['id' => '4', 'username' => '小赵', 'department' => '技术部', 'position' => 'PHP'],
            ],
        ],
        [
            'sheet_title' => '部门',
            'titles' => ['ID', '部门', '职位'],
            'keys' => ['id', 'department', 'position'],
            'data' => [
                ['id' => 1, 'department' => '运营部', 'position' => '产品运营'],
                ['id' => 2, 'department' => '技术部', 'position' => 'PHP'],
            ],
        ],
    ]
];

$res = (new Excel())->exportExcelWithMultipleSheets($tableName, $data);
print_r($res);
```

### 4、从浏览器导出多个sheet的excel
```php
<?php

declare(strict_types=1);

namespace App\Controller;

use Ezijing\HyperfExcel\Core\Constants\ExcelConstant;
use Ezijing\HyperfExcel\Core\Services\Excel;
use Hyperf\HttpServer\Annotation\AutoController;

/**
 * @AutoController
 */
class ExcelController extends AbstractController
{
    /**
     * @var Excel
     */
    protected $excel;

    public function __construct()
    {
        $this->excel = make(Excel::class);
    }

    public function download()
    {
        $tableName = 'sheets';
        $data = [
            'export_way' => ExcelConstant::DOWNLOAD_TO_BROWSER_BY_TMP,
            'sheets_params' => [
                [
                    'sheet_title' => '企业1',
                    'titles' => ['ID', '用户名', '部门', '职位'],
                    'keys' => ['id', 'username', 'department', 'position'],
                    'data' => [
                        ['id' => '1', 'username' => '小明', 'department' => '运营部', 'position' => '产品运营'],
                        ['id' => '2', 'username' => '小王', 'department' => '技术部', 'position' => 'PHP'],
                    ],
                ],
                [
                    'sheet_title' => '企业2',
                    'titles' => ['ID', '用户名', '部门', '职位'],
                    'keys' => ['id', 'username', 'department', 'position'],
                    'data' => [
                        ['id' => '3', 'username' => '小李', 'department' => '运营部', 'position' => '产品运营'],
                        ['id' => '4', 'username' => '小赵', 'department' => '技术部', 'position' => 'PHP'],
                    ],
                ],
                [
                    'sheet_title' => '部门',
                    'titles' => ['ID', '部门', '职位'],
                    'keys' => ['id', 'department', 'position'],
                    'data' => [
                        ['id' => 1, 'department' => '运营部', 'position' => '产品运营'],
                        ['id' => 2, 'department' => '技术部', 'position' => 'PHP'],
                    ],
                ],
            ]
        ];

        return $this->excel->exportExcelWithMultipleSheets($tableName, $data);
    }
}
```
