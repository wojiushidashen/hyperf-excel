<!-- START doctoc generated TOC please keep comment here to allow auto update -->
<!-- DON'T EDIT THIS SECTION, INSTEAD RE-RUN doctoc TO UPDATE -->
**Table of Contents**  *generated with [DocToc](https://github.com/thlorenz/doctoc)*

- [hyperf-excel](#hyperf-excel)
  - [安装准备](#%E5%AE%89%E8%A3%85%E5%87%86%E5%A4%87)
    - [1、确保在项目中安装了hyperf验证器](#1%E7%A1%AE%E4%BF%9D%E5%9C%A8%E9%A1%B9%E7%9B%AE%E4%B8%AD%E5%AE%89%E8%A3%85%E4%BA%86hyperf%E9%AA%8C%E8%AF%81%E5%99%A8)
    - [2、添加验证中间件 `config/autoload/middlewares.php`](#2%E6%B7%BB%E5%8A%A0%E9%AA%8C%E8%AF%81%E4%B8%AD%E9%97%B4%E4%BB%B6-configautoloadmiddlewaresphp)
    - [3、添加异常处理器 `config/autoload/exceptions.php`](#3%E6%B7%BB%E5%8A%A0%E5%BC%82%E5%B8%B8%E5%A4%84%E7%90%86%E5%99%A8-configautoloadexceptionsphp)
  - [安装](#%E5%AE%89%E8%A3%85)
    - [1、在项目根目录下执行](#1%E5%9C%A8%E9%A1%B9%E7%9B%AE%E6%A0%B9%E7%9B%AE%E5%BD%95%E4%B8%8B%E6%89%A7%E8%A1%8C)
    - [2、发布配置文件](#2%E5%8F%91%E5%B8%83%E9%85%8D%E7%BD%AE%E6%96%87%E4%BB%B6)
    - [3、配置文件 `config/autoload/excel_plugin.php`](#3%E9%85%8D%E7%BD%AE%E6%96%87%E4%BB%B6-configautoloadexcel_pluginphp)
    - [4、配置异常处理 `app/Exception/Handler/AppExceptionHandler.php`](#4%E9%85%8D%E7%BD%AE%E5%BC%82%E5%B8%B8%E5%A4%84%E7%90%86-appexceptionhandlerappexceptionhandlerphp)
  - [使用](#%E4%BD%BF%E7%94%A8)
    - [1、导出单个sheet的excel到本地](#1%E5%AF%BC%E5%87%BA%E5%8D%95%E4%B8%AAsheet%E7%9A%84excel%E5%88%B0%E6%9C%AC%E5%9C%B0)
    - [2、从浏览器导出单个sheet的excel](#2%E4%BB%8E%E6%B5%8F%E8%A7%88%E5%99%A8%E5%AF%BC%E5%87%BA%E5%8D%95%E4%B8%AAsheet%E7%9A%84excel)
    - [3、导出多个sheet的excel到本地](#3%E5%AF%BC%E5%87%BA%E5%A4%9A%E4%B8%AAsheet%E7%9A%84excel%E5%88%B0%E6%9C%AC%E5%9C%B0)
    - [4、从浏览器导出多个sheet的excel](#4%E4%BB%8E%E6%B5%8F%E8%A7%88%E5%99%A8%E5%AF%BC%E5%87%BA%E5%A4%9A%E4%B8%AAsheet%E7%9A%84excel)
    - [5、导入单个sheet的excel](#5%E5%AF%BC%E5%85%A5%E5%8D%95%E4%B8%AAsheet%E7%9A%84excel)
      - [测试数据](#%E6%B5%8B%E8%AF%95%E6%95%B0%E6%8D%AE)
      - [(1) 本地导入](#1-%E6%9C%AC%E5%9C%B0%E5%AF%BC%E5%85%A5)
        - [代码](#%E4%BB%A3%E7%A0%81)
        - [导入结果](#%E5%AF%BC%E5%85%A5%E7%BB%93%E6%9E%9C)
      - [(2) 接口导入](#2-%E6%8E%A5%E5%8F%A3%E5%AF%BC%E5%85%A5)
        - [请求方式](#%E8%AF%B7%E6%B1%82%E6%96%B9%E5%BC%8F)
        - [请求参数](#%E8%AF%B7%E6%B1%82%E5%8F%82%E6%95%B0)
        - [代码](#%E4%BB%A3%E7%A0%81-1)
        - [导入结果](#%E5%AF%BC%E5%85%A5%E7%BB%93%E6%9E%9C-1)

<!-- END doctoc generated TOC please keep comment here to allow auto update -->

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
    'export_way' => ExcelConstant::SAVE_TO_A_LOCAL_DIRECTORY, // 导出方式
    'enable_number' => true, // 是否开启序号
    'titles' => ['ID', '用户名', '部门', '职位'], // 设置表头
    'keys' => ['id', 'username', 'department', 'position'], // 设置表头标识，必须与要导出的数据的key对应
    // 要导出的数据
    'data' => [
        ['id' => '1', 'username' => '小明', 'department' => '运营部', 'position' => '产品运营'],
        ['id' => '2', 'username' => '小王', 'department' => '技术部', 'position' => 'PHP'],
    ],
    // 验证规则, 本地导入也适用
    'value_type' => [
        // 强转string
        ['key' => 'position', 'type' => 'string'],
        // 强转int
        ['key' => 'id', 'type' => 'int'],
        // 回调处理
        [
            'key' => 'department',
            'type' => 'function',
            'func' => function($value) {
                return (string) $value;
            },
        ],
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
            'enable_number' => false,
            'titles' => ['ID', '用户名', '部门', '职位'],
            'keys' => ['id', 'username', 'department', 'position'],
            'data' => [
                ['id' => '1', 'username' => '小明', 'department' => '运营部', 'position' => '产品运营'],
                ['id' => '2', 'username' => '小王', 'department' => '技术部', 'position' => 'PHP'],
            ],
            // 验证规则, 本地导入也适用
            'value_type' => [
                // 强转string
                ['key' => 'position', 'type' => 'string'],
                // 强转int
                ['key' => 'id', 'type' => 'int'],
                // 回调处理
                [
                    'key' => 'department',
                    'type' => 'function',
                    'func' => function($value) {
                        return (string) $value;
                    },
                ],
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
            'enable_number' => true, // 是否开启序号
            'titles' => ['ID', '用户名', '部门', '职位'],
            'keys' => ['id', 'username', 'department', 'position'],
            'data' => [
                ['id' => '1', 'username' => '小明', 'department' => '运营部', 'position' => '产品运营'],
                ['id' => '2', 'username' => '小王', 'department' => '技术部', 'position' => 'PHP'],
            ],
            // 验证规则, 本地导入也适用
            'value_type' => [
                // 强转string
                ['key' => 'position', 'type' => 'string'],
                // 强转int
                ['key' => 'id', 'type' => 'int'],
                // 回调处理
                [
                    'key' => 'department',
                    'type' => 'function',
                    'func' => function($value) {
                        return (string) $value;
                    },
                ],
            ]
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
            'enable_number' => false, // 是否开启序号
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
                    'enable_number' => true,
                    'titles' => ['ID', '用户名', '部门', '职位'],
                    'keys' => ['id', 'username', 'department', 'position'],
                    'data' => [
                        ['id' => '1', 'username' => '小明', 'department' => '运营部', 'position' => '产品运营'],
                        ['id' => '2', 'username' => '小王', 'department' => '技术部', 'position' => 'PHP'],
                    ],
                    // 验证规则, 本地导入也适用
                    'value_type' => [
                        // 强转string
                        ['key' => 'position', 'type' => 'string'],
                        // 强转int
                        ['key' => 'id', 'type' => 'int'],
                        // 回调处理
                        [
                            'key' => 'department',
                            'type' => 'function',
                            'func' => function($value) {
                                return (string) $value;
                            },
                        ],
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

### 5、导入单个sheet的excel

#### 测试数据
|ID|用户名|部门|职位|
|:---: |:----:|:----:|:-----:|
|1|小明|运营部|产品运营|
|2|小王|技术部|PHP|


#### (1) 本地导入
##### 代码
```php
$data = [
    // 带入方式
    'import_way' => ExcelConstant::THE_LOCAL_IMPORT,
    // 文件路径
    'file_path' => '/Users/ezijing/php_project/hyperf-demo/storage/excel/test_20220113_105250.xlsx',
    // 指定导入的title
    'titles' => ['部门', 'ID'],
    // 指定生成的key
    'keys' => ['position', 'id'],
];

$res = $this->excel->importExcelForASingleSheet($data);

print_r($res);
```
##### 导入结果
```
Array (
    [0] => Array
        (
            [id] => 1
            [position] => 运营部
        )

    [1] => Array
        (
            [id] => 2
            [position] => 技术部
        )

)
```

#### (2) 接口导入
##### 请求方式
> `POST`
##### 请求的类型
> `form-data`

##### 请求参数
|参数|类型|
|:---: |:----:|
|file|text|

##### 代码
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

    public function import()
    {
        $data = [
            'import_way' => ExcelConstant::BROWSER_IMPORT,
            // 指定导入的title
            'titles' => ['部门', 'ID', '职位'],
            // 指定生成的key
            'keys' => ['department', 'id', 'position'],
            // 验证规则, 本地导入也适用
            'value_type' => [
                // 强转string
                ['key' => 'position', 'type' => 'string'],
                // 强转int
                ['key' => 'id', 'type' => 'int'],
                // 回调处理
                [
                    'key' => 'department',
                    'type' => 'function',
                    'func' => function($value) {
                        return (string) $value;
                    },
                ],
            ]
        ];

        return $this->excel->importExcelForASingleSheet($data);
    }
}

```

##### 导入结果
```json
[
    {
        "id": 1,
        "department": "运营部",
        "position": "产品运营"
    },
    {
        "id": 2,
        "department": "技术部",
        "position": "PHP"
    }
]
```
