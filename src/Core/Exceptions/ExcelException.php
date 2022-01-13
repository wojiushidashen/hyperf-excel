<?php

declare(strict_types=1);

namespace Ezijing\HyperfExcel\Core\Exceptions;

use Ezijing\HyperfExcel\Core\Constants\ErrorCode;
use Throwable;

class ExcelException extends \RuntimeException
{
    public function __construct($code = ErrorCode::ERROR, $message = '', Throwable $previous = null)
    {
        if ($message == '') {
            $message = ErrorCode::getMessage($code) ?? '未知错误';
        }

        parent::__construct($message, $code, $previous);
    }
}
