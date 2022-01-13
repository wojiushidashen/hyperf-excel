<?php

declare(strict_types=1);

namespace Ezijing\HyperfExcel\Core\Constants;

use Hyperf\Constants\AbstractConstants;
use Hyperf\Constants\Annotation\Constants;

/**
 * @Constants
 */
class ErrorCode extends AbstractConstants
{
    /**
     * @Message("处理Excel异常")
     */
    public const ERROR = 500;

    /**
     * @Message("Excel参数错误")
     */
    public const PARAMETER_ERROR = 4007;
}
