<?php

declare(strict_types=1);

namespace Ezijing\HyperfExcel\Core\Constants;

use Hyperf\Constants\AbstractConstants;

class ExcelConstant extends AbstractConstants
{
    /**
     * @Message("保存到本地")
     */
    public const SAVE_TO_A_LOCAL_DIRECTORY = 1;

    /**
     * @Message("下载到浏览器")
     */
    public const DOWNLOAD_TO_BROWSER = 2;

    /**
     * @Message("缓存到临时文件并下载到浏览器")
     */
    public const DOWNLOAD_TO_BROWSER_BY_TMP = 3;

    public static function getExportWayMap()
    {
        return [
            self::SAVE_TO_A_LOCAL_DIRECTORY => self::getMessage(self::SAVE_TO_A_LOCAL_DIRECTORY),
            self::DOWNLOAD_TO_BROWSER => self::getMessage(self::DOWNLOAD_TO_BROWSER),
            self::DOWNLOAD_TO_BROWSER_BY_TMP => self::getMessage(self::DOWNLOAD_TO_BROWSER_BY_TMP),
        ];
    }
}
