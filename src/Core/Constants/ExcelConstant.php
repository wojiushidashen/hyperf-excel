<?php

declare(strict_types=1);

namespace Ezijing\HyperfExcel\Core\Constants;

use Hyperf\Constants\AbstractConstants;
use Hyperf\Constants\Annotation\Constants;

/**
 * @Constants
 */
class ExcelConstant extends AbstractConstants
{
    /**
     * @Message("保存到本地")
     */
    public const SAVE_TO_A_LOCAL_DIRECTORY = 1;

    /**
     * @Message("缓存到临时文件并下载到浏览器")
     */
    public const DOWNLOAD_TO_BROWSER_BY_TMP = 2;

    /**
     * @Message("本地导入")
     */
    public const THE_LOCAL_IMPORT = 1;

    /**
     * @Message("浏览器导入")
     */
    public const BROWSER_IMPORT = 2;

    /**
     * 导出方式列表.
     *
     * @return array
     */
    public static function getExportWayMap()
    {
        return [
            self::SAVE_TO_A_LOCAL_DIRECTORY => self::getMessage(self::SAVE_TO_A_LOCAL_DIRECTORY),
            self::DOWNLOAD_TO_BROWSER_BY_TMP => self::getMessage(self::DOWNLOAD_TO_BROWSER_BY_TMP),
        ];
    }

    /**
     * 导入方式列表.
     *
     * @return array
     */
    public static function getImportWayMap()
    {
        return [
            self::THE_LOCAL_IMPORT => self::getMessage(self::THE_LOCAL_IMPORT),
            self::BROWSER_IMPORT => self::getMessage(self::BROWSER_IMPORT),
        ];
    }
}
