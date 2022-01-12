<?php

declare(strict_types=1);

namespace Ezijing\HyperfExcel\Core\Services;

use Ezijing\HyperfExcel\Core\Constants\ErrorCode;
use Ezijing\HyperfExcel\Core\Exceptions\ExcelException;
use Hyperf\Di\Annotation\Inject;
use Hyperf\Validation\Contract\ValidatorFactoryInterface;

/**
 * 验证器.
 */
class Validator
{
    /**
     * @Inject
     * @var ValidatorFactoryInterface
     */
    public $validatorFactory;

    public function verify(array $params, array $validationParams, array $noticeMessage = [])
    {
        $inputData = [];
        foreach ($params as $paramName => $paramValue) {
            if (! is_null($paramValue)) {
                $inputData[$paramName] = $paramValue;
            }
        }
        $validator = $this->validatorFactory->make($inputData, $validationParams, $noticeMessage);
        if ($validator->fails()) {
            $errorMessage = $validator->errors()->getMessages()[0];

            throw new ExcelException(ErrorCode::PARAMETER_ERROR, $errorMessage);
        }

        return true;
    }
}
