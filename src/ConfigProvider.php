<?php

declare(strict_types=1);

namespace Ezijing\HyperfExcel;

class ConfigProvider
{
    public function __invoke(): array
    {
        return [
            'dependencies' => [
            ],
            'commands' => [
            ],
            'annotations' => [
                'scan' => [
                    'paths' => [
                        __DIR__,
                    ],
                ],
            ],
            'publish' => [
                [
                    'id' => 'config',
                    'description' => '发布注解路由配置.',
                ],
            ],
        ];
    }
}
