<?php

declare(strict_types=1);

namespace HyperfTest\Cases;

use Ezijing\HyperfExcel\FastExcel;

/**
 * @internal
 * @coversNothing
 */
class ExampleTest extends AbstractTestCase
{
    public function testExample()
    {
        $list = collect([
            [ 'id' => 1, 'name' => 'Jane' ],
            [ 'id' => 2, 'name' => 'John' ],
        ]);

        (new FastExcel($list))->export(__DIR__. '/file.xlsx');

        (new FastExcel($list))->export(__DIR__. '/invoices.csv');

        (new FastExcel($list))->export(__DIR__. '/file11.xlsx', function ($data) {
            return [
                'ID' => $data['id'],
                '姓名' => $data['name'],
            ];
        });

        $data = (new FastExcel($list))->download(__DIR__. '/file.xlsx');
        var_dump($data);
    }
}
