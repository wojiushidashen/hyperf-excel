<?php

namespace HyperfTest\Cases;

use Box\Spout\Common\Entity\Style\Color;
use Box\Spout\Writer\Common\Creator\Style\StyleBuilder;
use Ezijing\HyperfExcel\FastExcel;
use Ezijing\HyperfExcel\SheetCollection;
use PHPUnit\Framework\TestCase;

/**
 * Class FastExcelTest.
 */
class FastExcelTest extends TestCase
{
    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     */
    public function testImportXlsx()
    {
        $collection = (new FastExcel())->import(__DIR__ . '/test1.xlsx');
        print_r($collection);
        $this->assertEquals($this->collection(), $collection);
    }

    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     */
    public function testImportCsv()
    {
        $original_collection = $this->collection();

        $collection = (new FastExcel())->import(__DIR__.'/test2.csv');
        print_r($collection);
        $this->assertEquals($original_collection, $collection);

        $collection = (new FastExcel())->configureCsv(';')->import(__DIR__.'/test1.csv');
        print_r($collection);
        $this->assertEquals($original_collection, $collection);
    }

    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    private function export($file)
    {
        $original_collection = $this->collection();

        (new FastExcel(clone $original_collection))->export($file);
        $this->assertEquals($original_collection, (new FastExcel())->import($file));
        unlink($file);
    }

    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     */
    public function testExportXlsx()
    {
        $this->export(__DIR__.'/test2.xlsx');
    }

    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     */
    public function testExportCsv()
    {
        $this->export(__DIR__.'/test3.csv');
    }

    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     */
    public function testExcelImportWithCallback()
    {
        $collection = (new FastExcel())->import(__DIR__.'/test1.xlsx', function ($value) {
            return [
                'test' => $value['col1'],
            ];
        });
        $this->assertEquals(
            collect([['test' => 'row1 col1'], ['test' => 'row2 col1'], ['test' => 'row3 col1']]),
            $collection
        );

        $collection = (new FastExcel())->import(__DIR__.'/test1.xlsx', function ($value) {
            return new Dumb($value['col1']);
        });
        $this->assertEquals(
            collect([new Dumb('row1 col1'), new Dumb('row2 col1'), new Dumb('row3 col1')]),
            $collection
        );
    }

    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    public function testExcelExportWithCallback()
    {
        (new FastExcel(clone $this->collection()))->export(__DIR__.'/test2.xlsx', function ($value) {
            return [
                'test' => $value['col1'],
            ];
        });
        $this->assertEquals(
            collect([['test' => 'row1 col1'], ['test' => 'row2 col1'], ['test' => 'row3 col1']]),
            (new FastExcel())->import(__DIR__.'/test2.xlsx')
        );
        unlink(__DIR__.'/test2.xlsx');
    }

    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    public function testExportMultiSheetXLSX()
    {
        $file = __DIR__.'/test_multi_sheets.xlsx';
        $sheets = new SheetCollection([clone $this->collection(), clone $this->collection()]);
        (new FastExcel($sheets))->export($file);
        $this->assertEquals($this->collection(), (new FastExcel())->import($file));
        unlink($file);
    }

    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    public function testImportMultiSheetXLSX()
    {
        $collections = [
            collect([['test' => 'row1 col1'], ['test' => 'row2 col1'], ['test' => 'row3 col1']]),
            $this->collection(),
        ];
        $file = __DIR__.'/test_multi_sheets.xlsx';
        $sheets = new SheetCollection($collections);
        (new FastExcel($sheets))->export($file);

        $sheets = (new FastExcel())->importSheets($file);
        $this->assertInstanceOf(SheetCollection::class, $sheets);

        $this->assertEquals($collections[0], collect($sheets->first()));
        $this->assertEquals($collections[1], collect($sheets->all()[1]));

        unlink($file);
    }

    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    public function testImportMultiSheetWithSheetNamesXLSX()
    {
        $collections = [
            'Sheet with name A' => collect([['test' => 'row1 col1'], ['test' => 'row2 col1'], ['test' => 'row3 col1']]),
            'Sheet with name B' => $this->collection(),
        ];
        $file = __DIR__.'/test_multi_sheets_with_sheets_names.xlsx';
        $sheets = new SheetCollection($collections);
        (new FastExcel($sheets))->export($file);

        $sheets = (new FastExcel())->withSheetsNames()->importSheets($file);
        $this->assertInstanceOf(SheetCollection::class, $sheets);

        $this->assertEquals($collections['Sheet with name A'], collect($sheets->get('Sheet with name A')));
        $this->assertEquals($collections['Sheet with name B'], collect($sheets->get('Sheet with name B')));

        unlink($file);
    }

    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    public function testExportWithHeaderStyle()
    {
        $original_collection = $this->collection();
        $style = (new StyleBuilder())
           ->setFontBold()
           ->setBackgroundColor(Color::YELLOW)
           ->build();
        $file = __DIR__.'/test-header-style.xlsx';
        (new FastExcel(clone $original_collection))
            ->headerStyle($style)
            ->export($file);
        $this->assertEquals($original_collection, (new FastExcel())->import($file));

        unlink($file);
    }

    protected function collection()
    {
        return collect([
            ['col1' => 'row1 col1', 'col2' => 'row1 col2'],
            ['col1' => 'row2 col1', 'col2' => ''],
            ['col1' => 'row3 col1', 'col2' => 'row3 col2'],
        ]);
    }
}