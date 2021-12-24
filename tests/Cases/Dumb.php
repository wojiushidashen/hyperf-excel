<?php

namespace HyperfTest\Cases;


use Hyperf\Database\Model\Model;

/**
 * Class Dumb.
 */
class Dumb extends Model
{
    public $data;

    /**
     * Dumb constructor.
     *
     * @param $data
     */
    public function __construct($data)
    {
        parent::__construct();
        $this->data = $data;
    }
}
