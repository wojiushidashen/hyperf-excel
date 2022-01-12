<?php

declare(strict_types=1);

if (! function_exists('app')) {
    function app()
    {
        return \Hyperf\Utils\ApplicationContext::getContainer();
    }
}
