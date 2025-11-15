<?php


class LayoutTest extends PHPUnit\Framework\TestCase
{

    protected function setUp(): void
    {
        if (!defined('PHPEXCEL_ROOT')) {
            define('PHPEXCEL_ROOT', APPLICATION_PATH . '/');
        }
        require_once(PHPEXCEL_ROOT . 'PHPExcel/Autoloader.php');
    }

    public function testSetLayoutTarget()
    {
        $LayoutTargetValue = 'String';

        $testInstance = new PHPExcel_Chart_Layout;

        $result = $testInstance->setLayoutTarget($LayoutTargetValue);
        $this->assertTrue($result instanceof PHPExcel_Chart_Layout);
    }

    public function testGetLayoutTarget()
    {
        $LayoutTargetValue = 'String';

        $testInstance = new PHPExcel_Chart_Layout;
        $setValue = $testInstance->setLayoutTarget($LayoutTargetValue);

        $result = $testInstance->getLayoutTarget();
        $this->assertEquals($LayoutTargetValue, $result);
    }
}
