<?php

require_once dirname(__FILE__) . '/../src/Lite.php';

class PHPUnderControll_Lite_Test extends PHPUnit_FrameWork_TestCase {

    public function testHere() {

        $data=array(
            array('username'=>'zhangsan','password'=>"123456"),
            array('username'=>'lisi','password'=>"abcdefg"),
            array('username'=>'wangwu','password'=>"111111"),
        );

        $filename    = "test_excel.xlsx";
        $headArr     = array("用户名", "密码");

        \PhalApi\DI()->execl->exportExcel($filename, $data, $headArr);
    }
}
