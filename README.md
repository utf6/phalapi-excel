## PahlApi2.x 接口框架利用 PHPExcel 处理Excel 文件

##前言

在之前有提供过一个PhalApi-Excel拓展来提供导出,但是还是有一些不方便,比如没有导入,导出文件体积过大此类问题,也是应为自身业务有使用到重新来提供一个PHPExcel的拓展
来提供更为强大的Excel处理功能

附上:

官网地址:[http://www.phalapi.net/](http://www.phalapi.net/ "PhalApi官网")

开源中国Git地址:[http://git.oschina.net/dogstar/PhalApi/tree/release](http://git.oschina.net/dogstar/PhalApi/tree/release "开源中国Git地址")

开源中国拓展Git地址:[http://git.oschina.net/dogstar/PhalApi-Library](http://git.oschina.net/dogstar/PhalApi-Library "开源中国Git地址")

##1. 安装

    "require": {
        "utf6/phalapi-excel" : "*"
    },

##2. 初始化

    在di.php加入
    $di->execl = function() {
        return new \utf6\phalapiExcel\Lite();
    };

##3. 使用

PhalApi-PHPExcel提供两个基础封装好的方法分别是exportExcel,importExcel分表接触导出和接受的问题

exportExcel接受三个参数,$data基础数据,$headArr标题,$filename文件名称下面是一个例子

       $data=array(
            array('username'=>'zhangsan','password'=>"123456"),
            array('username'=>'lisi','password'=>"abcdefg"),
            array('username'=>'wangwu','password'=>"111111"),
        );

        $filename    = "test_excel.xlsx";
        $headArr     = array("用户名", "密码");
        
        \PhalApi\DI()->execl->exportExcel($filename, $data, $headArr);
        
PhalApi-PHPExcel可根据导出的文件后缀来导出不同格式的Excel文档

importExcel接受三个参数,$filename文件名称,$firstRowTitle标题(可选默认从第一行作为标题),$Sheet工作表(默认第一张工作表)

    $rs = \PhalApi\DI()->execl->importExcel("./test.xlsx");

**当然PHPExcel是一个强大的工具可以通过$PHPExcel->getPHPExcel();获得完整的PHPExcel实例自由使用**

##4. 总结

希望此拓展能够给大家带来方便以及实用,如有其他童鞋希望能加入其余常用功能可与笔者进行联系!

注:笔者能力有限有说的不对的地方希望大家能够指出,也希望多多交流!