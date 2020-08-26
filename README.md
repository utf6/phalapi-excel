PahlApi2.x 接口框架利用 PHPExcel 处理 Excel 文件

### 附上:

PhalApi 官网地址： [http://www.phalapi.net/](http://www.phalapi.net/ "PhalApi官网")

开源中国Git地址：[http://git.oschina.net/dogstar/PhalApi/tree/release](http://git.oschina.net/dogstar/PhalApi/tree/release "开源中国Git地址")

开源中国拓展Git地址：[http://git.oschina.net/dogstar/PhalApi-Library](http://git.oschina.net/dogstar/PhalApi-Library "开源中国Git地址")

### 1、安装

可以直接在 composer.json 文件中添加

    "require": {
        "utf6/phalapi-excel" : "*"
    },
    
或者直接使用 composer 安装    

    composer require utf6/phalapi-excel

### 2、初始化

在 di.php 加入
    
    $di->excel = function() {
        return new \utf6\phalapiExcel\Lite();
    };

### 3、使用

PhalApi-PHPExcel提供两个基础封装好的方法分别是 exportExcel、importExcel 分别处理导入、导出功能。

exportExcel 接受三个参数，$data基础数据，$headArr标题，$filename 文件名称。

下面是一个例子

    $data = [
        ['name' => '张三', 'password' => 'qwa3la'],
        ['name' => '李四', 'password' => 'vdf45s']
    ];
    
    $filename    = "用户信息.xlsx";
    $headArr     = array("用户名", "密码");
    
    \PhalApi\DI()->excel->exportExcel($filename, $data, $headArr);
        
PhalApi-PHPExcel可根据导出的文件后缀来导出不同格式的Excel文档

importExcel 接受三个参数：$filename 文件名称，$firstRowTitle 标题(可选默认从第一行作为标题)，$Sheet 工作表(默认第一张工作表)

    $rs = \PhalApi\DI()->excel->importExcel("./test.xlsx");

**当然 PHPExcel 是一个强大的工具可以通过$PHPExcel->getPHPExcel();获得完整的PHPExcel实例自由使用**

### 4、总结

希望此拓展能够给大家带来方便以及实用！