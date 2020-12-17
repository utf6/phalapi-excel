PahlApi2.x 接口框架利用 PHPExcel 处理 Excel 文件

### 附上:

PhalApi 官网地址： [http://www.phalapi.net/](http://www.phalapi.net/ "PhalApi官网")

### 1、安装

可以直接在 composer.json 文件中添加
```composer
"require": {
    "utf6/phalapi-excel" : "*"
},
```

或者直接使用 composer 安装    
```composer 
composer require utf6/phalapi-excel
```

### 2、初始化

在 di.php 加入
```composer 
$di->excel = function() {
    return new \utf6\phalapiExcel\Lite();
};
```

### 3、使用

PhalApi-PHPExcel 提供两个基础封装好的方法分别是 `exportExcel` 、`importExcel` 分别处理导入、导出功能。

exportExcel 接受4个参数，`$data` 基础数据，`$headArr`：标题，`$filename` ：文件名称，`$type` ：下载方式（默认 vnd.ms-excel，ajax 导出时为：json）。

下面是一个例子
```php 
$data = [
    ['name' => '张三', 'password' => 'qwa3la'],
    ['name' => '李四', 'password' => 'vdf45s']
];

$filename    = "用户信息.xlsx";
$headArr     = array("用户名", "密码");

\PhalApi\DI()->excel->exportExcel($filename, $data, $headArr, 'json');
```
        
PhalApi-PHPExcel 可根据导出的文件后缀来导出不同格式的Excel文档

importExcel 接受三个参数：$filename 文件名称，$keys 键名(选默为空， 可接受一个数组（比如数据库字段名）)，$Sheet 工作表(默认第一张工作表)

```php 
$data = \PhalApi\DI()->excel->importExcel("./test.xlsx");
//返回
$data = [
    [
        '张三',
        '男'
    ]
];
```
  
传递键名
```php
$keys = ['name', 'sex'];
$data = \PhalApi\DI()->excel->importExcel("./test.xlsx", $keys);
//返回
$data = [
    [
        'name' => '张三',
        'sex' => '男'
    ]
];
```
当然 PHPExcel 是一个强大的工具可以通过 `$PHPExcel->getPHPExcel()` 获得完整的 PHPExcel 实例自由使用！