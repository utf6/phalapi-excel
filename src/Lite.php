<?php
namespace utf6\phalapiExcel ;
/**
 * PHPExcel
 * Copyright (c) 2006 - 2013 PHPExcel
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 * @category   PHPExcel
 * @package    PHPExcel
 * @copyright  Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    1.7.9, 2013-06-02
 */

/** PHPExcel root directory */
if (!defined('PHPEXCEL_ROOT')) {
    define('PHPEXCEL_ROOT', dirname(__FILE__) . '/');
    require(PHPEXCEL_ROOT . 'PHPExcel/Autoloader.php');
}

class Lite {

    public $PHPExcel;

    public function __construct() {

        $this->PHPExcel = new \PHPExcel();
    }

    public function getPHPExcel() {

        return $this->PHPExcel;
    }

    /**
     * 导入excel表格
     * @param string $fileName
     * @param array $keys
     * @param int $Sheet
     * @return array|string
     * @throws \PHPExcel_Exception
     * @throws \PHPExcel_Reader_Exception
     */
    public function importExcel($fileName, $keys=[], $Sheet = 0) {
        // 读取excel文件
        try {
            //依据类型读取
            $excelType = self::getExcelType($fileName);

            $PHPExcel = \PHPExcel_IOFactory::createReader($excelType);

            $PHPExcel->setReadDataOnly(true); // 只需要添加这个方法实现表格数据格式转换

            $PHPExcel = $PHPExcel->load($fileName);
        } catch ( Exception $e ) {
            return $e->getMessage();
        }

        //获取表中的第一个工作表，如果要获取第二个，把 0 改为 1 依次类推
        $data = $PHPExcel->getSheet($Sheet)->toArray();
        array_shift($data);

        if ($keys){
            if (count($data[0]) != count($keys)){
                return "字段于数据数量不一致";
            }

            foreach ($data as $i => $item){
                foreach ($item as $k => $v){
                    $tmp[$keys[$k]] = $v;
                }
                $data[$i] = $tmp;
            }
        }

        return $data;
    }

    /**
     * 获取索引
     * @param $arr
     * @param $key
     * @param string $default
     * @return string
     */
    public function getIndex($arr, $key, $default = '') {

        return isset($arr[$key]) ? $arr[$key] : $default;
    }

    /**
     * 导出数据
     * @param $fileName
     * @param $data
     * @param $headArr
     * @param string $type
     * @throws \PHPExcel_Exception
     * @throws \PHPExcel_Reader_Exception
     * @throws \PHPExcel_Writer_Exception
     */
    public function exportExcel($fileName, $data, $headArr, $type="vnd.ms-excel") {

        //对数据进行检验
        if (empty($data) || !is_array($data)) {
            die("data must be a array");
        }
        //检查文件名
        if (empty($fileName)) {
            exit;
        }

        $objPHPExcel = $this->PHPExcel;

        //设置表头
        $key = ord("A");//A--65
        $key2 = ord("@");

        foreach ($headArr as $v) {
            if($key>ord("Z")){
                $key2 += 1;
                $key = ord("A");
                $colum = chr($key2).chr($key);//超过26个字母时才会启用  dingling 20150626
            }else {
                if ($key2 >= ord("A")) {
                    $colum = chr($key2) . chr($key);
                } else {
                    $colum = chr($key);
                }
            }

            $objPHPExcel->setActiveSheetIndex(0)->setCellValue( $colum . '1', $v);

            $key += 1;
        }

        $column      = 2;
        $objActSheet = $objPHPExcel->getActiveSheet();
        foreach ($data as $key => $rows) { //行写入

            $span = ord("A");
            $span2 = ord("@");

            foreach ($rows as $keyName => $value) {// 列写入

                if($span>ord("Z")){
                    $span2 += 1;
                    $span = ord("A");
                    $j = chr($span2).chr($span);//超过26个字母时才会启用  dingling 20150626
                }else{
                    if($span2 >= ord("A")){
                        $j = chr($span2).chr($span);
                    }else{
                        $j = chr($span);
                    }
                }

                $objActSheet->setCellValue($j . $column, $value);
                $objActSheet->getColumnDimension($j)->setAutoSize(true);

                $span++;
            }
            $column++;
        }

        $fileName = iconv("utf-8", "gb2312", $fileName);
        //设置活动单指数到第一个表,所以Excel打开这是第一个表
        $objPHPExcel->setActiveSheetIndex(0);
        header('Access-Control-Allow-Origin: *');
        header('Content-Type: application/json');
        header("Content-Disposition: attachment;filename=\"$fileName\"");
        header('Access-Control-Max-Age: 86400');    //缓存预检
        header('Cache-Control: max-age=0');

        $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, self::getExcelType($fileName));
        $objWriter->save('php://output'); //文件通过浏览器下载
        exit;
    }

    /**
     * 获取 excel 表格类型
     * @param $pFilename
     * @return string|null
     */
    public static function getExcelType($pFilename) {

        // First, lucky guess by inspecting file extension
        $pathinfo = pathinfo($pFilename);

        $extensionType = NULL;
        if (isset($pathinfo['extension'])) {
            switch (strtolower($pathinfo['extension'])) {
                case 'xlsx':            //	Excel (OfficeOpenXML) Spreadsheet
                case 'xlsm':            //	Excel (OfficeOpenXML) Macro Spreadsheet (macros will be discarded)
                case 'xltx':            //	Excel (OfficeOpenXML) Template
                case 'xltm':            //	Excel (OfficeOpenXML) Macro Template (macros will be discarded)
                    $extensionType = 'Excel2007';
                    break;
                case 'xls':                //	Excel (BIFF) Spreadsheet
                case 'xlt':                //	Excel (BIFF) Template
                    $extensionType = 'Excel5';
                    break;
                case 'ods':                //	Open/Libre Offic Calc
                case 'ots':                //	Open/Libre Offic Calc Template
                    $extensionType = 'OOCalc';
                    break;
                case 'slk':
                    $extensionType = 'SYLK';
                    break;
                case 'xml':                //	Excel 2003 SpreadSheetML
                    $extensionType = 'Excel2003XML';
                    break;
                case 'gnumeric':
                    $extensionType = 'Gnumeric';
                    break;
                case 'htm':
                case 'html':
                    $extensionType = 'HTML';
                    break;
                case 'csv':
                    // Do nothing
                    // We must not try to use CSV reader since it loads
                    // all files including Excel files etc.
                    break;
                default:
                    break;
            }
            return $extensionType;
        }
    }
}