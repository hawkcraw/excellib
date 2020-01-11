<?php

namespace Yxf\Excellib;

//Excel操作类
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

/*
 * examples
 * $data = [
            ['ID','姓名',['value'=>'头像'],'性别',['value'=>'爱好','width'=>30],'very import things','哈哈哈哈哈哈哈哈',],
            [12,'张三',['image'=>'/images/height_img.png'],'男','唱歌、小牧、玩游戏','不知道什么重要','搜if将诶噢王炯房屋及诶哦',],
            [31,'栗色',['image'=>'https://www.baidu.com/'],'女','多少分、小我无法玩游戏','不知雾f将重要','搜if将诶噢王炯房屋及诶哦',],
            [35,'网二',['image'=>'https://ss0.bdstatic.com/70cFuHSh_Q1YnxGkpoWK1HF6hhy/it/u=2688892464,3753125996&fm=111&gp=0.jpg'],'男','唱歌、小牧、玩游戏','不知道什f将重要','搜if将诶噢诶噢王炯及诶哦',],
            [1,'列大声道',['image'=>'https://timgsa.baidu.com/timg?image&quality=80&size=b9999_10000&sec=1514974949837&di=3d3ba7d3f9c03d6c88dd067f21e6253a&imgtype=0&src=http%3A%2F%2Fpic46.nipic.com%2F20140815%2F14008695_152008796000_2.jpg'],'男','唱歌、小牧、玩游戏','不f将什么重要','搜噢王炯房屋if将诶噢王炯房屋及诶哦',],
            [4,'广东省',['image'=>'images/height_img.png'],'','唱歌、小牧什么玩游戏','不知道f将重要','搜if将诶噢王炯房屋及诶哦',],
            [134,'多少分',['image'=>'http://static.googleadsserving.cn/pagead/imgad?id=CICAgKDLv-vJngEQoAEY2AQyCH9yHR4LAHzw'],'位','唱歌、小牧、玩游戏','不知道什f将要','搜if将诶噢王炯房屋及诶哦',],
        ];
   $data1 = [
        ['id'=>1,'name'=>'张三','gender'=>'男'],
        ['id'=>2,'name'=>'李文','gender'=>'女'],
    ];
    $path = excelPower::saveLocal($data,'uploads/bac.xls');//保存本地
    excelPower::download($data,'bac.xls');//直接下载
    excelPower::saveCsv($data,'bac.xls');//直接下载
    excelPower::saveCsv(excelPower::arrMap2CsvArr($data1),'bac.xls');//直接下载
 *
 *
 *
 *
 */
class excelPower
{

    public static $index = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z',
        'AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ',
        'BA','BB','BC','BD','BE','BF','BG','BH','BI','BJ','BK','BL','BM','BN','BO','BP','BQ','BR','BS','BT','BU','BV','BW','BX','BY','BZ',
        'CA','CB','CC','CD','CE','CF','CG','CH','CI','CJ','CK','CL','CM','CN','CO','CP','CQ','CR','CS','CT','CU','CV','CW','CX','CY','CZ',
    ];

    public static $_useTpl = false;
    //下载
    public static function download($data, $filename, $rowHeight=80, $type='Xlsx', $useTpl=false){

        $spreadsheet = self::createExcel($data,$rowHeight, $useTpl);

// Redirect output to a client’s web browser (Xls)
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="'.$filename.'.'.strtolower($type).'"');
        header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
        header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
        header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header('Pragma: public'); // HTTP/1.0

        $writer = IOFactory::createWriter($spreadsheet, $type);
        $writer->save('php://output');
        exit;
    }

    public static function saveLocal($data,$filename,$rowHeight=80,$type='Xlsx',$useTpl=true){
        $spreadsheet = self::createExcel($data,$rowHeight,$useTpl);
        $writer = IOFactory::createWriter($spreadsheet, $type);
        if(!is_dir(dirname($filename))) @mkdir(dirname($filename),0777,true);
        $writer->save($filename);
        if(file_exists($filename)){
            return $filename;
        }
        return false;
    }

    public static function saveCsv($data,$filename,$path=null,$type='download'){
        if (count($data) == 0) {
            return null;
        }
        $path = $path ? $path : sys_get_temp_dir().DIRECTORY_SEPARATOR;

        if(!is_dir($path)) @mkdir($path,0777,true);

        $file = $path.DIRECTORY_SEPARATOR.$filename.'.csv';
        $df = fopen($file, 'w');
        $header = array_shift($data);
        fputcsv($df, array_values($header));
        foreach ($data as $row) {
            fputcsv($df, $row);
        }
        fclose($df);
        if($type=='download'){
            header('Content-Type: application/octet-stream');
            header("Content-Transfer-Encoding: Binary");
            header("Content-disposition: attachment; filename=\"" . basename($file) . "\"");
            readfile($file);
            die;
        }

        return $file;
    }

    public static function saveBigCsv($func,$filename,$path=null,$type='download'){

        $path = $path ? $path : sys_get_temp_dir().DIRECTORY_SEPARATOR;
        $file = $path.$filename.'_'.date('YmdHis').'.csv';
        if(!is_dir($path)) @mkdir($path,0777,true);
        $fp = fopen($file, 'w');

        $func($fp);

        fclose($fp);

        if($type!='download') return $file;
        header('Content-Type: application/octet-stream');
        header("Content-Transfer-Encoding: Binary");
        header("Content-disposition: attachment; filename=\"" . basename($file) . "\"");
        readfile($file);
        die;
    }

    public static function getCsvData($file_path,$allow_fields=[]){

        if(!is_file($file_path)) return ['error'=>1,'msg'=>'文件不存在'];
        $file = fopen($file_path,'r');
        $list = [];

        while ($data = fgetcsv($file)) {
            $list[] = $data;
        }
        $header = array_map(array(__CLASS__, 'removeBomUtf8'),array_shift($list));
        if(!empty($allow_fields) && !empty(array_diff($allow_fields,$header))){
            return ['error'=>1, 'msg'=> "导入格式有误，请查看【demo】"];
        }
        $list = array_reduce($list,function($result, $item)use($header){
            array_push($result,array_combine($header,$item));
            return $result;
        },[]);
        return $list;
    }

    public static function arrMap2CsvArr($data){
        if(empty($data)) return [];
        $header = array_keys(reset($data));
        $data = array_map('array_values',$data);
        array_unshift($data,$header);
        return $data;
    }

    protected static function createExcel($data, $rowHeight=80, $useTpl=false){

        if($useTpl){
            self::$_useTpl = $useTpl;
            $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
            $spreadsheet = $reader->load(__DIR__."/template/data.xlsm");
        }else{
            $spreadsheet = new Spreadsheet();
        }

// Set document properties
//        $spreadsheet->getProperties()->setCreator('Maarten Balliauw')
//            ->setLastModifiedBy('Maarten Balliauw')
//            ->setTitle('Office 2007 XLSX Test Document')
//            ->setSubject('Office 2007 XLSX Test Document')
//            ->setDescription('Test document for Office 2007 XLSX, generated using PHP classes.')
//            ->setKeywords('office 2007 openxml php')
//            ->setCategory('Test result file');

// Add some data
        $currentSheet = $spreadsheet->setActiveSheetIndex(0);
        foreach ($data as $row => $row_data){
            foreach($row_data as $col => $val){
                self::setCellValue(self::$index[$col].($row+1),$val,$currentSheet,$col,$row);
            }
            $currentSheet->getRowDimension($row+1)->setRowHeight($rowHeight);
        }

// Rename worksheet
        $spreadsheet->getActiveSheet()->setTitle('Sheet1');
        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
//        $spreadsheet->setActiveSheetIndex(0);
        return $spreadsheet;
    }

    protected static function setCellValue($cell,$val,&$currentSheet,$col='',$row=''){
        //下面一行代码有问题，还没仔细看
        if($row===0) {
            $currentSheet->getColumnDimension(self::$index[$col])->setWidth(20);
        }
        if(is_array($val)){
            if(!empty($val['image'])){
                self::setImageCell($cell,$val['image'],$currentSheet);
            }
            if(!empty($val['value'])){
                $currentSheet->setCellValue($cell, $val['value']);
            }
            if(!empty($val['width'])){
                $currentSheet->getColumnDimension(self::$index[$col])->setWidth($val['width']);
            }
            if(!empty($val['row'])){
                //$currentSheet->mergeCells('A18:E22');
                $currentSheet->mergeCells($cell.':'.self::$index[$col].($row+$val['row']+1));
            }
            if(!empty($val['col'])){
                $currentSheet->mergeCells($cell.':'.self::$index[$col+$val['col']].($row+1));
            }
        }else{
            $currentSheet->setCellValue($cell, $val);
        }
        $currentSheet->getStyle($cell)->getAlignment()->setWrapText(true);
    }

    private static function setImageCell($cell,$image_url,&$currentSheet){
        if(self::$_useTpl){
            $currentSheet->setCellValue($cell,$image_url);
            return;
        }
        $drawing = new Drawing();
        $drawing->setName('Image');
        $drawing->setDescription('');
        $path = self::getImagePath($image_url);
        if(!file_exists($path)) return false;
        $drawing->setPath($path);

        $drawing->setCoordinates($cell);
        $drawing->setOffsetX(6);                       //setOffsetX works properly
        $drawing->setOffsetY(6);
        $drawing->setWidthAndHeight(100,100);
        $drawing->setWorksheet($currentSheet);
//        $currentSheet->getColumnDimension(substr($cell,0,1))->setWidth(1000);
//        $currentSheet->getRowDimension(substr($cell,1))->setRowHeight(60);
    }
    protected static function getImagePath($url){
        $filename = md5($url).'.'.pathinfo($url,PATHINFO_EXTENSION);
        $path = getcwd().'/downloads/excel_image/';
        if(file_exists($path.$filename)) return $path.$filename;
        if(strpos($url,'http')===0){
            if(!empty(getimagesize($url))){
                if(!is_dir($path)) @mkdir($path,0777,true);
                @copy($url,$path.$filename);
                return $path.$filename;
            }
            return '';
        }
        return ltrim($url,'/');
    }

    protected static function removeBomUtf8($s){
        if(substr($s,0,3)==chr(hexdec('EF')).chr(hexdec('BB')).chr(hexdec('BF'))){
            return substr($s,3);
        }else{
            return $s;
        }
    }
}
