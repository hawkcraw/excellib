<?php
namespace Home\Logic;
/**
 *
 * excel操作类
 * 依附类库PHPExcel
 * @author yxf <561641083@qq.com>
 *
 */

class excelPower
{
    
    //将数据导出为excel
    public static function outputExcel($list,$filename = 'demo', $output = 'php://output') {
        $objPHPExcel = new \PHPExcel();
        
        //获取当前活动的表
        $objActSheet = $objPHPExcel->getActiveSheet();
        $objActSheet->getDefaultRowDimension()->setRowHeight(-1);

        $index = range('A', 'Z');
        $i = 1;
        foreach ($list as $key => $item) {
            $j = 0;
            foreach ($item as $k1 => $v1) {
                if(is_array($v1)){
                    switch (@$v1['type']){
                        case 'image':
                            self::setImageCell($index[$j], $i, $v1, $objActSheet);
                            break;
                    }

                }else{
                    //$objActSheet->setCellValueExplicit($index[$j] . $i, $v1, \PHPExcel_Cell_DataType::TYPE_STRING);
                    self::setCellValue($index[$j] . $i, $v1, $objPHPExcel);
                }

                
                //$objActSheet->setCellValue($index[$j] . $i, $v1);
                $j++;
            }
            $i++;
        }
        
        if($output === 'php://output'){//保存文件
            header ( 'Content-Type: application/vnd.ms-excel' );
            header ( 'Content-Disposition: attachment;filename="' . $filename . '.xls"' ); //"'.$filename.'.xls"
            header ( 'Cache-Control: max-age=0' );
        }
        
        $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
         //在内存中准备一个excel2003文件
        $objWriter->save($output);
    }

    //从excel导入数据
    public static function inputExcel($filename){
//        $objReader = \PHPExcel_IOFactory::createReader('Excel5');//use excel2007 for 2007 format
//        $objPHPExcel = $objReader->load($filename);
        $objPHPExcel = \PHPExcel_IOFactory::load($filename);//support xls / xlsx
        $sheet = $objPHPExcel->getSheet(0); 
        $highestRow = $sheet->getHighestRow();           //取得总行数 
        $highestColumn = $sheet->getHighestColumn(); //取得总列数

        //循环读取excel文件,读取一条,插入一条
        $data = null;
        for($j=1;$j<=$highestRow;$j++)                        //从第一行开始读取数据
        { 
            for($k='A',$i=0;$k<=$highestColumn;$k++,$i++)            //从A列读取数据
            { 
                //
                // 这种方法简单，但有不妥，以''合并为数组，再分割为字段值插入到数据库
                // 实测在excel中，如果某单元格的值包含了导入的数据会为空        
                //
                $data[$j][$i]=$objPHPExcel->getActiveSheet()->getCell("$k$j")->getValue().'';//读取单元格
            } 
            
     
        }  
        return $data;
    }

    //设置图片单元格
    private static function setImageCell($col, $row, $v, &$objActSheet){
        // 图片生成
        if(!file_exists($v['path'])) return ;
        $objDrawing[$row] = new \PHPExcel_Worksheet_Drawing();
        $objDrawing[$row]->setPath($v['path']);
        // 设置宽度高度
        if(!empty($v['width'])) $objActSheet->getColumnDimension($col)->setWidth($v['width']);
        $objActSheet->getRowDimension($row)->setRowHeight(70);
        $objDrawing[$row]->setHeight(80);//照片高度
        $objDrawing[$row]->setWidth(80); //照片宽度
        /*设置图片要插入的单元格*/
        $objDrawing[$row]->setCoordinates($col.$row);
        // 图片偏移距离
        $objDrawing[$row]->setOffsetX(1);
        $objDrawing[$row]->setOffsetY(2);
        $objDrawing[$row]->setWorksheet($objActSheet);
    }

    private function setCellValue($cell,$val, &$objPHPExcel){

        $objPHPExcel->getActiveSheet(0)->setCellValue($cell, trim($val));

        if(is_numeric(trim($val)) && strlen(trim($val)) <= 15){//excel只能显示15数字
            $number = trim($val);
            if(strpos($number,'.')>0){
                $objPHPExcel->getActiveSheet(0)->getStyle($cell)
                    ->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
            }else{
                $objPHPExcel->getActiveSheet(0)->getStyle($cell)
                    ->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
            }

        }
    }
}

//=================demo==================//
//输出excel
// $data = [
//     ['黎明',28,'唱歌'],
//     ['张三',38,'画画'],
//     ['李四',128,'吃饭'],
//     ['王武',8,'睡觉'],
// ];
// array_unshift($data, ['姓名','年龄','任务']);
//excelPower::outputExcel($data,'人员');//输出到浏览器下载
//excelPower::outputExcel($data,'人员','./'.date('YmdHis').'.xls');//保存到本地 注：需要有操作目录权限


//导入excel
// $data = excelPower::inputExcel('20151224180030.xls');
// var_dump($data);
?>