<?php

/**
 * @link https://github.com/PHPOffice/PHPExcel/tree/1.8/Classes
 */
include_once 'PHPExcel.php';

class ExcelSimpleDecorator
{
    protected $objPHPExcel;
    protected $objReader;
    protected $fileName;
    protected $fileType;

    public function __construct($fileName, $fileType)
    {
        $objReader = PHPExcel_IOFactory::createReader($fileType);
        $objReader->setReadDataOnly(true);
        $this->objReader = $objReader;
        $this->fileName = $fileName;
        $this->fileType = $fileType;
    }

    public function getValue($name)
    {
        $item = $this->existValue($name);
        if(!$item){
            $this->addValue([$name]);
        }
        return $item;
    }

    protected function fileExists()
    {
        return file_exists($this->fileName);
    }

    protected function addValue($data)
    {
        $objPHPExcel = $this->getObj();
        $worksheet = $objPHPExcel->getActiveSheet();
        $num_rows = $objPHPExcel->getActiveSheet()->getHighestRow();
        $worksheet->fromArray($data, NULL, 'A'.($num_rows + 1));
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, $this->fileType);
        $objWriter->save($this->fileName);
    }

    public function getPrice($name)
    {
        $item = $this->getValue($name);
        if($item && array_key_exists(1, $item)){
            return $item[1];
        }
        return null;
    }

    protected function getObj()
    {
        if(!$this->objPHPExcel){
            $this->objPHPExcel = ($this->fileExists()) ? $this->objReader->load($this->fileName) : new PHPExcel();
        }
        return $this->objPHPExcel;
    }

    protected function existValue($name)
    {
        foreach($this->toArray() as $item){
            if(isset($item[0]) && $item[0] == $name){
                return $item;
            }
        }
        return null;
    }

    public function toArray()
    {
        $objPHPExcel = $this->getObj();
        return $objPHPExcel->getActiveSheet()->toArray();
    }
}
//$simple = new ExcelSimpleDecorator('test.xls', 'Excel5');
//var_dump($simple->getPrice('Test item на русском'), $simple->toArray());