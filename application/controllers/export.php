<?php
    class export extends Admin_Controller {

        public function __construct() {
            parent::__construct();
            $this->load->model('product_m');
            $this->load->model('product_detail_m');
            $this->load->model('product_type_m');
            $this->load->library('excel');
        }

        public function index() {
            $product_type = $this->product_type_m->get_by('status = 1');
            foreach ($product_type as $val) {
                if($val->parent==0)
                {
                    $sheet_array[] = $val;
                }
                $list_detail_types[$val->id] = $val;
            }
            $list_products = $this->product_m->get_by("status = 1");
            foreach ($sheet_array as $k => $type) {
                $list_type_childs = $this->get_all_childs_of_type($product_type,$type->id);
                $list_type_childs[] = $type->id;
                
                //activate worksheet number 1
                $this->excel->createSheet($k);
                $this->excel->setActiveSheetIndex($k);
                //name the worksheet
                $this->excel->getActiveSheet()->setTitle($type->name);
                //set cell A1 content with some text
                $this->excel->getActiveSheet()->setCellValue('A1', 'Danh sách sản phẩm: ' . $type->name)
                ->setCellValue('A2','STT')
                ->setCellValue('B2','Mã SP')
                ->setCellValue('C2','Tên SP')
                ->setCellValue('D2','Loại SP')
                ->setCellValue('E2','Số lượng')
                ->setCellValue('F2','Giá')
                ->setCellValue('G2','Giá bán');
                $row=3;
                foreach ($list_products as $key => $val) {
                    if(in_array($val->product_type,$list_type_childs))
                    {
                        $detail = $this->product_detail_m->get_detail_item($val->id,TRUE);
                        $type = $list_detail_types[$val->product_type];
                        $this->excel->getActiveSheet()->setCellValue('A'.$row, $row-2);
                        $this->excel->getActiveSheet()->setCellValue('B'.$row, $detail->model);
                        $this->excel->getActiveSheet()->setCellValue('C'.$row, $val->name);
                        $this->excel->getActiveSheet()->getCell("C".$row)->getHyperlink()->setUrl(strip_tags(build_link_to_detail_product($val->slug,$val->id)));
                        $this->excel->getActiveSheet()->setCellValue('D'.$row, $type->name);
                        $this->excel->getActiveSheet()->setCellValue('E'.$row, $val->number);
                        $this->excel->getActiveSheet()->setCellValue('F'.$row, $val->price);
                        $this->excel->getActiveSheet()->setCellValue('G'.$row, $val->sale_price);
                        $row++;
                    }
                }
                //change the font size
                $this->excel->getActiveSheet()->getStyle('A1')->getFont()->setSize(20);
                $this->excel->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);
                $this->excel->getActiveSheet()->getStyle('A2')->getFont()->setSize(14);
                $this->excel->getActiveSheet()->getStyle('A2')->getFont()->setBold(true);
                $this->excel->getActiveSheet()->getStyle('B2')->getFont()->setSize(14);
                $this->excel->getActiveSheet()->getStyle('B2')->getFont()->setBold(true);
                $this->excel->getActiveSheet()->getStyle('C2')->getFont()->setSize(14);
                $this->excel->getActiveSheet()->getStyle('C2')->getFont()->setBold(true);
                $this->excel->getActiveSheet()->getStyle('D2')->getFont()->setSize(14);
                $this->excel->getActiveSheet()->getStyle('D2')->getFont()->setBold(true);
                $this->excel->getActiveSheet()->getStyle('E2')->getFont()->setSize(14);
                $this->excel->getActiveSheet()->getStyle('E2')->getFont()->setBold(true);
                $this->excel->getActiveSheet()->getStyle('F2')->getFont()->setSize(14);
                $this->excel->getActiveSheet()->getStyle('F2')->getFont()->setBold(true);
                $this->excel->getActiveSheet()->getStyle('G2')->getFont()->setSize(14);
                $this->excel->getActiveSheet()->getStyle('G2')->getFont()->setBold(true);
                //merge cell A1 until D1
                $this->excel->getActiveSheet()->mergeCells('A1:I1');
                //set aligment to center for that merged cell (A1 to D1)
                $this->excel->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                $this->excel->getActiveSheet()->getStyle('A2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                $this->excel->getActiveSheet()->getStyle('B2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                $this->excel->getActiveSheet()->getStyle('C2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                $this->excel->getActiveSheet()->getStyle('D2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                $this->excel->getActiveSheet()->getStyle('E2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                $this->excel->getActiveSheet()->getStyle('F2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                $this->excel->getActiveSheet()->getStyle('G2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            }

            $filename='maychuhanoi.xls'; //save our workbook as this file name
            header('Content-Type: application/vnd.ms-excel'); //mime type
            header('Content-Disposition: attachment;filename="'.$filename.'"'); //tell browser what's the file name
            header('Cache-Control: max-age=0'); //no cache

            //save it to Excel5 format (excel 2003 .XLS file), change this to 'Excel2007' (and adjust the filename extension, also the header mime type)
            //if you want to save it as .XLSX Excel 2007 format
            $objWriter = PHPExcel_IOFactory::createWriter($this->excel, 'Excel5');  
            ob_end_clean();
            //force user to download the Excel file without writing it to server's HD
            $objWriter->save('php://output');
        }

        private function get_all_childs_of_type ($list_type,$parent=0){
            $data = $this->get_type_child($list_type,$parent);
            foreach ($data as $val) {
                $array = $this->get_all_childs_of_type($list_type,$val);
                $data = array_merge_recursive($data,$array);
            }
            return $data;
        }

        private function get_type_child($list_type,$parent){
            $data = array();
            foreach ($list_type as $val) {
                if($val->parent==$parent)
                    $data[] = $val->id;
            }
            return $data;
        }
}