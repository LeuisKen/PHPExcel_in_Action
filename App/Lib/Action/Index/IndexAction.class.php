<?php
/**
 * leuisken.github.io  IndexAction.class.php
 * ============================================================================
 * 版权所有 (C) 2015 星辰工作室
 * 作者:   魏嘉汛
 * Email:  leuisken@gmail.com
 * ----------------------------------------------------------------------------
 * 许可声明：这是一个开源程序，未经许可不得将本软件的整体或任何部分用于商业用途及再发布。
 * ============================================================================
 * $Author: 魏嘉汛 (leuisken@gmail.com) $
 * $Date: 2015-03-29 上午00:47 $
 */
class IndexAction extends Action{
	public function index(){
		$this->display();
	}

	//上传类
	public function upload(){
		//导入上传类
		import('ORG.Net.UploadFile');
		//初始化上传类并进行一些配置
		$upload = new UploadFile();
		$upload->maxSize = 3145728;
		$upload->allowExts = array('xls', 'xlsx');
		$upload->savePath = __ROOT__ . '/Public/Uploads/';
		//上传文件并返回文件信息
		$info = $upload->uploadOne($_FILES['excelData']);
		//保存文件路径
		$filename = $info[0]['savepath'] . $info[0]['savename'];
		//保存文件扩展名
		$extension = $info[0]['extension'];
		//判断是否上传成功，若否，抛出对应错误
		if(!$info){
			$this->error($upload->getErrorMsg());
		}else{
			$this->read_excel($filename, $extension);
		}
	}

	//excel文件读取类
	private function read_excel($filename, $extension = 'xls'){
		//导入PHPExcel类库
		import('ORG.Util.PHPExcel');
		$PHPExcel = new PHPExcel();
		//根据扩展名判断所用的解析方式
		if($extension == 'xls'){
			import("ORG.Util.PHPExcel.Reader.Excel5");
			$PHPReader = new PHPExcel_Reader_Excel5();
		}else if($extension == 'xlsx'){
			import("ORG.Util.PHPExcel.Reader.Excel2007");
			$PHPReader = new PHPExcel_Reader_Excel2007();
		}
		//读取excel文件
		$PHPExcel = $PHPReader->load($filename);
		//获取Sheet0中的数据
		$currentSheet = $PHPExcel->getSheet(0);
		//获取行数、列数
		$allColumn = $currentSheet->getHighestColumn();
		$allRow = $currentSheet->getHighestRow();
		//循环遍历excel表格，重点是getCell和getValue方法
		for($currentRow = 1; $currentRow <= $allRow; $currentRow++){
			for($currentColumn = 'A'; $currentColumn <= $allColumn; $currentColumn++){
				$address = $currentColumn . $currentRow;
				$data[$currentRow][$currentColumn] = $currentSheet->getCell($address)->getValue();
			}
		}
		//循环所得数据插入数据库，至于从2开始嘛，执行var_dump($data)你就懂了
		// echo "<pre>";
		// var_dump($data);
		for ($i=2; $i < count($data); $i++) {
			$insert = array(
				'username' => $data[$i]['A'],
				'stu_id' => $data[$i]['B'],
				'tel' => $data[$i]['C'],
				'qq' => $data[$i]['D']
				);
			$result = M('user')->data($insert)->add();
		}
		//然而我并没有想出该怎么判断成功和失败，感觉下面的方法并不好，虽然没出问题
		if($result){
			$this->success('Excel导入成功' ,'Index/Index/Index');
		}else{
			$this->error('Excel导入失败');
		}
	}

	//Excel输出类
	public function exports_excel(){
		//导入所需要的类库
		import('ORG.Util.PHPExcel');
		import('ORG.Util.PHPExcel.Writer.Excel5');
		import('ORG.Util.PHPExcel.IOFactory.php');
		//获取当前日期作为文件名
		$date = date("Y_m_d", time());
		$fileName = "$date" . ".xls";
		//初始化输出类并进行配置
		$objPHPExcel = new PHPExcel();
		$objProps = $objPHPExcel->getProperties();
		//尤其是这个，设置对拿个Sheet输出
		$objPHPExcel->setActiveSheetIndex(0);
		$objActSheet = $objPHPExcel->getActiveSheet();
		//对输出Excel设置表头，内容因具体项目而定
		$objActSheet->setCellValue('A1', '姓名');
		$objActSheet->setCellValue('B1', '学号');
		$objActSheet->setCellValue('C1', '联系电话');
		$objActSheet->setCellValue('D1', 'QQ');
		//获取数据库中的信息条目数，并取出
		$count = M('user')->count();
		$cell = M('user')->order('id')->select();
		//循环遍历，添加到Excel中，重点是setCellValue方法
		for($i = 1; $i<= $count; $i++){
			$objActSheet->setCellValue('A'.($i+1), $cell[$i-1]['username']);
			$objActSheet->setCellValue('B'.($i+1), $cell[$i-1]['stu_id']);
			$objActSheet->setCellValue('C'.($i+1), $cell[$i-1]['tel']);
			$objActSheet->setCellValue('D'.($i+1), $cell[$i-1]['qq']);
		}
		//文件名转换为gb2312编码，不然Windows下会乱码
		$fileName = iconv("utf-8", "gb2312", $fileName);
		//自定义http请求的返回内容
		header('Content-Type: application/vnd.ms-excel');
		header("Content-Disposition: attachment;filename=\"$fileName\"");
		header('Cache-Control: max-age=0');
		//将抽象的PHPExcel对象输出为Excel文件
		$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
		$objWriter->save('php://output');
		exit;
	}
}

?>