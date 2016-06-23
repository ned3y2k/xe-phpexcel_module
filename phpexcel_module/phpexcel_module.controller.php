<?php

/**
 * User: Kyeongdae
 * Date: 2016-06-08
 * Time: 오전 4:57
 */
class phpexcel_moduleController extends phpexcel_module {
	private $rowCount = 1;

	function init() { }

	public function printDownloadHttpHeader($fileName) {
		ob_clean();
		flush();

		$fileName = iconv('utf-8', 'euc-kr', $fileName);

		header("Pragma: public");
		header("Expires: 0");
		header('Cache-Control: must-revalidate');
		header("Content-Description: File Transfer");
		header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
		header("Content-Disposition: attachment; filename=\"$fileName\"");
		header("Content-Transfer-Encoding: binary");

	}
	
	public function writeToSheet(PHPExcel_Worksheet $sheet, array $fields, phpexcel_moduleModel $model) {
		$this->writeHeader($sheet, $model->getHeader());
		$this->writeResultSet($sheet, $fields, $model);

		$this->rowCount = 1;
	}

	public function printPHPOutStream(PHPExcel $excel) {
		$writer = new PHPExcel_Writer_Excel2007($excel);
		$writer->save('php://output');

		exit;
	}

	private function writeHeader(PHPExcel_Worksheet $sheet, array $header = array()) {
		if ($header) {
			$columnCount = 0;
			foreach ($header as $column) {
				$sheet->setCellValueByColumnAndRow($columnCount, $this->rowCount, $column);
				$columnCount++;
			}

			unset($columnCount);
			$this->rowCount++;
		}
	}

	private function writeResultSet(PHPExcel_Worksheet $sheet, array $fields, phpexcel_moduleModel $model) {
		$resultSet = $model->getResultSet();
		$fieldConverter = $model->getFieldConverter();
		
		foreach ($resultSet as $result) {
			$columnCount = 0;
			foreach ($fields as $field) {
				$val = $this->resolveValue($field, $fieldConverter, $result);

				$sheet->setCellValueByColumnAndRow($columnCount, $this->rowCount, $val);
				$columnCount++;
			}
			$this->rowCount++;
		}
	}

	/**
	 * @param $field
	 * @param $fieldConverter
	 * @param $result
	 *
	 * @return mixed
	 */
	private function resolveValue($field, $fieldConverter, $result) {
		$val = array_key_exists($field, $fieldConverter)
			? $fieldConverter[$field]($result)
			: $result->{$field};

		return $val;
	}
}