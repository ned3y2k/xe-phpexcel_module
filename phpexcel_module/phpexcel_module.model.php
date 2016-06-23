<?php
/**
 * User: Kyeongdae
 * Date: 2016-06-08
 * Time: 오전 5:02
 */

class phpexcel_moduleModel extends phpexcel_module {
	/** @var array */
	private $header = array();
	/** @var array */
	private $resultSet;
	/** @var array */
	private $fieldConverter = array();

	public function setHeader(array $header = array()) { $this->header = $header; }

	/** @return array */
	public function getHeader() { return $this->header; }

	public function setFiledConverter(array $fieldConverter) { $this->fieldConverter = $fieldConverter; }
	public function getFieldConverter() { return $this->fieldConverter; }
	
	public function setResultSet(array $resultSet = array()) { $this->resultSet = $resultSet; }
	public function &getResultSet() { return $this->resultSet; }
	
	public function fetchAll($queryId, $pageVarName = 'page', $page = 1, $args = null) {
		$output = $this->fetch($queryId, $pageVarName, $page, $args);

		$this->resultSet = $output->data;
		if($output->page_navigation->total_page <= 1) {
			return;
		}

		$total_page = $output->page_navigation->total_page;
		for($i = 2; $i<=$total_page; $i++) {
			$output = $this->fetch($queryId, $pageVarName, $i, $args);
			$this->resultSet = array_merge($this->resultSet, $output->data);
		}
	}

	private function fetch($queryId, $pageVarName, $page, stdClass $args = null) {
		if(!$args)
			$args = new stdClass();

		$args->{$pageVarName} = $page;

		$output = executeQueryArray($queryId, $args);
		
		if($output->error)
			throw new RuntimeException($output->message);
		
		return $output;
	}

	public function resultSetDump() {
		header('content-type: text/plain; charset=utf-8');
		var_dump($this->resultSet);
	}
}