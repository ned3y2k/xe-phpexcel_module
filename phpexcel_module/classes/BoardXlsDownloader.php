<?php
/**
 * User: Kyeongdae
 * Date: 2016-06-24
 * Time: 오전 2:56
 */
class BoardXlsDownloader {
	private $headers = array(
		'문서번호',
		'모듈번호',
		'분류번호',
		'언어코드',
		'공지사항여부',
		'제목',
		'내용',
		'조회수',
		'추천수',
		'비추천수',
		'댓글수',
		'사용자ID',
		'사용자명',
		'별명',
		'사용자번호',
		'이메일주소',
		'홈페이지',
		'문서 태그',
		'등록일',
		'최종갱신일',
		'IP주소',
	);
	private $fields = array(
		'document_srl',
		'module_srl',
		'category_srl',
		'lang_code',
		'is_notice',
		'title',
		'content',
		'readed_count',
		'voted_count',
		'blamed_count',
		'comment_count',
		'user_id',
		'user_name',
		'nick_name',
		'member_srl',
		'email_address',
		'homepage',
		'tags',
		'regdate',
		'last_update',
		'ipaddress',
	);

	/** @var stdClass */
	private $moduleInfo;
	/** @var phpexcel_moduleModel */
	private $model;
	/** @var string[] */
	private $eids;

	/**
	 * BoardXlsDownloader constructor.
	 *
	 * @param phpexcel_moduleModel $model
	 * @param                      $moduleSrl
	 * @param array                $eids
	 */
	public function __construct(phpexcel_moduleModel $model, $moduleSrl, array $eids = null) {
		$this->model = $model;
		$this->eids  = is_array($eids) ? $eids : array();
		$this->setModuleSrl($moduleSrl);
		$this->prepareHeadersAndFields();
		$model->setHeader($this->headers);
	}

	/**
	 * @param $moduleSrl
	 *
	 * @return BoardXlsDownloader
	 */
	public function setModuleSrl($moduleSrl) {
		/** @var moduleModel $oModuleModel */
		$oModuleModel = getModel('module');
		$moduleInfo   = $oModuleModel->getModuleInfoByModuleSrl($moduleSrl);
		if (!$moduleInfo)
			throw new InvalidArgumentException('해당모듈을 찾을수 없습니다.');

		$this->moduleInfo = $moduleInfo;

		return $this;
	}

	private function prepareHeadersAndFields() {
		$oDocumentModel = getModel('document');
		/** @var ExtraItem[] $extraVars */
		$extraVars = $oDocumentModel->getExtraKeys($this->moduleInfo->module_srl);
		foreach ($extraVars as $extra) {
			$this->headers[] = $extra->name;
			$this->fields[] = $extra->eid;
		}
	}

	public function fetch() {
		$args             = new stdClass();
		$args->module_srl = $this->moduleInfo->module_srl;
		$this->model->fetchAll('document.getDocumentList', 'page', 1, $args);

		/** @var documentModel $oDocumentModel */
		$oDocumentModel = getModel('document');

		foreach ($this->model->getResultSet() as $row) {
			foreach ($oDocumentModel->getExtraVars($this->moduleInfo->module_srl, $row->document_srl) as $ext) {
				/** @var ExtraItem $ext */
				if (!in_array($ext->eid, $this->eids))
					continue;
				$row->{$ext->eid} = $ext->value;

			}
		}

		return $this;
	}

	public function download() {
		ini_set('memory_limit','-1');

		/** @var phpexcel_moduleController $oController */
		$oController = getController('phpexcel_module');
		$oController->printDownloadHttpHeader($this->moduleInfo->mid.'.xlsx');
		$excel = $oController->createPHPExcel();
		$oController->writeToSheet($excel->getActiveSheet(), $this->fields, $this->model);
		$oController->printPHPOutStream($excel);
	}
}