<?php
/**
 * User: Kyeongdae
 * Date: 2016-06-23
 * Time: 오후 9:44
 */
class phpexcel_moduleAdminView extends phpexcel_module {
	/** @var phpexcel_moduleModel */
	private $model;

	function init() {
		$this->model = getModel('phpexcel_module');
		$this->setTemplatePath($this->module_path . 'tpl');
	}

	function dispPhpexcel_moduleAdminIndex() {
		$this->setTemplate('게시판 데이터 추출', 'board.export');

		$output = executeQueryArray('board.getAllBoard');
		if ($output->error) {
			return new Object($output->error, $output->message);
		}
		ModuleModel::syncModuleToSite($output->data);

		// get the module category list
		/** @var moduleModel $oModuleModel */
		$oModuleModel    = getModel('module');
		$module_category = $oModuleModel->getModuleCategories();
		$this->getContext()->set('module_category', $module_category);

		$this->getContext()->set('boards', $output->data);
	}

	function dispPhpexcel_moduleAdminDownloadBoard_extVarsFilter() {
		/** @var documentModel $oDocumentModel */
		$oDocumentModel = getModel('document');
		$extraVars      = $oDocumentModel->getExtraKeys($this->getContextVar('module_srl'));
		$this->getContext()->set('extraVars', $extraVars);
		$this->setTemplate('게시판 데이터 추출 - 확장변수 선택', 'board.extraVars');
	}

	function dispPhpexcel_moduleAdminDownloadBoard_xlsx() {
		require_once 'classes/BoardXlsDownloader.php';
		$downloader = new BoardXlsDownloader($this->model, $this->getContextVar('module_srl'), $this->getContextVar('eid'));
		$downloader
			->fetch()
			->download();
	}

	private function setTemplate($title, $tplFile) {
		$this->getContext()->set('title', $title);
		$this->setTemplateFile($tplFile);
	}
}