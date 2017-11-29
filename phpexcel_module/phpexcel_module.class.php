<?php
/**
 * User: Kyeongdae
 * Date: 2016-06-08
 * Time: 오전 3:41
 */

require_once 'vendor/phpexcel/Classes/PHPExcel.php';
class phpexcel_module extends ModuleObject {

	public function moduleInstall() { return $this->createObject(); }

	public function checkUpdate() { }

	public function moduleUpdate() { return $this->createObject(0, 'success_updated'); }

	public function recompileCache() { }

	/** @return PHPExcel */
	public function createPHPExcel() {
		return new PHPExcel();
	}

	/** @return Context */
	public function getContext() { return Context::getInstance(); }

	public function getContextVar($key) { return $this->getContext()->get($key); }

    protected function createObject($error = 0, $message = 'success') {
        if(class_exists("BaseObject")) {
            return new BaseObject($error, $message);
        } else {
            return new Object($error, $message);
        }
    }
}