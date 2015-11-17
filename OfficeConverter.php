<?php

class OfficeConverter {
	private $file;
	private $filename;
	private $extension;
	private $path;
	private $bin;
	private $tempPath;

	public function __construct($filename, $bin = 'libreoffice', $tempPath = '/tmp/libreoffice'){
		$this->bin = $bin;
		$this->tempPath = $tempPath;

		$this->filename = $filename;
		$this->extension = pathinfo($filename, PATHINFO_EXTENSION);

		if ( !$this->isFilename() ){
			$this->path = dirname($this->filename).DIRECTORY_SEPARATOR;
			$this->filename = basename($this->filename);
			if ( substr($this->path, 0, 1) === DIRECTORY_SEPARATOR);{
				$this->path = $this->getCallerPath().DIRECTORY_SEPARATOR.$this->path;
			}
		} else {
			$this->path = $this->getCallerPath().DIRECTORY_SEPARATOR;
		}

		$this->file = $this->path.$this->filename;
		if ( !file_exists($this->file) ){
			throw new Exception('File not exist');
		}
	}

	public function convertTo($filename){
		$ext = pathinfo($filename, PATHINFO_EXTENSION);
		$allowedExt = $this->getAllowedConverter($this->extension);
		$path = $this->getCallerPath().DIRECTORY_SEPARATOR.dirname($filename).DIRECTORY_SEPARATOR;
		$filename = basename($filename);
		$oriFilename = str_replace('.'.$this->extension, '.'.$ext, $this->filename);
		$file = $path.$filename;
		if ( !in_array($ext, $allowedExt) ){
			throw new Exception('File extension not supported');
		}
		$oriFile = escapeshellarg($this->file);
		$cmd = "{$this->bin} --headless --convert-to {$ext} --outdir {$this->tempPath} {$oriFile} 2>&1";
		exec($cmd, $result, $returnVar);
		if ( $returnVar !== 0 ){
			$result = implode(' ', $result);
			throw new Exception("{$cmd} command cant run. {$result}. Error code: {$returnVar}");
			return false;
		}
		try {
			if ( rename($this->tempPath.DIRECTORY_SEPARATOR.$oriFilename, $file) ){
				return true;
			}
		} catch (Exception $e) {
		}
		return false;
	}

	public function toArray(){
		return [
			'file' => $this->file,
			'filename' => $this->filename,
			'extension' => $this->extension,
			'path' => $this->path,
		];
	}

	private function isFilename(){
		return basename($this->filename) === $this->filename;
	}

	private function getAllowedConverter($extension = null){
		$allowedConverter = [
			'pptx' => ['pdf'],
			'ppt' => ['pdf'],
			'pdf' => ['pdf'],
			'docx' => ['pdf', 'odt'],
			'doc' => ['pdf', 'odt'],
			'xlsx' => ['pdf'],
			'xls' => ['pdf'],
		];
		if ( $extension !== null ){
			if ( isset( $allowedConverter[$extension] ) ){
				return $allowedConverter[$extension];
			}
			return [];
		}
		return $allowedConverter;
	}

	private function getCallerPath(){
		$trace = debug_backtrace();
		if ( empty($trace) || !isset($trace[1]) || !isset($trace[1]['file']) ){
			return null;
		}
		return dirname($trace[1]['file']);
	}
}
