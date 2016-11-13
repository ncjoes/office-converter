<?php

namespace NcJoes\OfficeConverter;

class OfficeConverter
{
    private $file;
    private $bin;
    private $tempPath;
    private $extension;
    private $basename;

    public function __construct($filename, $tempPath = null, $bin = 'soffice')
    {
        $this->file = $this->open($filename);
        $this->basename = pathinfo($this->file, PATHINFO_BASENAME);
        $extension = pathinfo($this->file, PATHINFO_EXTENSION);

        //Check for valid input file extension
        if (!array_key_exists($extension, $this->getAllowedConverter())) {
            throw new OfficeConverterException('Input file extension not supported -- '.$extension);
        }
        $this->extension = $extension;

        //setup output path
        if (!is_dir($tempPath)) {
            $tempPath = dirname($this->file);
        }
        $this->tempPath = $tempPath;

        //binary location
        $this->bin = $bin;
    }

    protected function open($filename)
    {
        if (!file_exists($filename)) {
            throw new OfficeConverterException('File does not exist --'.$filename);
        }

        return realpath($filename);
    }

    public function convertTo($filename)
    {
        $ext = pathinfo($filename, PATHINFO_EXTENSION);
        $allowedExt = $this->getAllowedConverter($this->extension);

        if (!in_array($ext, $allowedExt)) {
            throw new OfficeConverterException("Output extension({$ext}) not supported for input file({$this->basename})");
        }

        $oriFile = escapeshellarg($this->file);
        $outdir = escapeshellarg($this->tempPath);
        $cmd = "{$this->bin} --headless -convert-to {$ext} {$oriFile} -outdir {$outdir}";
        shell_exec($cmd);
    }

    private function getAllowedConverter($extension = null)
    {
        $allowedConverter = [
            'pptx' => ['pdf'],
            'ppt' => ['pdf'],
            'pdf' => ['pdf'],
            'docx' => ['pdf', 'odt', 'html'],
            'doc' => ['pdf', 'odt', 'html'],
            'xlsx' => ['pdf'],
            'xls' => ['pdf'],
        ];

        if ($extension !== null) {
            if (isset($allowedConverter[ $extension ])) {
                return $allowedConverter[ $extension ];
            }

            return [];
        }

        return $allowedConverter;
    }
}
