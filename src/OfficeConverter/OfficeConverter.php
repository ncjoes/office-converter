<?php

namespace NcJoes\OfficeConverter;

class OfficeConverter
{
    /** @var string */
    private $file;
    /** @var string */
    private $bin;
    /** @var string */
    private $tempPath;
    /** @var string */
    private $extension;
    /** @var string */
    private $basename;
    /** @var bool */
    private $prefixExecWithExportHome;
    /** @var string */
    private $filter = '';
    private $logPath;

    /**
     * OfficeConverter constructor.
     *
     * @param string      $filename
     * @param string|null $tempPath
     * @param string      $bin
     * @param bool        $prefixExecWithExportHome
     */
    public function __construct($filename, $tempPath = null, $bin = 'libreoffice', $prefixExecWithExportHome = true, $logPath = null)
    {
        if ($this->open($filename)) {
            $this->setup($tempPath, $bin, $prefixExecWithExportHome, $logPath);
        }
    }

    /**
     * @param string $filename
     *
     * @return string|null
     *
     * @throws OfficeConverterException
     */
    public function convertTo($filename)
    {
        $outputExtension = pathinfo($filename, PATHINFO_EXTENSION);
        $supportedExtensions = $this->getAllowedConverter($this->extension);

        if (!in_array($outputExtension, $supportedExtensions)) {
            throw new OfficeConverterException("Output extension ($outputExtension) not supported for input file($this->basename)");
        }

        $outdir = $this->tempPath;
        $this->exec($this->makeCommand($outdir, $outputExtension));

        return $this->prepOutput($outdir, $filename, $outputExtension);
    }

    protected static function trimString($value, $limit = 200, $end = '...')
    {
        return (mb_strwidth($value, 'UTF-8') <= $limit)
            ? $value
            : rtrim(mb_strimwidth($value, 0, $limit, '', 'UTF-8')) . $end;
    }

    /**
     * @param string $filename
     *
     * @return bool
     *
     * @throws OfficeConverterException
     */
    protected function open($filename)
    {
        if (!file_exists($filename) || false === realpath($filename)) {
            throw new OfficeConverterException('File does not exist --' . $filename);
        }

        $this->file = realpath($filename);

        return true;
    }

    /**
     * @param string|null $tempPath
     * @param string      $bin
     * @param bool        $prefixExecWithExportHome
     *
     * @return void
     *
     * @throws OfficeConverterException
     */
    protected function setup($tempPath, $bin, $prefixExecWithExportHome, $logPath)
    {
        //basename
        $this->basename = pathinfo($this->file, PATHINFO_BASENAME);

        //extension
        $extension = pathinfo($this->file, PATHINFO_EXTENSION);

        //Check for valid input file extension
        if (!array_key_exists($extension, $this->getAllowedConverter())) {
            throw new OfficeConverterException('Input file extension not supported -- ' . $extension);
        }
        $this->extension = $extension;

        //setup output path
        if (null === $tempPath || !is_dir($tempPath)) {
            $tempPath = dirname($this->file);
        }

        if (false === realpath($tempPath)) {
            $this->tempPath = sys_get_temp_dir();
        } else {
            $this->tempPath = realpath($tempPath);
        }

        //binary location
        $this->bin = $bin;

        //use prefix export home or not
        $this->prefixExecWithExportHome = $prefixExecWithExportHome;

        // log path
        $this->logPath = realpath($logPath);
    }

    /**
     * @param string $outputDirectory
     * @param string $outputExtension
     *
     * @return string
     */
    protected function makeCommand($outputDirectory, $outputExtension)
    {
        $oriFile = escapeshellarg($this->file);
        $outputDirectory = escapeshellarg($outputDirectory);
        $logCmd = $this->logPath ? ">> {$this->logPath}" : '';

        $randomNumber = mt_rand(1, 20);

        // Add the userInstallationDirectory option
        $userInstallationDirectoryOption = "-env:UserInstallation=file://{$_SERVER['HOME']}/.config/libreoffice-profile{$randomNumber}";

        return "\"$this->bin\" --headless --convert-to {$outputExtension}{$this->filter} $userInstallationDirectoryOption $oriFile --outdir $outputDirectory";
    }

    /**
     * @param string $filter
     *
     * @return OfficeConverter
     */
    public function setFilter($filter)
    {
        $this->filter = ':' . $filter;

        return $this;
    }

    /**
     * @param string $outdir
     * @param string $filename
     * @param string $outputExtension
     *
     * @return string|null
     */
    protected function prepOutput($outdir, $filename, $outputExtension)
    {
        $DS = DIRECTORY_SEPARATOR;
        $tmpName = ($this->extension ? basename($this->basename, $this->extension) : $this->basename . '.') . $outputExtension;
        if (rename($outdir . $DS . $tmpName, $outdir . $DS . $filename)) {
            return $outdir . $DS . $filename;
        }

        if (is_file($outdir . $DS . $tmpName)) {
            return $outdir . $DS . $tmpName;
        }

        return null;
    }

    /**
     * @param string|null $extension
     *
     * @return array|mixed
     */
    private function getAllowedConverter($extension = null)
    {
        $allowedConverter = [
            '' => ['pdf'],
            'html' => ['pdf', 'docx'],
            'pptx' => ['pdf'],
            'ppt' => ['pdf'],
            'pdf' => ['pdf'],
            'docx' => ['pdf', 'odt', 'html'],
            'doc' => ['pdf', 'odt', 'html'],
            'wps' => ['pdf', 'odt', 'html'],
            'dotx' => ['pdf', 'odt', 'html'],
            'docm' => ['pdf', 'odt', 'html'],
            'dotm' => ['pdf', 'odt', 'html'],
            'dot' => ['pdf', 'odt', 'html'],
            'odt' => ['pdf', 'html'],
            'xlsx' => ['pdf'],
            'xls' => ['pdf'],
            'png' => ['pdf'],
            'jpg' => ['pdf'],
            'jpeg' => ['pdf'],
            'jfif' => ['pdf'],
            'PPTX' => ['pdf'],
            'PPT' => ['pdf'],
            'PDF' => ['pdf'],
            'DOCX' => ['pdf', 'odt', 'html'],
            'DOC' => ['pdf', 'odt', 'html'],
            'WPS' => ['pdf', 'odt', 'html'],
            'DOTX' => ['pdf', 'odt', 'html'],
            'DOCM' => ['pdf', 'odt', 'html'],
            'DOTM' => ['pdf', 'odt', 'html'],
            'DOT' => ['pdf', 'odt', 'html'],
            'ODT' => ['pdf', 'html'],
            'XLSX' => ['pdf'],
            'XLS' => ['pdf'],
            'PNG' => ['pdf'],
            'JPG' => ['pdf'],
            'JPEG' => ['pdf'],
            'JFIF' => ['pdf'],
            'Pptx' => ['pdf'],
            'Ppt' => ['pdf'],
            'Pdf' => ['pdf'],
            'Docx' => ['pdf', 'odt', 'html'],
            'Doc' => ['pdf', 'odt', 'html'],
            'Wps' => ['pdf', 'odt', 'html'],
            'Dotx' => ['pdf', 'odt', 'html'],
            'Docm' => ['pdf', 'odt', 'html'],
            'Dotm' => ['pdf', 'odt', 'html'],
            'Dot' => ['pdf', 'odt', 'html'],
            'Ddt' => ['pdf', 'html'],
            'Xlsx' => ['pdf'],
            'Xls' => ['pdf'],
            'Png' => ['pdf'],
            'Jpg' => ['pdf'],
            'Jpeg' => ['pdf'],
            'Jfif' => ['pdf'],
            'rtf' => ['docx', 'txt', 'pdf'],
            'txt' => ['pdf', 'odt', 'doc', 'docx', 'html'],
            'csv' => ['pdf'],
        ];

        if (null !== $extension) {
            if (isset($allowedConverter[$extension])) {
                return $allowedConverter[$extension];
            }

            return [];
        }

        return $allowedConverter;
    }

    /**
     * More intelligent interface to system calls.
     *
     * @see http://php.net/manual/en/function.system.php
     *
     * @param string $cmd
     * @param string $input
     */
    private function exec($cmd, $input = '')
    {
        // Cannot use $_SERVER superglobal since that's empty during UnitUnishTestCase
        // getenv('HOME') isn't set on Windows and generates a Notice.
        if ($this->prefixExecWithExportHome && false === stripos(PHP_OS, 'WIN')) {
            $home = getenv('HOME');
            if (!is_writable($home)) {
                $cmd = 'export HOME=/tmp && ' . $cmd;
            }
        }

        $exec = exec($cmd . ' 2>&1', $output, $rtn);

        if (false === $exec || 0 !== $rtn) {
            $croppedStderr = self::trimString(implode("\n", $output), 1000);

            throw new OfficeConverterException('Convertion Failure! Contact Server Admin:' . "code $rtn \nerror: $croppedStderr");
        }
    }
}
