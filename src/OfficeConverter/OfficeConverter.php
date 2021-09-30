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

    /**
     * OfficeConverter constructor.
     *
     * @param string      $filename
     * @param string|null $tempPath
     * @param string      $bin
     * @param bool        $prefixExecWithExportHome
     */
    public function __construct($filename, $tempPath = null, $bin = 'libreoffice', $prefixExecWithExportHome = true)
    {
        if ($this->open($filename)) {
            $this->setup($tempPath, $bin, $prefixExecWithExportHome);
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
            throw new OfficeConverterException("Output extension({$outputExtension}) not supported for input file({$this->basename})");
        }

        $outdir = $this->tempPath;
        $shell = $this->exec($this->makeCommand($outdir, $outputExtension));
        if (0 != $shell['return']) {
            throw new OfficeConverterException('Convertion Failure! Contact Server Admin.');
        }

        return $this->prepOutput($outdir, $filename, $outputExtension);
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
            throw new OfficeConverterException('File does not exist --'.$filename);
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
    protected function setup($tempPath, $bin, $prefixExecWithExportHome)
    {
        //basename
        $this->basename = pathinfo($this->file, PATHINFO_BASENAME);

        //extension
        $extension = pathinfo($this->file, PATHINFO_EXTENSION);

        //Check for valid input file extension
        if (!array_key_exists($extension, $this->getAllowedConverter())) {
            throw new OfficeConverterException('Input file extension not supported -- '.$extension);
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

        return "{$this->bin} --headless --convert-to {$outputExtension} {$oriFile} --outdir {$outputDirectory}";
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
        $tmpName = ($this->extension ? basename($this->basename, $this->extension) : $this->basename . '.').$outputExtension;
        if (rename($outdir.$DS.$tmpName, $outdir.$DS.$filename)) {
            return $outdir.$DS.$filename;
        } elseif (is_file($outdir.$DS.$tmpName)) {
            return $outdir.$DS.$tmpName;
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
            'rtf'  => ['docx', 'txt', 'pdf'],
            'txt'  => ['pdf', 'odt', 'doc', 'docx', 'html'],
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
     *
     * @return array
     */
    private function exec($cmd, $input = '')
    {
        // Cannot use $_SERVER superglobal since that's empty during UnitUnishTestCase
        // getenv('HOME') isn't set on Windows and generates a Notice.
        if ($this->prefixExecWithExportHome) {
            $home = getenv('HOME');
            if (!is_writable($home)) {
                $cmd = 'export HOME=/tmp && '.$cmd;
            }
        }
        $process = proc_open($cmd, [0 => ['pipe', 'r'], 1 => ['pipe', 'w'], 2 => ['pipe', 'w']], $pipes);

        if (false === $process) {
            throw new OfficeConverterException('Cannot obtain ressource for process to convert file');
        }

        fwrite($pipes[0], $input);
        fclose($pipes[0]);
        $stdout = stream_get_contents($pipes[1]);
        fclose($pipes[1]);
        $stderr = stream_get_contents($pipes[2]);
        fclose($pipes[2]);
        $rtn = proc_close($process);

        return [
            'stdout' => $stdout,
            'stderr' => $stderr,
            'return' => $rtn,
        ];
    }
}
