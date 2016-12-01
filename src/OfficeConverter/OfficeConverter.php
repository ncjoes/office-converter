<?php

namespace NcJoes\OfficeConverter;

/**
 * Class OfficeConverter
 *
 * @package NcJoes\OfficeConverter
 */
class OfficeConverter
{
    private $file;
    private $bin;
    private $tempPath;
    private $extension;
    private $basename;

    /**
     * OfficeConverter constructor.
     *
     * @param $filename
     * @param null $tempPath
     * @param string $bin
     */
    public function __construct($filename, $tempPath = null, $bin = 'libreoffice')
    {
        if ($this->open($filename)) {
            $this->setup($tempPath, $bin);
        }
    }

    /**
     * @param $filename
     *
     * @return null|string
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
        if ($shell['return'] != 0) {
            throw new OfficeConverterException("Convertion Failure! Contact Server Admin.");
        }

        return $this->prepOutput($outdir, $filename, $outputExtension);
    }

    /**
     * @param $filename
     *
     * @return bool
     * @throws OfficeConverterException
     */
    protected function open($filename)
    {
        if (!file_exists($filename)) {
            throw new OfficeConverterException('File does not exist --'.$filename);
        }
        $this->file = realpath($filename);

        return true;
    }

    /**
     * @param $tempPath
     * @param $bin
     *
     * @throws OfficeConverterException
     */
    protected function setup($tempPath, $bin)
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
        if (!is_dir($tempPath)) {
            $tempPath = dirname($this->file);
        }
        $this->tempPath = realpath($tempPath);

        //binary location
        $this->bin = $bin;
    }

    /**
     * @param $outputDirectory
     * @param $outputExtension
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
     * @param $outdir
     * @param $filename
     * @param $outputExtension
     *
     * @return null|string
     */
    protected function prepOutput($outdir, $filename, $outputExtension)
    {
        $DS = DIRECTORY_SEPARATOR;
        $tmpName = str_replace($this->extension, '', $this->basename).$outputExtension;
        if (rename($outdir.$DS.$tmpName, $outdir.$DS.$filename)) {
            return $outdir.$DS.$filename;
        }
        elseif (is_file($outdir.$DS.$tmpName)) {
            return $outdir.$DS.$tmpName;
        }

        return null;
    }

    /**
     * @param null $extension
     *
     * @return array|mixed
     */
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

    /**
     * More intelligent interface to system calls
     *
     * @link http://php.net/manual/en/function.system.php
     *
     * @param $cmd
     * @param string $input
     *
     * @return array
     */
    private function exec($cmd, $input = '')
    {
        $process = proc_open($cmd, [0 => ['pipe', 'r'], 1 => ['pipe', 'w'], 2 => ['pipe', 'w']], $pipes);
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
            'return' => $rtn
        ];
    }
}
