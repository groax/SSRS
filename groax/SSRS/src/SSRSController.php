<?php


namespace Groax\Ssrs;

use SSRS\Report;

class SSRSController
{
    protected $options = [];
    protected $conn;

    /**
     * SSRSController constructor.
     */
    public function __construct()
    {
        $this->options = ["username" => env('SSRS_UID'), "password" => env('SSRS_PASWD')];
        $this->conn = new Report(env('SSRS_URL'), $this->options, env('SSRS_SERVICE_URL'));
    }

    /**
     * @param string $path
     * @param array $parameters
     * @param string|null $historyId
     * @return \App\Model\SSRS\SSRSController
     */
    public function loadReport(string $path, array $parameters, string $historyId = null): SSRSController
    {
        $executionInfo = $this->conn->loadReport(strtoupper($path), $historyId);
        $this->conn->setSessionId($executionInfo->executionInfo->ExecutionID);

        if (!empty($parameters)) {
            $this->conn->setExecutionParameters($parameters);
        }

        return $this;
    }

    /**
     * @param string $format
     * @return \SSRS\Object\ReportOutput|\SSRS\SSRS\Object\ReportOutput
     */
    public function render(string $format = 'HTML4.0')
    {
        return $this->conn->render($format); // PDF | XML | CSV
    }

    /**
     * @param string $report_name
     * @param string $format
     */
    public function download(string $report_name, string $format)
    {
        $file = $this->conn->render($format);

        header('Content-Description: File Transfer');
        header('Content-Type: application/'.strtolower($format));
        header('Content-Disposition: attachment; filename='.$report_name. $this->getExtension($format));
        header('Content-Transfer-Encoding: binary');
        header('Expires: 0');
        header('Cache-Control: must-revalidate');
        header('Pragma: public');
        header('Content-Length: ' . strlen($file));
        ob_clean();
        flush();
        echo $file;
        exit;
    }

    /**
     * @param string $type
     * @return string|null
     */
    public function getExtension(string $type)
    {
        switch($type)
        {
            case "CSV":
                return ".csv";
            case "EXCEL":
                return ".xls";
            case "IMAGE":
                return ".jpg";
            case "PDF":
                return ".pdf";
            case "WORD":
                return ".doc";
            case "XML":
                return ".xml";
            default:
                return null;
        }
    }
}
