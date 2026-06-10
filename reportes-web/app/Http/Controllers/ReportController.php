<?php

namespace App\Http\Controllers;

use App\Services\LegacyReportService;
use App\Services\ReportExportService;
use Illuminate\Database\QueryException;
use Illuminate\Http\Request;
use Throwable;
use Symfony\Component\HttpFoundation\StreamedResponse;

class ReportController extends Controller
{
    public function __construct(
        private readonly LegacyReportService $reports,
        private readonly ReportExportService $exports,
    )
    {
    }

    public function index()
    {
        return view('reports.index', [
            'reports' => $this->reports->reports(),
        ]);
    }

    public function show(string $slug)
    {
        $report = $this->reports->find($slug);

        return view('reports.show', [
            'slug' => $slug,
            'report' => $report,
            'rows' => [],
            'params' => $this->reports->normalizeParams($report, []),
            'error' => null,
            'lookups' => $this->safeLookups($report),
        ]);
    }

    public function run(Request $request, string $slug)
    {
        $report = $this->reports->find($slug);
        $params = $this->reports->normalizeParams($report, $request->all());

        try {
            $rows = $this->reports->run($slug, $request->all());
            $error = null;
        } catch (QueryException $exception) {
            $rows = [];
            $error = $exception->getMessage();
        }

        return view('reports.show', [
            'slug' => $slug,
            'report' => $report,
            'rows' => $rows,
            'params' => $params,
            'error' => $error,
            'lookups' => $this->safeLookups($report),
        ]);
    }

    public function export(Request $request, string $slug): StreamedResponse
    {
        $report = $this->reports->find($slug);
        $rows = $this->reports->run($slug, $request->query());
        $filename = $slug.'-'.now()->format('Ymd-His').'.csv';

        return response()->streamDownload(function () use ($rows): void {
            $out = fopen('php://output', 'w');

            if ($rows !== []) {
                fputcsv($out, array_keys($rows[0]));

                foreach ($rows as $row) {
                    fputcsv($out, $row);
                }
            }

            fclose($out);
        }, $filename, [
            'Content-Type' => 'text/csv; charset=UTF-8',
            'X-Report-Title' => $report['title'],
        ]);
    }

    public function excel(Request $request, string $slug)
    {
        $report = $this->reports->find($slug);
        $params = $this->reports->normalizeParams($report, $request->query());
        $rows = $this->reports->run($slug, $request->query());
        $path = $this->exports->excel($slug, $report, $params, $rows);

        return response()->download($path)->deleteFileAfterSend(true);
    }

    public function pdf(Request $request, string $slug)
    {
        $report = $this->reports->find($slug);
        $params = $this->reports->normalizeParams($report, $request->query());
        $rows = $this->reports->run($slug, $request->query());
        $path = $this->exports->pdf($slug, $report, $params, $rows);

        return response()->download($path)->deleteFileAfterSend(true);
    }

    private function safeLookups(array $report): array
    {
        try {
            return $this->reports->lookupsFor($report);
        } catch (Throwable) {
            return [];
        }
    }
}
