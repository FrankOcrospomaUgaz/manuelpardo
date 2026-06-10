<?php

namespace App\Services;

use Carbon\Carbon;
use Dompdf\Dompdf;
use Dompdf\Options;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class ReportExportService
{
    public function __construct(private readonly LegacyReportService $legacy)
    {
    }

    public function excel(string $slug, array $report, array $params, array $rows): string
    {
        if (($report['special_export'] ?? '') === 'private_monthly') {
            return $this->privateMonthlyExcel($params);
        }

        return $this->genericExcel($report, $params, $rows);
    }

    public function pdf(string $slug, array $report, array $params, array $rows): string
    {
        $html = (($report['special_export'] ?? '') === 'private_monthly')
            ? $this->privateMonthlyPdfHtml($this->legacy->privateMonthlyPayload($params))
            : view('reports.exports.generic-pdf', compact('report', 'params', 'rows'))->render();

        $options = new Options();
        $options->set('isRemoteEnabled', false);
        $options->set('defaultFont', 'Arial');

        $dompdf = new Dompdf($options);
        $dompdf->loadHtml($html);
        $dompdf->setPaper('A4', (($report['special_export'] ?? '') === 'private_monthly') ? 'landscape' : 'portrait');
        $dompdf->render();

        $path = storage_path('app/exports/'.$this->filename($slug, 'pdf'));
        $this->ensureExportDirectory();
        file_put_contents($path, $dompdf->output());

        return $path;
    }

    private function genericExcel(array $report, array $params, array $rows): string
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle('Reporte');

        $columns = $rows === [] ? ['Sin datos'] : array_keys($rows[0]);
        $lastColumn = Coordinate::stringFromColumnIndex(count($columns));

        $sheet->mergeCells("A1:{$lastColumn}1");
        $sheet->setCellValue('A1', mb_strtoupper($report['title']));
        $sheet->setCellValue('A2', 'Generado: '.now()->format('d/m/Y H:i'));
        $sheet->setCellValue('A3', $this->paramsText($params));

        $sheet->fromArray($columns, null, 'A5');
        $sheet->fromArray(array_map(fn ($row) => array_values($row), $rows), null, 'A6');

        $sheet->getStyle("A1:{$lastColumn}1")->applyFromArray([
            'font' => ['bold' => true, 'size' => 16],
            'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER],
        ]);
        $sheet->getStyle("A5:{$lastColumn}5")->applyFromArray($this->headerStyle('146C94', 'FFFFFF'));
        $sheet->getStyle("A5:{$lastColumn}".max(6, count($rows) + 5))->applyFromArray($this->thinBorders());
        $sheet->freezePane('A6');
        $sheet->setAutoFilter("A5:{$lastColumn}5");

        foreach (range(1, count($columns)) as $column) {
            $sheet->getColumnDimension(Coordinate::stringFromColumnIndex($column))->setAutoSize(true);
        }

        return $this->saveSpreadsheet($spreadsheet, $this->filename('reporte', 'xlsx'));
    }

    private function privateMonthlyExcel(array $params): string
    {
        $payload = $this->legacy->privateMonthlyPayload($params);
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle('PRIVADO');

        $this->buildPrivateSheet($sheet, $payload);

        $raw = $spreadsheet->createSheet();
        $raw->setTitle('BDPLANILLA');
        $raw->fromArray($payload['rows'] === [] ? [['Sin datos']] : array_merge([array_keys($payload['rows'][0])], array_map(fn ($row) => array_values($row), $payload['rows'])));
        foreach (range(1, max(1, count($payload['rows'][0] ?? ['Sin datos']))) as $column) {
            $raw->getColumnDimension(Coordinate::stringFromColumnIndex($column))->setAutoSize(true);
        }

        $spreadsheet->setActiveSheetIndex(0);

        return $this->saveSpreadsheet($spreadsheet, $this->filename('datos-personal-privado', 'xlsx'));
    }

    private function buildPrivateSheet($sheet, array $payload): void
    {
        $days = $payload['days'];
        $rows = $payload['rows'];
        $dayStart = 4;
        $dayEnd = $dayStart + count($days) - 1;
        $totalCol = $dayEnd + 1;
        $obsCol = $dayEnd + 2;
        $lastCol = $obsCol;
        $lastLetter = Coordinate::stringFromColumnIndex($lastCol);
        $monthStart = Coordinate::stringFromColumnIndex($dayStart);
        $monthEnd = Coordinate::stringFromColumnIndex($dayEnd);
        $totalLetter = Coordinate::stringFromColumnIndex($totalCol);
        $obsLetter = Coordinate::stringFromColumnIndex($obsCol);

        $sheet->mergeCells("A1:{$lastLetter}1");
        $sheet->setCellValue('A1', 'DATOS PERSONAL PRIVADO');
        $sheet->getStyle('A1')->getFont()->setBold(true)->setSize(18);

        $sheet->mergeCells("A2:A4");
        $sheet->mergeCells("B2:B4");
        $sheet->mergeCells("C2:C4");
        $sheet->mergeCells("{$monthStart}2:{$monthEnd}2");
        $sheet->mergeCells("{$totalLetter}2:{$totalLetter}4");
        $sheet->mergeCells("{$obsLetter}2:{$obsLetter}4");

        $sheet->setCellValue('A2', 'N');
        $sheet->setCellValue('B2', 'APELLIDOS Y NOMBRES');
        $sheet->setCellValue('C2', "INGRESO BASE\nCONTRATO");
        $sheet->setCellValue("{$monthStart}2", mb_strtoupper(Carbon::parse($payload['start'])->translatedFormat('F')));
        $sheet->setCellValue("{$totalLetter}2", "TOTAL\nMINUTOS\nACUMULADO");
        $sheet->setCellValue("{$obsLetter}2", 'OBSERVACIONES');

        foreach ($days as $offset => $day) {
            $col = Coordinate::stringFromColumnIndex($dayStart + $offset);
            $sheet->setCellValue("{$col}3", $day['weekday']);
            $sheet->setCellValue("{$col}4", $day['day']);
            $sheet->getColumnDimension($col)->setWidth(4);

            if ($day['weekend']) {
                $sheet->getStyle("{$col}3:{$col}".(count($rows) + 5))->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setRGB('D7DDE5');
            }
        }

        $sheet->getColumnDimension('A')->setWidth(5);
        $sheet->getColumnDimension('B')->setWidth(42);
        $sheet->getColumnDimension('C')->setWidth(16);
        $sheet->getColumnDimension($totalLetter)->setWidth(12);
        $sheet->getColumnDimension($obsLetter)->setWidth(48);

        $sheet->getStyle('A2:B4')->applyFromArray($this->headerStyle('70AD47', '000000'));
        $sheet->getStyle('C2:C4')->applyFromArray($this->headerStyle('D9EAF7', '000000'));
        $sheet->getStyle("{$monthStart}2:{$monthEnd}4")->applyFromArray($this->headerStyle('FFE699', '000000'));
        $sheet->getStyle("{$totalLetter}2:{$totalLetter}4")->applyFromArray($this->headerStyle('70AD47', '000000'));
        $sheet->getStyle("{$obsLetter}2:{$obsLetter}4")->applyFromArray($this->headerStyle('FFE699', '000000'));

        $rowNumber = 5;
        foreach ($rows as $row) {
            $sheet->setCellValue("A{$rowNumber}", $row['N']);
            $sheet->setCellValue("B{$rowNumber}", $row['Apellidos y nombres']);
            $sheet->setCellValue("C{$rowNumber}", $row['Ingreso base contrato']);
            $sheet->getStyle("C{$rowNumber}")->getFont()->setBold(true)->setItalic(true);

            foreach ($days as $offset => $day) {
                $col = Coordinate::stringFromColumnIndex($dayStart + $offset);
                $value = $row[$day['weekday'].' '.$day['day']] ?? '';
                $sheet->setCellValue("{$col}{$rowNumber}", $value);

                if (is_numeric($value) && (int) $value > 0) {
                    $sheet->getStyle("{$col}{$rowNumber}")->getFont()->getColor()->setRGB('FF0000');
                    $sheet->getStyle("{$col}{$rowNumber}")->getFont()->setBold(true);
                }
            }

            $sheet->setCellValue("{$totalLetter}{$rowNumber}", $row['Total minutos acumulado']);
            $sheet->setCellValue("{$obsLetter}{$rowNumber}", $row['Observaciones']);
            $sheet->getStyle("{$totalLetter}{$rowNumber}")->getFont()->setBold(true)->setItalic(true);
            $rowNumber++;
        }

        $tableEnd = max(5, $rowNumber - 1);
        $sheet->getStyle("A2:{$lastLetter}{$tableEnd}")->applyFromArray($this->thinBorders());
        $sheet->getStyle("A2:{$lastLetter}{$tableEnd}")->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
        $sheet->getStyle("A2:{$totalLetter}{$tableEnd}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle("B5:B{$tableEnd}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
        $sheet->getStyle("{$obsLetter}5:{$obsLetter}{$tableEnd}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
        $sheet->getStyle("A2:{$lastLetter}4")->getAlignment()->setWrapText(true);
        $sheet->freezePane('D5');
        $sheet->setAutoFilter("A4:{$lastLetter}4");

        $legendStart = $tableEnd + 3;
        $sheet->mergeCells("A{$legendStart}:C{$legendStart}");
        $sheet->setCellValue("A{$legendStart}", 'LEYENDA');
        $sheet->getStyle("A{$legendStart}:C{$legendStart}")->applyFromArray($this->headerStyle('BFBFBF', '000000'));

        $legend = [
            ['ASISTENCIA NORMAL', '•'],
            ['DESCANSO MEDICO', 'DM'],
            ['PERMISOS', 'P'],
            ['TARDANZAS EN MINUTOS', '1, 2, 3 ...'],
            ['LICENCIA PATERNIDAD LP', 'LP'],
            ['VACACIONES', 'V'],
            ['SUSPENSION SIN GOCE DE HABER', 'SSGH'],
            ['FALTAS', 'F'],
        ];

        foreach ($legend as $index => [$label, $mark]) {
            $r = $legendStart + 1 + $index;
            $sheet->mergeCells("A{$r}:B{$r}");
            $sheet->setCellValue("A{$r}", $label);
            $sheet->setCellValue("C{$r}", $mark);
        }

        $sheet->getStyle("A{$legendStart}:C".($legendStart + count($legend)))->applyFromArray($this->thinBorders());
        $sheet->getStyle("A{$legendStart}:C".($legendStart + count($legend)))->getFont()->setBold(true);
        $sheet->getStyle("C".($legendStart + 1).":C".($legendStart + count($legend)))->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $sheet->getPageSetup()->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE);
        $sheet->getPageSetup()->setFitToWidth(1)->setFitToHeight(0);
    }

    private function privateMonthlyPdfHtml(array $payload): string
    {
        return view('reports.exports.private-monthly-pdf', ['payload' => $payload])->render();
    }

    private function saveSpreadsheet(Spreadsheet $spreadsheet, string $filename): string
    {
        $this->ensureExportDirectory();
        $path = storage_path('app/exports/'.$filename);
        (new Xlsx($spreadsheet))->save($path);

        return $path;
    }

    private function ensureExportDirectory(): void
    {
        if (! is_dir(storage_path('app/exports'))) {
            mkdir(storage_path('app/exports'), 0775, true);
        }
    }

    private function filename(string $base, string $extension): string
    {
        return $base.'-'.now()->format('Ymd-His').'.'.$extension;
    }

    private function paramsText(array $params): string
    {
        return collect($params)->map(fn ($value, $key) => "{$key}: {$value}")->implode(' | ');
    }

    private function headerStyle(string $background, string $font): array
    {
        return [
            'font' => ['bold' => true, 'color' => ['rgb' => $font]],
            'fill' => ['fillType' => Fill::FILL_SOLID, 'startColor' => ['rgb' => $background]],
            'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER],
        ];
    }

    private function thinBorders(): array
    {
        return [
            'borders' => [
                'allBorders' => ['borderStyle' => Border::BORDER_THIN, 'color' => ['rgb' => '000000']],
            ],
        ];
    }
}
