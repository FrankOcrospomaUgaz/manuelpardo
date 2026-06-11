<?php

namespace App\Services;

use Carbon\Carbon;
use Illuminate\Support\Facades\DB;
use InvalidArgumentException;

class LegacyReportService
{
    public function reports(): array
    {
        return config('reports');
    }

    public function find(string $slug): array
    {
        $report = $this->reports()[$slug] ?? null;

        if (! $report) {
            throw new InvalidArgumentException("Reporte no encontrado: {$slug}");
        }

        return $report;
    }

    public function run(string $slug, array $input): array
    {
        $report = $this->find($slug);
        $params = $this->normalizeParams($report, $input);

        if (isset($report['handler'])) {
            return $this->{$report['handler']}($params);
        }

        return $this->callProcedure($report['procedure'], array_values($params));
    }

    public function normalizeParams(array $report, array $input): array
    {
        $params = [];

        foreach ($report['params'] ?? [] as $name => $definition) {
            $type = $definition['type'] ?? 'text';
            $value = $input[$name] ?? $this->defaultValueFor($name, $definition);

            if ($type === 'checkbox') {
                $params[$name] = filter_var($value, FILTER_VALIDATE_BOOLEAN);
                continue;
            }

            if ($type === 'multiselect') {
                $params[$name] = $this->parseIds($value);
                continue;
            }

            if ($value === '' || $value === null) {
                $params[$name] = $this->emptyValueFor($name, $type, $definition);
                continue;
            }

            $params[$name] = in_array($type, ['number', 'lookupselect'], true) ? (int) $value : $value;
        }

        return $params;
    }

    public function callProcedure(string $procedure, array $params = []): array
    {
        $params = array_map(fn (mixed $param) => $this->legacySqlParam($param), $params);
        $placeholders = implode(', ', array_fill(0, count($params), '?'));
        $sql = trim("EXEC {$procedure} {$placeholders}");

        return $this->rowsToArrays(DB::select($sql, $params));
    }

    public function lookupsFor(array $report): array
    {
        $lookups = [];

        foreach ($report['params'] ?? [] as $definition) {
            $lookup = $definition['lookup'] ?? null;

            if ($lookup && ! isset($lookups[$lookup])) {
                $lookups[$lookup] = $this->lookup($lookup);
            }
        }

        return $lookups;
    }

    public function lookup(string $name): array
    {
        return match ($name) {
            'groups' => $this->storedLookup('up_BuscarGrupo', 'IdGrupo'),
            'positions' => $this->storedLookup('up_BuscarCargo', 'IdCargo'),
            'categories' => $this->storedLookup('up_BuscarCategoria', 'IdCategoria'),
            'areas' => $this->storedLookup('up_BuscarArea', 'IdArea'),
            'reasons' => $this->storedLookup('up_BuscarMotivo', 'IdMotivoJustificacion'),
            'hour_types' => $this->tableLookup('TipoHora', 'IdTipoHora', 'Descripcion'),
            'employees' => $this->employeeLookup(),
            default => [],
        };
    }

    public function groups(): array
    {
        return $this->lookup('groups');
    }

    private function storedLookup(string $procedure, string $idColumn): array
    {
        return collect($this->callProcedure($procedure, [0]))
            ->map(fn (array $row) => [
                'value' => (int) $this->value($row, $idColumn, strtolower($idColumn), strtoupper($idColumn)),
                'label' => (string) $this->value($row, 'Descripcion', 'Grupo', 'Nombre', 'DESCRIPCION'),
            ])
            ->filter(fn (array $row) => $row['value'] > 0 && $row['label'] !== '')
            ->values()
            ->all();
    }

    private function tableLookup(string $table, string $idColumn, string $labelColumn): array
    {
        return collect(DB::table($table)->select($idColumn, $labelColumn)->orderBy($labelColumn)->get()->all())
            ->map(fn (object $row) => ['value' => (int) $row->{$idColumn}, 'label' => (string) $row->{$labelColumn}])
            ->filter(fn (array $row) => $row['value'] > 0 && $row['label'] !== '')
            ->values()
            ->all();
    }

    private function employeeLookup(): array
    {
        return collect(DB::table('Trabajador')
            ->select('IdTrabajador', 'Apellido', 'Nombre')
            ->orderBy('Apellido')
            ->orderBy('Nombre')
            ->get()
            ->all())
            ->map(fn (object $row) => [
                'value' => (int) $row->IdTrabajador,
                'label' => trim((string) $row->Apellido.' '.(string) $row->Nombre),
            ])
            ->filter(fn (array $row) => $row['value'] > 0 && $row['label'] !== '')
            ->values()
            ->all();
    }

    public function datosPersonalPrivado(array $p): array
    {
        return $this->privateMonthlyPayload($p)['rows'];
    }

    public function privateMonthlyPayload(array $p): array
    {
        $start = Carbon::parse($p['fecha_inicio'])->startOfDay();
        $end = Carbon::parse($p['fecha_fin'])->startOfDay();
        $groups = $this->parseIds($p['ids_grupo'] ?? []);
        $sources = $groups === [] ? [0] : $groups;
        $employees = [];

        foreach ($sources as $groupId) {
            foreach ($this->callProcedure('up_ReporteHorariosTrab', [0, $groupId, $p['id_cargo'], $start->toDateString(), $end->toDateString()]) as $row) {
                $id = (int) $this->value($row, 'IdTrabajador', 'idtrabajador', 'IDTRABAJADOR');

                if ($id <= 0 || isset($employees[$id])) {
                    continue;
                }

                $employees[$id] = [
                    'id' => $id,
                    'name' => $this->employeeName($row),
                    'base_time' => $this->baseTime($row),
                    'group_id' => $groupId,
                    'days' => [],
                    'total_minutes' => 0,
                    'observation' => '',
                ];
            }
        }

        foreach ($employees as $id => $employee) {
            $attendance = $this->callProcedure('up_ReporteDetalleAsistenciaxTrabajador', [
                $start->toDateString(),
                $end->toDateString(),
                $this->toleranciaTardanza(),
                $this->toleranciaJornadaIncompleta(),
                $id,
            ]);

            foreach ($attendance as $row) {
                $date = $this->dateKey($this->value($row, 'Fecha', 'FECHA', 'fecha'));

                if (! $date) {
                    continue;
                }

                $employees[$id]['days'][$date] = $this->attendanceMark($row);
                $late = (int) $this->value($row, 'Tardanza', 'MinutosTardanza', 'TARDANZA', 'TardanzaMinutos');

                if ($late > 0) {
                    $employees[$id]['total_minutes'] += $late;
                }

                $obs = $this->observationText($row);
                if ($obs !== '' && ! str_contains($employees[$id]['observation'], $obs)) {
                    $employees[$id]['observation'] = trim($employees[$id]['observation'].' '.$obs);
                }
            }
        }

        $this->applyPrivateMonthlyExceptions($employees, $start, $end, $sources, (bool) ($p['solo_dia'] ?? true));

        $days = [];
        for ($date = $start->copy(); $date <= $end; $date->addDay()) {
            $days[] = [
                'date' => $date->toDateString(),
                'day' => $date->day,
                'weekday' => $this->spanishWeekday($date),
                'weekend' => $date->isWeekend(),
            ];
        }

        $rows = [];
        foreach (array_values($employees) as $index => $employee) {
            $row = [
                'N' => $index + 1,
                'Apellidos y nombres' => $employee['name'],
                'Ingreso base contrato' => $employee['base_time'],
            ];

            foreach ($days as $day) {
                $row[$day['weekday'].' '.$day['day']] = $employee['days'][$day['date']] ?? ($day['weekend'] ? '' : $this->normalMark());
            }

            $row['Total minutos acumulado'] = $employee['total_minutes'];
            $row['Observaciones'] = $employee['observation'];
            $rows[] = $row;
        }

        return [
            'title' => 'DATOS PERSONAL PRIVADO',
            'period' => $start->translatedFormat('F Y'),
            'start' => $start,
            'end' => $end,
            'days' => $days,
            'rows' => $rows,
            'employees' => array_values($employees),
        ];
    }

    private function applyPrivateMonthlyExceptions(array &$employees, Carbon $start, Carbon $end, array $groups, bool $soloDia): void
    {
        if ($employees === []) {
            return;
        }

        $absenceProcedure = $soloDia ? 'up_ReporteRecordInasistenciasSoloDiaD' : 'up_ReporteRecordInasistenciasD';

        foreach ($groups as $groupId) {
            foreach ($this->callProcedure('up_ReporteRecordTardanzasD', [0, $groupId, $start->toDateString(), $end->toDateString(), $this->toleranciaTardanza()]) as $row) {
                $this->markPrivateMonthlyLate($employees, $row, $start, $end);
            }

            foreach ($this->callProcedure($absenceProcedure, [0, $groupId, $start->toDateString(), $end->toDateString()]) as $row) {
                $this->markPrivateMonthlyException($employees, $row, $start, $end, 'F');
            }

            foreach ($this->callProcedure('up_ReporteJustificacionTrab', [0, $groupId, 0, $start->toDateString(), $end->toDateString()]) as $row) {
                $mark = $this->justificationMark($row);
                $this->markPrivateMonthlyException($employees, $row, $start, $end, $mark, $this->observationText($row));
            }
        }
    }

    private function markPrivateMonthlyLate(array &$employees, array $row, Carbon $start, Carbon $end): void
    {
        $workerId = (int) $this->value($row, 'IdTrabajador', 'idtrabajador', 'IDTRABAJADOR', 'Codigo', 'CODIGO');
        $minutes = $this->lateMinutes($row);

        if ($workerId <= 0 || ! isset($employees[$workerId]) || $minutes <= 0) {
            return;
        }

        foreach ($this->datesFromRow($row, $start, $end) as $date) {
            $current = $employees[$workerId]['days'][$date] ?? null;

            if ($current && ! $this->isNormalOrLateMark($current)) {
                continue;
            }

            if (! is_numeric((string) $current)) {
                $employees[$workerId]['total_minutes'] += $minutes;
            }

            $employees[$workerId]['days'][$date] = (string) $minutes;
        }
    }

    private function lateMinutes(array $row): int
    {
        foreach (['Tardanza', 'MinutosTardanza', 'TARDANZA', 'TardanzaMinutos', 'Minutos', 'MINUTOS', 'Diferencia', 'DIFERENCIA'] as $key) {
            $value = $this->value($row, $key);

            if (is_numeric($value) && (int) $value > 0) {
                return (int) $value;
            }
        }

        foreach ($row as $key => $value) {
            $key = $this->plainUpper((string) $key);

            if ((str_contains($key, 'TARD') || str_contains($key, 'MIN') || str_contains($key, 'DIF')) && is_numeric($value) && (int) $value > 0) {
                return (int) $value;
            }
        }

        return 0;
    }

    private function isNormalOrLateMark(mixed $mark): bool
    {
        if ($mark === null || $mark === '') {
            return true;
        }

        $mark = trim((string) $mark);

        return $mark === $this->normalMark() || is_numeric($mark);
    }

    private function markPrivateMonthlyException(array &$employees, array $row, Carbon $start, Carbon $end, string $mark, string $observation = ''): void
    {
        $workerId = (int) $this->value($row, 'IdTrabajador', 'idtrabajador', 'IDTRABAJADOR', 'Codigo', 'CODIGO');

        if ($workerId <= 0 || ! isset($employees[$workerId])) {
            return;
        }

        foreach ($this->datesFromRow($row, $start, $end) as $date) {
            $employees[$workerId]['days'][$date] = $mark;
        }

        if ($observation !== '' && ! str_contains($employees[$workerId]['observation'], $observation)) {
            $employees[$workerId]['observation'] = trim($employees[$workerId]['observation'].' '.$observation);
        }
    }

    private function datesFromRow(array $row, Carbon $start, Carbon $end): array
    {
        $single = $this->value($row, 'Fecha', 'FECHA', 'fecha', 'FechaInasistencia', 'FechaTardanza', 'FechaMarcacion', 'FechaRegistro');
        $from = $this->value($row, 'FechaInicio', 'FECHAINICIO', 'fecha_inicio', 'Desde', 'FechaDesde') ?? $single;
        $to = $this->value($row, 'FechaFin', 'FECHAFIN', 'fecha_fin', 'Hasta', 'FechaHasta') ?? $single ?? $from;

        if (! $from) {
            $dates = [];

            foreach ($row as $value) {
                $date = $this->parseLegacyDate($value);

                if ($date && $date->betweenIncluded($start, $end)) {
                    $dates[] = $date->toDateString();
                }
            }

            return array_values(array_unique($dates));
        }

        $from = $this->parseLegacyDate($from);
        $to = $this->parseLegacyDate($to ?: $from);

        if (! $from || ! $to) {
            return [];
        }

        $from = $from->startOfDay();
        $to = $to->startOfDay();

        if ($from->lessThan($start)) {
            $from = $start->copy();
        }

        if ($to->greaterThan($end)) {
            $to = $end->copy();
        }

        $dates = [];

        for ($date = $from->copy(); $date <= $to; $date->addDay()) {
            $dates[] = $date->toDateString();
        }

        return $dates;
    }

    private function justificationMark(array $row): string
    {
        $text = $this->plainUpper(trim($this->observationText($row).' '.(string) $this->value($row, 'MotivoJustificacion', 'Motivo', 'Descripcion')));

        return match (true) {
            str_contains($text, 'VACACION') || str_contains($text, 'VACACIONES') => 'V',
            str_contains($text, 'DESCANSO MEDICO') || preg_match('/\bDM\b/', $text) => 'DM',
            str_contains($text, 'PATERNIDAD') => 'LP',
            str_contains($text, 'SUSPENSION') => 'SSGH',
            str_contains($text, 'FALTA') || str_contains($text, 'INASIST') => 'F',
            default => 'P',
        };
    }

    public function registroMarcaciones(array $p): array
    {
        $fechaFin = $p['fecha_fin'] ?: $p['fecha_inicio'];
        $hasTipoHora = (int) $p['id_tipo_hora'] > 0;
        $singleDay = $fechaFin === $p['fecha_inicio'];

        if ($singleDay && $hasTipoHora) {
            return $this->callProcedure('up_ReporteRegistroMarcacionDiaTipo', [
                $p['hora_inicio'], $p['hora_fin'], $p['fecha_inicio'], $p['id_trabajador'], $p['id_tipo_hora'],
            ]);
        }

        if ($singleDay) {
            return $this->callProcedure('up_ReporteRegistroMarcacionDia', [
                $p['hora_inicio'], $p['hora_fin'], $p['fecha_inicio'], $p['id_trabajador'],
            ]);
        }

        if ($hasTipoHora) {
            return $this->callProcedure('up_ReporteRegistroMarcacionFechaTipo', [
                $p['hora_inicio'], $p['hora_fin'], $p['fecha_inicio'], $fechaFin, $p['id_trabajador'], $p['id_tipo_hora'],
            ]);
        }

        return $this->callProcedure('up_ReporteRegistroMarcacionFecha', [
            $p['hora_inicio'], $p['hora_fin'], $p['fecha_inicio'], $fechaFin, $p['id_trabajador'],
        ]);
    }

    public function comparacionHoras(array $p): array
    {
        return match ($p['tipo']) {
            'asistencia' => $this->callProcedure('up_ReporteComparacionHoraTrabAsistencia', [
                $p['id_trabajador'], $p['id_grupo'], $p['fecha_inicio'], $p['fecha_fin'],
            ]),
            'listado' => $this->callProcedure('up_ReporteComparacionHoraTrabL', [
                $p['id_trabajador'], $p['fecha_inicio'], $p['fecha_fin'],
            ]),
            default => $this->callProcedure('up_ReporteComparacionHoraTrab', [
                $p['id_trabajador'], $p['fecha_inicio'], $p['fecha_fin'],
            ]),
        };
    }

    public function justificaciones(array $p): array
    {
        $procedure = $p['formato'] === 'listado'
            ? 'up_ReporteJustificacionTrabL'
            : 'up_ReporteJustificacionTrab';

        return $this->callProcedure($procedure, [
            $p['id_trabajador'], $p['id_grupo'], $p['id_motivo'], $p['fecha_inicio'], $p['fecha_fin'],
        ]);
    }

    public function noMarcaciones(array $p): array
    {
        $suffix = $p['formato'] === 'listado' ? 'L' : 'D';
        $base = match ($p['tipo']) {
            'entrada_salida' => 'up_ReporteRegistroNoMarcacionesES',
            'ambas' => 'up_ReporteRegistroNoMarcacionesAmbas',
            default => 'up_ReporteRegistroNoMarcaciones',
        };

        return $this->callProcedure($base.$suffix, [
            $p['id_trabajador'], $p['id_grupo'], $p['fecha_inicio'], $p['fecha_fin'],
        ]);
    }

    public function tardanzas(array $p): array
    {
        $procedure = $p['formato'] === 'listado'
            ? 'up_ReporteRecordTardanzasL'
            : 'up_ReporteRecordTardanzasD';

        return $this->callProcedure($procedure, [
            $p['id_trabajador'], $p['id_grupo'], $p['fecha_inicio'], $p['fecha_fin'], $this->toleranciaTardanza(),
        ]);
    }

    public function inasistencias(array $p): array
    {
        $procedure = match (true) {
            $p['solo_dia'] && $p['formato'] === 'listado' => 'up_ReporteRecordInasistenciasSoloDiaL',
            $p['solo_dia'] => 'up_ReporteRecordInasistenciasSoloDiaD',
            $p['formato'] === 'listado' => 'up_ReporteRecordInasistenciasL',
            default => 'up_ReporteRecordInasistenciasD',
        };

        return $this->callProcedure($procedure, [
            $p['id_trabajador'], $p['id_grupo'], $p['fecha_inicio'], $p['fecha_fin'],
        ]);
    }

    public function sobretiempos(array $p): array
    {
        $procedure = $p['formato'] === 'listado'
            ? 'up_ReporteRecordSobretiemposL'
            : 'up_ReporteRecordSobretiemposD';

        return $this->callProcedure($procedure, [
            $p['id_trabajador'], $p['id_grupo'], $p['fecha_inicio'], $p['fecha_fin'],
        ]);
    }

    public function detalleAsistencia(array $p): array
    {
        $procedure = match ($p['formato']) {
            'justificado' => 'up_ReporteDetalleAsistenciaxTrabajadorJust',
            'listado' => 'up_ReporteDetalleAsistenciaxTrabajadorL',
            default => 'up_ReporteDetalleAsistenciaxTrabajador',
        };

        return $this->callProcedure($procedure, [
            $p['fecha_inicio'], $p['fecha_fin'], $this->toleranciaTardanza(), $this->toleranciaJornadaIncompleta(), $p['id_trabajador'],
        ]);
    }

    public function bitacora(array $p): array
    {
        $query = DB::table('Bitacora');

        foreach (['usuario' => 'Usuario', 'tabla' => 'Tabla', 'operacion' => 'Operacion'] as $param => $column) {
            if ($p[$param] !== '') {
                $query->where($column, $p[$param]);
            }
        }

        if ($p['fecha_inicio'] !== '') {
            $query->where('Fecha', '>=', $p['fecha_inicio']);
        }

        if ($p['fecha_fin'] !== '') {
            $query->where('Fecha', '<=', $p['fecha_fin']);
        }

        return $this->rowsToArrays($query->orderByDesc('Fecha')->get()->all());
    }

    public function consolidadoTrabajador(array $p): array
    {
        $marcaciones = $this->callProcedure('up_BuscarMarcacionesTrabajador', [
            $p['id_trabajador'], $p['fecha_inicio'], $p['fecha_fin'],
        ]);

        $trabajador = $this->trabajador($p['id_trabajador']);
        $suspension = $this->firstRow($this->callProcedure('up_BuscarMinutoSuspensionTrabajador', [
            $p['id_trabajador'], $p['fecha_inicio'], $p['fecha_fin'],
        ]));

        $result = [
            'IdTrabajador' => $p['id_trabajador'],
            'Apellidos' => trim(($trabajador['Apellido'] ?? '').', '.ucwords(strtolower($trabajador['Nombre'] ?? ''))),
            'FID' => 0, 'FNRD' => 0, 'SPD' => (int) ($suspension['Cantidad'] ?? 0),
            'FIH' => 0, 'FNRH' => 0, 'SPH' => (float) ($suspension['Diferencia'] ?? 0),
            'FRD' => 0, 'FRH' => 0, 'TARDRD' => 0, 'TARDNRD' => 0,
            'TARDRH' => 0, 'TARDNRH' => 0, 'HORASTRAB' => 0, 'FJNRD' => 0,
            'FJNRH' => 0, 'JIRD' => 0, 'JIRH' => 0, 'JIINJD' => 0, 'JIINJH' => 0,
            'JIJUSTNRD' => 0, 'JIJUSTNRH' => 0, 'FJRD' => 0, 'FJRH' => 0,
        ];

        $holguraEntrada = $this->holgura('Entrada');
        $holguraSalida = $this->holgura('Salida');
        $ultimoFeriadoNoRemunerado = null;
        $ultimoFeriadoRemunerado = null;

        foreach ($marcaciones as $marcacion) {
            $fecha = $this->value($marcacion, 'Fecha', 'FECHA');
            $feriado = $this->feriado($fecha);
            $justificacion = $this->value($marcacion, 'IdJustificacion');
            $remunerada = $justificacion ? (bool) ($this->motivoJustificacion($justificacion)['Remunerada'] ?? false) : null;
            $ideal = $this->minutesBetween($this->value($marcacion, 'HoraIngresoIdeal'), $this->value($marcacion, 'HoraSalidaIdeal'));
            $entradaReal = $this->value($marcacion, 'HoraIngresoReal');
            $salidaReal = $this->value($marcacion, 'HoraSalidaReal');

            if ($feriado) {
                if ((bool) ($feriado['Remunerado'] ?? false)) {
                    $result['FRH'] += $ideal;
                    if ($ultimoFeriadoRemunerado !== $fecha) {
                        $result['FRD']++;
                        $ultimoFeriadoRemunerado = $fecha;
                    }
                } else {
                    $result['FNRH'] += $ideal;
                    if ($ultimoFeriadoNoRemunerado !== $fecha) {
                        $result['FNRD']++;
                        $ultimoFeriadoNoRemunerado = $fecha;
                    }
                }

                continue;
            }

            if (! $entradaReal || ! $salidaReal) {
                $keyD = $justificacion ? ($remunerada ? 'FJRD' : 'FJNRD') : 'FID';
                $keyH = $justificacion ? ($remunerada ? 'FJRH' : 'FJNRH') : 'FIH';
                $result[$keyD]++;
                $result[$keyH] += $ideal;
                continue;
            }

            $tarde = max(0, $this->minutesBetween($this->value($marcacion, 'HoraIngresoIdeal'), $entradaReal));
            $incompleta = max(0, $this->minutesBetween($salidaReal, $this->value($marcacion, 'HoraSalidaIdeal')));

            if ($tarde > $holguraEntrada) {
                $keyD = $justificacion ? ($remunerada ? 'TARDRD' : 'TARDNRD') : 'TARDNRD';
                $keyH = $justificacion ? ($remunerada ? 'TARDRH' : 'TARDNRH') : 'TARDNRH';
                $result[$keyD]++;
                $result[$keyH] += $tarde;
            }

            if ($incompleta > $holguraSalida) {
                $keyD = $justificacion ? ($remunerada ? 'JIRD' : 'JIJUSTNRD') : 'JIINJD';
                $keyH = $justificacion ? ($remunerada ? 'JIRH' : 'JIJUSTNRH') : 'JIINJH';
                $result[$keyD]++;
                $result[$keyH] += $incompleta;
            }

            $result['HORASTRAB'] += $this->minutesBetween($entradaReal, $salidaReal);
        }

        return [$result];
    }

    private function toleranciaTardanza(): int
    {
        $row = $this->firstRow($this->rowsToArrays(DB::select('SELECT TOP 1 ToleranciaTardanza FROM Cliente')));

        return (int) ($row['ToleranciaTardanza'] ?? 0);
    }

    private function toleranciaJornadaIncompleta(): int
    {
        $row = $this->firstRow($this->rowsToArrays(DB::select('SELECT TOP 1 ToleranciaJornadaIncompleta FROM Cliente')));

        return (int) ($row['ToleranciaJornadaIncompleta'] ?? 0);
    }

    private function holgura(string $descripcion): int
    {
        $row = $this->firstRow($this->callProcedure('up_BuscarHolgura', [$descripcion]));

        return (int) ($row['ValorMin'] ?? 0);
    }

    private function feriado(string $fecha): ?array
    {
        $row = DB::table('Feriado')
            ->select('IdFeriado', 'FechaFeriado', 'Remunerado')
            ->whereDate('FechaFeriado', $fecha)
            ->first();

        return $row ? (array) $row : null;
    }

    private function motivoJustificacion(int $id): array
    {
        return $this->firstRow($this->callProcedure('up_BuscarMotivoJustificacion', [$id]));
    }

    private function trabajador(int $id): array
    {
        $row = DB::table('Trabajador')->where('IdTrabajador', $id)->first();

        return $row ? (array) $row : [];
    }

    private function defaultValueFor(string $name, array $definition): mixed
    {
        if (array_key_exists('default', $definition)) {
            return $definition['default'];
        }

        if ($name === 'fecha_inicio') {
            return now()->subMonthNoOverflow()->startOfMonth()->toDateString();
        }

        if ($name === 'fecha_fin') {
            return now()->subMonthNoOverflow()->endOfMonth()->toDateString();
        }

        return null;
    }

    private function emptyValueFor(string $name, string $type, array $definition): mixed
    {
        if (in_array($name, ['fecha_inicio', 'fecha_fin'], true)) {
            return $this->defaultValueFor($name, $definition);
        }

        return match ($type) {
            'number' => (int) ($definition['default'] ?? 0),
            'checkbox' => false,
            default => $definition['default'] ?? '',
        };
    }

    private function rowsToArrays(array $rows): array
    {
        return array_map(fn ($row) => (array) $row, $rows);
    }

    private function firstRow(array $rows): array
    {
        return $rows[0] ?? [];
    }

    private function value(array $row, string ...$keys): mixed
    {
        foreach ($keys as $key) {
            if (array_key_exists($key, $row)) {
                return $row[$key];
            }
        }

        return null;
    }

    private function minutesBetween(mixed $start, mixed $end): int
    {
        if (! $start || ! $end) {
            return 0;
        }

        return (int) Carbon::parse($start)->diffInMinutes(Carbon::parse($end), false);
    }

    private function legacySqlParam(mixed $param): mixed
    {
        if (is_string($param) && preg_match('/^\d{4}-\d{2}-\d{2}$/', $param)) {
            return Carbon::parse($param)->format('Y-m-d');
        }

        return $param;
    }

    private function parseIds(array|string|null $value): array
    {
        if (is_array($value)) {
            return array_values(array_unique(array_filter(array_map('intval', $value))));
        }

        if ($value === null) {
            return [];
        }

        return array_values(array_filter(array_map(
            fn (string $id) => (int) trim($id),
            preg_split('/[,;\\s]+/', $value) ?: []
        )));
    }

    private function employeeName(array $row): string
    {
        $name = $this->value($row, 'Apellidos', 'Nombres', 'Trabajador', 'Empleado', 'NombreCompleto');

        if ($name) {
            return (string) $name;
        }

        return trim(($this->value($row, 'Apellido') ?? '').' '.($this->value($row, 'Nombre') ?? ''));
    }

    private function baseTime(array $row): string
    {
        $value = $this->value($row, 'IngresoBaseContrato', 'HoraIngresoIdeal', 'HoraIngreso', 'Ingreso', 'Horario');

        if (! $value) {
            return '';
        }

        try {
            return Carbon::parse($value)->format('g.i a');
        } catch (\Throwable) {
            return (string) $value;
        }
    }

    private function dateKey(mixed $value): ?string
    {
        return $this->parseLegacyDate($value)?->toDateString();
    }

    private function attendanceMark(array $row): string
    {
        $text = $this->plainUpper($this->observationText($row));
        $late = (int) $this->value($row, 'Tardanza', 'MinutosTardanza', 'TARDANZA', 'TardanzaMinutos');

        if ($late > 0) {
            return (string) $late;
        }

        return match (true) {
            str_contains($text, 'VACACION') || str_contains($text, 'VACACIONES') => 'V',
            str_contains($text, 'DESCANSO MEDICO') || str_contains($text, 'DM') => 'DM',
            str_contains($text, 'PATERNIDAD') => 'LP',
            str_contains($text, 'PERMISO') => 'P',
            str_contains($text, 'SUSPENSION') => 'SSGH',
            str_contains($text, 'FALTA') || str_contains($text, 'INASIST') => 'F',
            default => $this->normalMark(),
        };
    }

    private function parseLegacyDate(mixed $value): ?Carbon
    {
        if (! $value) {
            return null;
        }

        if ($value instanceof \DateTimeInterface) {
            return Carbon::instance($value);
        }

        $value = trim((string) $value);

        if ($value === '') {
            return null;
        }

        if (preg_match('/^\d+$/', $value)) {
            return null;
        }

        if (preg_match('/^(\d{1,2})\/(\d{1,2})\/(\d{4})/', $value, $matches)) {
            try {
                return Carbon::createFromFormat('d/m/Y', "{$matches[1]}/{$matches[2]}/{$matches[3]}");
            } catch (\Throwable) {
                return null;
            }
        }

        try {
            return Carbon::parse($value);
        } catch (\Throwable) {
            return null;
        }
    }

    private function plainUpper(string $text): string
    {
        $text = mb_strtoupper($text);

        return strtr($text, [
            'Á' => 'A',
            'É' => 'E',
            'Í' => 'I',
            'Ó' => 'O',
            'Ú' => 'U',
            'Ü' => 'U',
            'Ñ' => 'N',
        ]);
    }

    private function normalMark(): string
    {
        return html_entity_decode('&bull;', ENT_QUOTES | ENT_HTML5, 'UTF-8');
    }

    private function observationText(array $row): string
    {
        foreach (['Observacion', 'Observaciones', 'Motivo', 'Justificacion', 'Descripcion', 'Concepto'] as $key) {
            $value = $this->value($row, $key);

            if ($value) {
                return trim((string) $value);
            }
        }

        return '';
    }

    private function spanishWeekday(Carbon $date): string
    {
        return ['Do', 'Lu', 'Ma', 'Mi', 'Ju', 'Vi', 'Sa'][$date->dayOfWeek];
    }
}
