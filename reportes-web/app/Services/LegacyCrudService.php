<?php

namespace App\Services;

use Carbon\Carbon;
use Illuminate\Support\Facades\DB;

class LegacyCrudService
{
    /**
     * Execute a stored procedure and return rows as arrays.
     */
    public function callProcedure(string $procedure, array $params = []): array
    {
        $params = array_map(fn(mixed $p) => $this->formatParam($p), $params);
        $placeholders = implode(', ', array_fill(0, count($params), '?'));
        $sql = trim("EXEC {$procedure} {$placeholders}");

        return array_map(fn($r) => (array) $r, DB::select($sql, $params));
    }

    /**
     * List all records for an entity.
     */
    public function listAll(array $config): array
    {
        $sp = $config['sp_buscar'];
        if (str_starts_with(strtoupper(trim($sp)), 'SELECT')) {
            return array_map(fn($r) => (array) $r, DB::select($sp));
        }
        $defaultParam = $config['buscar_default'] ?? [0];

        return $this->callProcedure($sp, $defaultParam);
    }

    /**
     * Get a single record by ID.
     */
    public function findById(array $config, int $id): ?array
    {
        $sp = $config['sp_buscar'];
        if (str_starts_with(strtoupper(trim($sp)), 'SELECT')) {
            $sql = $sp . " AND " . $config['id_column'] . " = ?";
            $rows = array_map(fn($r) => (array) $r, DB::select($sql, [$id]));
            return $rows[0] ?? null;
        }
        $rows = $this->callProcedure($sp, [$id]);

        return $rows[0] ?? null;
    }

    /**
     * Create a new record.
     */
    public function create(array $config, array $data): bool
    {
        $params = $this->buildParams($config['fields'], $data, 'create');

        return $this->executeProcedure($config['sp_agregar'], $params);
    }

    /**
     * Update an existing record.
     */
    public function update(array $config, int $id, array $data): bool
    {
        $params = array_merge([$id], $this->buildParams($config['fields'], $data, 'update'));

        return $this->executeProcedure($config['sp_modificar'], $params);
    }

    /**
     * Delete a record by ID.
     */
    public function delete(array $config, int $id): bool
    {
        return $this->executeProcedure($config['sp_eliminar'], [$id]);
    }

    /**
     * Get lookup data for select dropdowns.
     */
    public function getLookup(string $type): array
    {
        return match ($type) {
            'areas' => $this->simpleLookup('up_BuscarArea', 'IdArea', ['Descripcion']),
            'cargos' => $this->simpleLookup('up_BuscarCargo', 'IdCargo', ['Descripcion']),
            'categorias' => $this->simpleLookup('up_BuscarCategoria', 'IdCategoria', ['Descripcion']),
            'grupos' => $this->simpleLookup('up_BuscarGrupo', 'IdGrupo', ['Descripcion', 'Grupo']),
            'motivos' => $this->simpleLookup('up_BuscarMotivo', 'IdMotivoJustificacion', ['Descripcion']),
            'tipos_hora' => $this->tableLookup('TipoHora', 'IdTipoHora', 'Descripcion'),
            'trabajadores' => $this->employeeLookup(),
            'usuarios' => $this->simpleLookup('up_BuscarUsuario', 'IdUsuario', ['Nombre']),
            'horarios' => $this->tableLookup('Horario', 'IdHorario', 'Descripcion'),
            default => [],
        };
    }

    /**
     * Build ordered parameters from form data based on field config.
     */
    private function buildParams(array $fields, array $data, string $action): array
    {
        $params = [];
        foreach ($fields as $name => $field) {
            if (($field['exclude_on'] ?? null) === $action) {
                continue;
            }
            if ($field['only_on'] ?? null) {
                if ($field['only_on'] !== $action) {
                    continue;
                }
            }

            $value = $data[$name] ?? ($field['default'] ?? '');

            // Type casting
            $params[] = match ($field['type'] ?? 'text') {
                'integer', 'number' => (int) $value,
                'date' => $value ?: null,
                'boolean' => $value ? 1 : 0,
                default => (string) $value,
            };
        }

        return $params;
    }

    /**
     * Execute a procedure that returns true/false (INSERT/UPDATE/DELETE).
     */
    private function executeProcedure(string $procedure, array $params): bool
    {
        try {
            $params = array_map(fn(mixed $p) => $this->formatParam($p), $params);
            $placeholders = implode(', ', array_fill(0, count($params), '?'));
            DB::statement("EXEC {$procedure} {$placeholders}", $params);

            return true;
        } catch (\Throwable $e) {
            throw $e;
        }
    }

    private function simpleLookup(string $sp, string $idCol, array $labelCols): array
    {
        $rows = $this->callProcedure($sp, [0]);

        return collect($rows)->map(function (array $row) use ($idCol, $labelCols) {
            $id = 0;
            foreach ([$idCol, strtolower($idCol), strtoupper($idCol)] as $key) {
                if (isset($row[$key])) { $id = (int) $row[$key]; break; }
            }

            $label = '';
            foreach ($labelCols as $col) {
                foreach ([$col, strtolower($col), strtoupper($col)] as $key) {
                    if (isset($row[$key]) && $row[$key] !== '') { $label = (string) $row[$key]; break 2; }
                }
            }

            return ['value' => $id, 'label' => $label];
        })->filter(fn($r) => $r['value'] > 0 && $r['label'] !== '')->values()->all();
    }

    private function tableLookup(string $table, string $idCol, string $labelCol): array
    {
        return collect(DB::table($table)->select($idCol, $labelCol)->orderBy($labelCol)->get()->all())
            ->map(fn($r) => ['value' => (int) $r->{$idCol}, 'label' => (string) $r->{$labelCol}])
            ->filter(fn($r) => $r['value'] > 0 && $r['label'] !== '')
            ->values()->all();
    }

    private function employeeLookup(): array
    {
        return collect(DB::table('Trabajador')->select('IdTrabajador', 'Apellido', 'Nombre')->orderBy('Apellido')->orderBy('Nombre')->get()->all())
            ->map(fn($r) => ['value' => (int) $r->IdTrabajador, 'label' => trim($r->Apellido . ' ' . $r->Nombre)])
            ->filter(fn($r) => $r['value'] > 0 && $r['label'] !== '')
            ->values()->all();
    }

    private function formatParam(mixed $param): mixed
    {
        if (is_string($param) && preg_match('/^\d{4}-\d{2}-\d{2}$/', $param)) {
            return Carbon::parse($param)->format('Y-m-d');
        }

        return $param;
    }
}
