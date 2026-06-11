<?php

namespace App\Http\Controllers;

use App\Services\LegacyCrudService;
use Illuminate\Http\Request;

class CrudController extends Controller
{
    public function __construct(private readonly LegacyCrudService $crud)
    {
    }

    /**
     * List all records.
     */
    public function index(string $entity)
    {
        $config = $this->resolveConfig($entity);

        try {
            $rows = $this->crud->listAll($config);
        } catch (\Throwable $e) {
            $rows = [];
            session()->flash('error', 'Error al cargar datos: ' . $e->getMessage());
        }

        return view('crud.index', compact('entity', 'config', 'rows'));
    }

    /**
     * Show create form.
     */
    public function create(string $entity)
    {
        $config = $this->resolveConfig($entity);
        $lookups = $this->loadLookups($config);

        return view('crud.form', [
            'entity' => $entity,
            'config' => $config,
            'record' => null,
            'lookups' => $lookups,
            'mode' => 'create',
        ]);
    }

    /**
     * Store a new record.
     */
    public function store(Request $request, string $entity)
    {
        $config = $this->resolveConfig($entity);

        try {
            $this->crud->create($config, $request->all());

            return redirect("mantenimiento/{$entity}")
                ->with('success', "{$config['singular']} creado(a) correctamente.");
        } catch (\Throwable $e) {
            return back()->withInput()
                ->with('error', 'Error al crear: ' . $e->getMessage());
        }
    }

    /**
     * Show edit form.
     */
    public function edit(string $entity, int $id)
    {
        $config = $this->resolveConfig($entity);
        $record = $this->crud->findById($config, $id);

        if (! $record) {
            return redirect("mantenimiento/{$entity}")
                ->with('error', 'Registro no encontrado.');
        }

        $lookups = $this->loadLookups($config);

        return view('crud.form', [
            'entity' => $entity,
            'config' => $config,
            'record' => $record,
            'lookups' => $lookups,
            'mode' => 'edit',
            'id' => $id,
        ]);
    }

    /**
     * Update an existing record.
     */
    public function update(Request $request, string $entity, int $id)
    {
        $config = $this->resolveConfig($entity);

        try {
            $this->crud->update($config, $id, $request->all());

            return redirect("mantenimiento/{$entity}")
                ->with('success', "{$config['singular']} actualizado(a) correctamente.");
        } catch (\Throwable $e) {
            return back()->withInput()
                ->with('error', 'Error al actualizar: ' . $e->getMessage());
        }
    }

    /**
     * Delete a record.
     */
    public function destroy(string $entity, int $id)
    {
        $config = $this->resolveConfig($entity);

        try {
            $this->crud->delete($config, $id);

            return redirect("mantenimiento/{$entity}")
                ->with('success', "{$config['singular']} eliminado(a) correctamente.");
        } catch (\Throwable $e) {
            return back()
                ->with('error', 'Error al eliminar: ' . $e->getMessage());
        }
    }

    /**
     * Resolve entity config or abort.
     */
    private function resolveConfig(string $entity): array
    {
        $config = config("crud.{$entity}");

        if (! $config) {
            abort(404, "Entidad '{$entity}' no configurada.");
        }

        return $config;
    }

    /**
     * Load all lookup data needed by the form.
     */
    private function loadLookups(array $config): array
    {
        $lookups = [];

        foreach ($config['fields'] as $field) {
            $lookup = $field['lookup'] ?? null;
            if ($lookup && ! isset($lookups[$lookup])) {
                try {
                    $lookups[$lookup] = $this->crud->getLookup($lookup);
                } catch (\Throwable) {
                    $lookups[$lookup] = [];
                }
            }
        }

        return $lookups;
    }
}
