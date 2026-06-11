<?php

namespace App\Http\Controllers\Movimiento;

use App\Http\Controllers\Controller;
use App\Services\LegacyCrudService;
use Illuminate\Http\Request;

class PermisoController extends Controller
{
    public function __construct(private readonly LegacyCrudService $crud)
    {
    }

    public function index(Request $request)
    {
        $trabajadores = $this->crud->getLookup('trabajadores');
        $grupos = $this->crud->getLookup('grupos');

        $filters = [
            'id_trabajador' => (int) $request->query('id_trabajador', 0),
            'id_grupo' => (int) $request->query('id_grupo', 0),
            'fecha_inicio' => $request->query('fecha_inicio', ''),
            'fecha_fin' => $request->query('fecha_fin', ''),
        ];

        $rows = [];
        if ($request->has('buscar')) {
            try {
                $rows = $this->crud->callProcedure('up_BuscarPermiso', [
                    0, $filters['id_trabajador'],
                    $filters['fecha_inicio'] ?: '01/01/1900',
                    $filters['fecha_fin'] ?: '31/12/2099',
                ]);
            } catch (\Throwable $e) {
                session()->flash('error', 'Error: ' . $e->getMessage());
            }
        }

        return view('movimiento.permisos.index', compact('trabajadores', 'grupos', 'rows', 'filters'));
    }

    public function create()
    {
        $trabajadores = $this->crud->getLookup('trabajadores');

        return view('movimiento.permisos.form', [
            'mode' => 'create', 'record' => null, 'trabajadores' => $trabajadores,
        ]);
    }

    public function store(Request $request)
    {
        try {
            $this->crud->callProcedure('up_AgregarPermiso', [
                $request->input('Fecha'),
                (int) $request->input('IdTrabajador'),
                $request->input('HoraInicio'),
                $request->input('HoraFin'),
                $request->input('Motivo', ''),
            ]);

            return redirect('movimiento/permisos')->with('success', 'Permiso registrado correctamente.');
        } catch (\Throwable $e) {
            return back()->withInput()->with('error', 'Error: ' . $e->getMessage());
        }
    }
}
