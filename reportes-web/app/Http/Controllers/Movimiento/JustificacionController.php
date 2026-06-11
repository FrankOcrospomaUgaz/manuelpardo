<?php

namespace App\Http\Controllers\Movimiento;

use App\Http\Controllers\Controller;
use App\Services\LegacyCrudService;
use Illuminate\Http\Request;

class JustificacionController extends Controller
{
    public function __construct(private readonly LegacyCrudService $crud)
    {
    }

    public function index(Request $request)
    {
        $trabajadores = $this->crud->getLookup('trabajadores');
        $grupos = $this->crud->getLookup('grupos');
        $motivos = $this->crud->getLookup('motivos');

        $filters = [
            'id_trabajador' => (int) $request->query('id_trabajador', 0),
            'id_grupo' => (int) $request->query('id_grupo', 0),
            'id_motivo' => (int) $request->query('id_motivo', 0),
            'fecha_inicio' => $request->query('fecha_inicio', ''),
            'fecha_fin' => $request->query('fecha_fin', ''),
        ];

        $rows = [];
        if ($request->has('buscar')) {
            try {
                $rows = $this->crud->callProcedure('up_BuscarJustificacion', [
                    0, $filters['id_trabajador'], $filters['id_grupo'],
                    $filters['fecha_inicio'] ?: '01/01/1900',
                    $filters['fecha_fin'] ?: '31/12/2099',
                ]);
            } catch (\Throwable $e) {
                session()->flash('error', 'Error: ' . $e->getMessage());
            }
        }

        return view('movimiento.justificaciones.index', compact('trabajadores', 'grupos', 'motivos', 'rows', 'filters'));
    }

    public function create()
    {
        $trabajadores = $this->crud->getLookup('trabajadores');
        $motivos = $this->crud->getLookup('motivos');

        return view('movimiento.justificaciones.form', [
            'mode' => 'create', 'record' => null,
            'trabajadores' => $trabajadores, 'motivos' => $motivos,
        ]);
    }

    public function store(Request $request)
    {
        try {
            $this->crud->callProcedure('up_AgregarJustificacion', [
                $request->input('Fecha'),
                (int) $request->input('IdTrabajador'),
                (int) $request->input('IdMotivoJustificacion'),
                $request->input('FechaInicio'),
                $request->input('FechaFin'),
                $request->input('Observacion', ''),
            ]);

            return redirect('movimiento/justificaciones')->with('success', 'Justificación registrada correctamente.');
        } catch (\Throwable $e) {
            return back()->withInput()->with('error', 'Error: ' . $e->getMessage());
        }
    }

    public function destroy(int $id)
    {
        try {
            $this->crud->callProcedure('up_EliminarJustificacion', [$id]);

            return redirect('movimiento/justificaciones')->with('success', 'Justificación eliminada.');
        } catch (\Throwable $e) {
            return back()->with('error', 'Error: ' . $e->getMessage());
        }
    }
}
