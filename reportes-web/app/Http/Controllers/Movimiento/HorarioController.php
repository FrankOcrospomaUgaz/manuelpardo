<?php

namespace App\Http\Controllers\Movimiento;

use App\Http\Controllers\Controller;
use App\Services\LegacyCrudService;
use Illuminate\Http\Request;

class HorarioController extends Controller
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
            'fecha_inicio' => $request->query('fecha_inicio', now()->startOfMonth()->format('Y-m-d')),
            'fecha_fin' => $request->query('fecha_fin', now()->endOfMonth()->format('Y-m-d')),
        ];

        $rows = [];
        if ($request->has('buscar')) {
            try {
                $rows = $this->crud->callProcedure('up_BuscarHorarioxTrabajador', [
                    $filters['id_trabajador'],
                    $filters['id_grupo'],
                    $filters['fecha_inicio'],
                    $filters['fecha_fin'],
                ]);
            } catch (\Throwable $e) {
                session()->flash('error', 'Error: ' . $e->getMessage());
            }
        }

        return view('movimiento.horarios.index', compact('trabajadores', 'grupos', 'rows', 'filters'));
    }

    public function create()
    {
        $trabajadores = $this->crud->getLookup('trabajadores');
        $horarios = [];
        try {
            $horarios = $this->crud->callProcedure('up_BuscarHorario', [0]);
        } catch (\Throwable) {}

        return view('movimiento.horarios.form', [
            'mode' => 'create', 'trabajadores' => $trabajadores, 'horarios' => $horarios,
        ]);
    }

    public function store(Request $request)
    {
        try {
            $this->crud->callProcedure('up_AgregarHorario', [
                $request->input('Fecha'),
                (int) $request->input('IdTrabajador'),
                $request->input('HoraIngresoIdeal'),
                $request->input('HoraSalidaIdeal'),
                (int) $request->input('IdTipoHora', 1),
            ]);

            return redirect('movimiento/horarios')->with('success', 'Horario asignado correctamente.');
        } catch (\Throwable $e) {
            return back()->withInput()->with('error', 'Error: ' . $e->getMessage());
        }
    }

    public function destroy(int $id)
    {
        try {
            $this->crud->callProcedure('up_EliminarHorario', [$id]);

            return redirect('movimiento/horarios')->with('success', 'Horario eliminado.');
        } catch (\Throwable $e) {
            return back()->with('error', 'Error: ' . $e->getMessage());
        }
    }
}
