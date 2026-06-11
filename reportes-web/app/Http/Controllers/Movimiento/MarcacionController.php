<?php

namespace App\Http\Controllers\Movimiento;

use App\Http\Controllers\Controller;
use App\Services\LegacyCrudService;
use Illuminate\Http\Request;

class MarcacionController extends Controller
{
    public function __construct(private readonly LegacyCrudService $crud)
    {
    }

    public function index(Request $request)
    {
        $trabajadores = $this->crud->getLookup('trabajadores');
        $filters = [
            'id_trabajador' => (int) $request->query('id_trabajador', 0),
            'fecha_inicio' => $request->query('fecha_inicio', now()->format('Y-m-d')),
            'fecha_fin' => $request->query('fecha_fin', now()->format('Y-m-d')),
        ];

        $rows = [];
        if ($request->has('buscar')) {
            try {
                $rows = $this->crud->callProcedure('up_BuscarMarcacion', [
                    $filters['id_trabajador'],
                    $filters['fecha_inicio'],
                    $filters['fecha_fin'],
                ]);
            } catch (\Throwable $e) {
                session()->flash('error', 'Error: ' . $e->getMessage());
            }
        }

        return view('movimiento.marcacion.index', compact('trabajadores', 'rows', 'filters'));
    }

    public function updateEntry(Request $request, int $id)
    {
        try {
            $this->crud->callProcedure('up_ModificarHoraEntradaYSalidaMarcacion', [
                $id,
                $request->input('HoraIngresoReal'),
                $request->input('HoraSalidaReal'),
            ]);

            return back()->with('success', 'Marcación actualizada.');
        } catch (\Throwable $e) {
            return back()->with('error', 'Error: ' . $e->getMessage());
        }
    }
}
