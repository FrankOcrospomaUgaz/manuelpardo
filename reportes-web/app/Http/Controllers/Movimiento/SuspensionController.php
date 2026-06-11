<?php

namespace App\Http\Controllers\Movimiento;

use App\Http\Controllers\Controller;
use App\Services\LegacyCrudService;
use Illuminate\Http\Request;

class SuspensionController extends Controller
{
    public function __construct(private readonly LegacyCrudService $crud)
    {
    }

    public function index(Request $request)
    {
        $trabajadores = $this->crud->getLookup('trabajadores');
        $grupos = $this->crud->getLookup('grupos');

        $idTrabajador = (int) $request->query('id_trabajador', 0);
        $idGrupo = (int) $request->query('id_grupo', 0);
        $fechaInicio = $request->query('fecha_inicio', '');
        $fechaFin = $request->query('fecha_fin', '');

        $rows = [];
        if ($request->has('buscar')) {
            try {
                $rows = $this->crud->callProcedure('up_BuscarSuspension', [
                    0, $idTrabajador, $idGrupo,
                    $fechaInicio ?: '01/01/1900',
                    $fechaFin ?: '31/12/2099',
                ]);
            } catch (\Throwable $e) {
                session()->flash('error', 'Error: ' . $e->getMessage());
            }
        }

        return view('movimiento.suspensiones.index', compact(
            'trabajadores', 'grupos', 'rows',
            'idTrabajador', 'idGrupo', 'fechaInicio', 'fechaFin'
        ));
    }

    public function create()
    {
        $trabajadores = $this->crud->getLookup('trabajadores');

        return view('movimiento.suspensiones.form', [
            'mode' => 'create',
            'record' => null,
            'trabajadores' => $trabajadores,
        ]);
    }

    public function store(Request $request)
    {
        try {
            $this->crud->callProcedure('up_AgregarSuspension', [
                (int) $request->input('IdTrabajador'),
                $request->input('FechaInicio'),
                $request->input('FechaFin'),
                $request->input('Motivo', ''),
            ]);

            return redirect('movimiento/suspensiones')->with('success', 'Suspensión registrada correctamente.');
        } catch (\Throwable $e) {
            return back()->withInput()->with('error', 'Error: ' . $e->getMessage());
        }
    }

    public function edit(int $id)
    {
        $trabajadores = $this->crud->getLookup('trabajadores');
        $rows = $this->crud->callProcedure('up_BuscarSuspension', [$id, 0, 0, '01/01/1900', '31/12/2099']);
        $record = $rows[0] ?? null;

        if (! $record) {
            return redirect('movimiento/suspensiones')->with('error', 'Registro no encontrado.');
        }

        return view('movimiento.suspensiones.form', [
            'mode' => 'edit',
            'record' => $record,
            'id' => $id,
            'trabajadores' => $trabajadores,
        ]);
    }

    public function update(Request $request, int $id)
    {
        try {
            $this->crud->callProcedure('up_ModificarSuspension', [
                $id,
                (int) $request->input('IdTrabajador'),
                $request->input('FechaInicio'),
                $request->input('FechaFin'),
                $request->input('Motivo', ''),
            ]);

            return redirect('movimiento/suspensiones')->with('success', 'Suspensión actualizada.');
        } catch (\Throwable $e) {
            return back()->withInput()->with('error', 'Error: ' . $e->getMessage());
        }
    }

    public function destroy(int $id)
    {
        try {
            $this->crud->callProcedure('up_EliminarSuspension', [$id]);

            return redirect('movimiento/suspensiones')->with('success', 'Suspensión eliminada.');
        } catch (\Throwable $e) {
            return back()->with('error', 'Error: ' . $e->getMessage());
        }
    }
}
