<?php

namespace App\Http\Controllers\Movimiento;

use App\Http\Controllers\Controller;
use App\Services\LegacyCrudService;
use Illuminate\Http\Request;

class ConsolidadoController extends Controller
{
    public function __construct(private readonly LegacyCrudService $crud)
    {
    }

    public function index(Request $request)
    {
        $mes = (int) $request->query('mes', date('m'));
        $anio = (int) $request->query('anio', date('Y'));

        $rows = [];
        $exists = false;

        if ($request->has('buscar')) {
            try {
                $rows = $this->crud->callProcedure('up_ReporteConsolidados', [$mes, $anio]);
            } catch (\Throwable $e) {
                session()->flash('error', 'Error: ' . $e->getMessage());
            }
        }

        return view('movimiento.consolidado.index', compact('mes', 'anio', 'rows'));
    }
}
