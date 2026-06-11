<?php

namespace App\Http\Controllers;

use App\Services\LegacyCrudService;

class DashboardController extends Controller
{
    public function __construct(private readonly LegacyCrudService $crud)
    {
    }

    public function index()
    {
        $stats = [];

        try {
            $stats['trabajadores'] = count($this->crud->callProcedure('up_BuscarTrabajador', [0, '', 0, 0, 0]));
        } catch (\Throwable) {
            $stats['trabajadores'] = '—';
        }

        try {
            $stats['areas'] = count($this->crud->callProcedure('up_BuscarArea', [0]));
        } catch (\Throwable) {
            $stats['areas'] = '—';
        }

        try {
            $stats['grupos'] = count($this->crud->callProcedure('up_BuscarGrupo', [0]));
        } catch (\Throwable) {
            $stats['grupos'] = '—';
        }

        try {
            $stats['usuarios'] = count($this->crud->callProcedure('up_BuscarUsuario', [0, '']));
        } catch (\Throwable) {
            $stats['usuarios'] = '—';
        }

        return view('dashboard', compact('stats'));
    }
}
