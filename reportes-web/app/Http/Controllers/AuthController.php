<?php

namespace App\Http\Controllers;

use App\Services\LegacyCrudService;
use Illuminate\Http\Request;

class AuthController extends Controller
{
    public function __construct(private readonly LegacyCrudService $crud)
    {
    }

    public function showLogin()
    {
        if (session()->has('usuario_id')) {
            return redirect('/');
        }

        return view('auth.login');
    }

    public function login(Request $request)
    {
        $request->validate([
            'usuario' => 'required|string',
            'clave' => 'required|string',
        ]);

        try {
            $rows = $this->crud->callProcedure('up_BuscarValidaUsuario', [
                $request->input('usuario'),
                $request->input('clave'),
            ]);

            if (empty($rows)) {
                return back()->with('error', 'Usuario o contraseña incorrectos.')->withInput();
            }

            $user = $rows[0];
            $userId = (int) ($user['IdUsuario'] ?? $user['idusuario'] ?? 0);

            session([
                'usuario_id' => $userId,
                'usuario_nombre' => strtoupper($request->input('usuario')),
            ]);

            // Load menu permissions
            try {
                $permisos = $this->crud->callProcedure('up_BuscarAccesoUsuario', [$userId]);
                $accessString = '';
                foreach ($permisos as $p) {
                    $accessString .= ($p['Habilitado'] ?? $p['habilitado'] ?? false) ? '1' : '0';
                }
                session(['usuario_permisos' => $accessString]);
            } catch (\Throwable) {
                session(['usuario_permisos' => str_repeat('1', 30)]);
            }

            return redirect('/');
        } catch (\Throwable $e) {
            return back()->with('error', 'Error de conexión: ' . $e->getMessage())->withInput();
        }
    }

    public function logout()
    {
        session()->flush();

        return redirect('login');
    }
}
