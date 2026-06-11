<?php

use App\Http\Controllers\AuthController;
use App\Http\Controllers\DashboardController;
use App\Http\Controllers\CrudController;
use App\Http\Controllers\ReportController;
use App\Http\Controllers\Movimiento\SuspensionController;
use App\Http\Controllers\Movimiento\JustificacionController;
use App\Http\Controllers\Movimiento\PermisoController;
use App\Http\Controllers\Movimiento\MarcacionController;
use App\Http\Controllers\Movimiento\HorarioController;
use App\Http\Controllers\Movimiento\ConsolidadoController;
use App\Http\Middleware\LegacyAuth;
use Illuminate\Support\Facades\Route;

// Autenticación
Route::get('/login', [AuthController::class, 'showLogin'])->name('login');
Route::post('/login', [AuthController::class, 'login']);
Route::post('/logout', [AuthController::class, 'logout'])->name('logout');

// Rutas Protegidas
Route::middleware([LegacyAuth::class])->group(function () {
    
    // Panel de Inicio
    Route::get('/', [DashboardController::class, 'index']);

    // Mantenimientos (CRUDs)
    Route::prefix('mantenimiento')->group(function () {
        Route::get('/{entity}', [CrudController::class, 'index']);
        Route::get('/{entity}/crear', [CrudController::class, 'create']);
        Route::post('/{entity}', [CrudController::class, 'store']);
        Route::get('/{entity}/{id}/editar', [CrudController::class, 'edit']);
        Route::put('/{entity}/{id}', [CrudController::class, 'update']);
        Route::delete('/{entity}/{id}', [CrudController::class, 'destroy']);
    });

    // Movimientos
    Route::prefix('movimiento')->group(function () {
        // Suspensiones
        Route::get('/suspensiones', [SuspensionController::class, 'index']);
        Route::get('/suspensiones/crear', [SuspensionController::class, 'create']);
        Route::post('/suspensiones', [SuspensionController::class, 'store']);
        Route::get('/suspensiones/{id}/editar', [SuspensionController::class, 'edit']);
        Route::put('/suspensiones/{id}', [SuspensionController::class, 'update']);
        Route::delete('/suspensiones/{id}', [SuspensionController::class, 'destroy']);

        // Justificaciones
        Route::get('/justificaciones', [JustificacionController::class, 'index']);
        Route::get('/justificaciones/crear', [JustificacionController::class, 'create']);
        Route::post('/justificaciones', [JustificacionController::class, 'store']);
        Route::delete('/justificaciones/{id}', [JustificacionController::class, 'destroy']);

        // Permisos
        Route::get('/permisos', [PermisoController::class, 'index']);
        Route::get('/permisos/crear', [PermisoController::class, 'create']);
        Route::post('/permisos', [PermisoController::class, 'store']);

        // Marcaciones
        Route::get('/marcacion', [MarcacionController::class, 'index']);
        Route::post('/marcacion/{id}', [MarcacionController::class, 'updateEntry']);

        // Horarios
        Route::get('/horarios', [HorarioController::class, 'index']);
        Route::get('/horarios/crear', [HorarioController::class, 'create']);
        Route::post('/horarios', [HorarioController::class, 'store']);
        Route::delete('/horarios/{id}', [HorarioController::class, 'destroy']);

        // Consolidado
        Route::get('/consolidado', [ConsolidadoController::class, 'index']);
    });

    // Reportes (Existentes y funcionales)
    Route::get('/reportes', [ReportController::class, 'index'])->name('reports.index');
    Route::get('/reportes/{slug}', [ReportController::class, 'show'])->name('reports.show');
    Route::post('/reportes/{slug}', [ReportController::class, 'run'])->name('reports.run');
    Route::get('/reportes/{slug}/exportar', [ReportController::class, 'export'])->name('reports.export');
    Route::get('/reportes/{slug}/excel', [ReportController::class, 'excel'])->name('reports.excel');
    Route::get('/reportes/{slug}/pdf', [ReportController::class, 'pdf'])->name('reports.pdf');
});
