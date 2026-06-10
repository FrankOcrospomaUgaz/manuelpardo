<?php

use App\Http\Controllers\ReportController;
use Illuminate\Support\Facades\Route;

Route::redirect('/', '/reportes');

Route::get('/reportes', [ReportController::class, 'index'])->name('reports.index');
Route::get('/reportes/{slug}', [ReportController::class, 'show'])->name('reports.show');
Route::post('/reportes/{slug}', [ReportController::class, 'run'])->name('reports.run');
Route::get('/reportes/{slug}/exportar', [ReportController::class, 'export'])->name('reports.export');
Route::get('/reportes/{slug}/excel', [ReportController::class, 'excel'])->name('reports.excel');
Route::get('/reportes/{slug}/pdf', [ReportController::class, 'pdf'])->name('reports.pdf');
