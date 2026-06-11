@extends('layouts.app', ['title' => ($mode === 'create' ? 'Nueva' : 'Editar') . ' Suspensión'])

@section('topbar-title', ($mode === 'create' ? 'Nueva' : 'Editar') . ' Suspensión')
@section('breadcrumb', 'Movimientos / Suspensiones / ' . ($mode === 'create' ? 'Nueva' : 'Editar'))

@section('content')
    <div class="page-header">
        <div>
            <h1>⛔ {{ $mode === 'create' ? 'Nueva' : 'Editar' }} Suspensión</h1>
            <p class="subtitle">Ingrese los datos de la suspensión de labor del trabajador</p>
        </div>
        <a class="btn btn-secondary" href="{{ url('movimiento/suspensiones') }}">← Volver al listado</a>
    </div>

    <div class="card form-card">
        <form method="POST" action="{{ $mode === 'create' ? url('movimiento/suspensiones') : url("movimiento/suspensiones/{$id}") }}">
            @csrf
            @if ($mode === 'edit')
                @method('PUT')
            @endif

            <div class="form-grid">
                <div class="form-group">
                    <label class="form-label">Trabajador <span class="required">*</span></label>
                    <select name="IdTrabajador" class="form-control" required>
                        <option value="">— Seleccione Trabajador —</option>
                        @foreach ($trabajadores as $t)
                            <option value="{{ $t['value'] }}" @selected($record && ($record['IdTrabajador'] ?? $record['idtrabajador'] ?? 0) == $t['value'])>{{ $t['label'] }}</option>
                        @endforeach
                    </select>
                </div>

                <div class="form-group">
                    <label class="form-label">Fecha Inicio <span class="required">*</span></label>
                    <input type="date" name="FechaInicio" class="form-control" required
                           value="{{ $record ? \Carbon\Carbon::parse($record['FechaInicio'] ?? $record['fechainicio'])->format('Y-m-d') : '' }}">
                </div>

                <div class="form-group">
                    <label class="form-label">Fecha Fin <span class="required">*</span></label>
                    <input type="date" name="FechaFin" class="form-control" required
                           value="{{ $record ? \Carbon\Carbon::parse($record['FechaFin'] ?? $record['fechafin'])->format('Y-m-d') : '' }}">
                </div>

                <div class="form-group full">
                    <label class="form-label">Motivo</label>
                    <input type="text" name="Motivo" class="form-control" placeholder="Escriba el motivo de la suspensión..."
                           value="{{ $record['Motivo'] ?? $record['motivo'] ?? '' }}">
                </div>
            </div>

            <div class="form-actions">
                <button type="submit" class="btn btn-primary">
                    {{ $mode === 'create' ? '＋ Registrar Suspensión' : '💾 Guardar Cambios' }}
                </button>
                <a class="btn btn-secondary" href="{{ url('movimiento/suspensiones') }}">Cancelar</a>
            </div>
        </form>
    </div>
@endsection
