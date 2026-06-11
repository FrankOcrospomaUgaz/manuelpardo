@extends('layouts.app', ['title' => 'Registrar Permiso'])

@section('topbar-title', 'Registrar Permiso')
@section('breadcrumb', 'Movimientos / Permisos / Nuevo')

@section('content')
    <div class="page-header">
        <div>
            <h1>🎫 Registrar Permiso</h1>
            <p class="subtitle">Ingrese los datos del permiso temporal por horas</p>
        </div>
        <a class="btn btn-secondary" href="{{ url('movimiento/permisos') }}">← Volver al listado</a>
    </div>

    <div class="card form-card">
        <form method="POST" action="{{ url('movimiento/permisos') }}">
            @csrf

            <div class="form-grid">
                <div class="form-group">
                    <label class="form-label">Trabajador <span class="required">*</span></label>
                    <select name="IdTrabajador" class="form-control" required>
                        <option value="">— Seleccione Trabajador —</option>
                        @foreach ($trabajadores as $t)
                            <option value="{{ $t['value'] }}">{{ $t['label'] }}</option>
                        @endforeach
                    </select>
                </div>

                <div class="form-group">
                    <label class="form-label">Fecha del Permiso <span class="required">*</span></label>
                    <input type="date" name="Fecha" class="form-control" required value="{{ date('Y-m-d') }}">
                </div>

                <div class="form-group">
                    <label class="form-label">Hora Inicio <span class="required">*</span></label>
                    <input type="time" name="HoraInicio" class="form-control" required value="08:00">
                </div>

                <div class="form-group">
                    <label class="form-label">Hora Fin <span class="required">*</span></label>
                    <input type="time" name="HoraFin" class="form-control" required value="12:00">
                </div>

                <div class="form-group full">
                    <label class="form-label">Motivo</label>
                    <input type="text" name="Motivo" class="form-control" placeholder="Escriba el motivo del permiso...">
                </div>
            </div>

            <div class="form-actions">
                <button type="submit" class="btn btn-primary">＋ Registrar Permiso</button>
                <a class="btn btn-secondary" href="{{ url('movimiento/permisos') }}">Cancelar</a>
            </div>
        </form>
    </div>
@endsection
