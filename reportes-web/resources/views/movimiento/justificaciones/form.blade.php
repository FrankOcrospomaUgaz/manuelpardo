@extends('layouts.app', ['title' => 'Nueva Justificación'])

@section('topbar-title', 'Nueva Justificación')
@section('breadcrumb', 'Movimientos / Justificaciones / Nueva')

@section('content')
    <div class="page-header">
        <div>
            <h1>📝 Nueva Justificación</h1>
            <p class="subtitle">Ingrese los detalles para justificar la inasistencia o tardanza</p>
        </div>
        <a class="btn btn-secondary" href="{{ url('movimiento/justificaciones') }}">← Volver al listado</a>
    </div>

    <div class="card form-card">
        <form method="POST" action="{{ url('movimiento/justificaciones') }}">
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
                    <label class="form-label">Motivo de Justificación <span class="required">*</span></label>
                    <select name="IdMotivoJustificacion" class="form-control" required>
                        <option value="">— Seleccione Motivo —</option>
                        @foreach ($motivos as $m)
                            <option value="{{ $m['value'] }}">{{ $m['label'] }}</option>
                        @endforeach
                    </select>
                </div>

                <div class="form-group">
                    <label class="form-label">Fecha de Registro <span class="required">*</span></label>
                    <input type="date" name="Fecha" class="form-control" required value="{{ date('Y-m-d') }}">
                </div>

                <div class="form-group">
                    <label class="form-label">Desde (Fecha) <span class="required">*</span></label>
                    <input type="date" name="FechaInicio" class="form-control" required value="{{ date('Y-m-d') }}">
                </div>

                <div class="form-group">
                    <label class="form-label">Hasta (Fecha) <span class="required">*</span></label>
                    <input type="date" name="FechaFin" class="form-control" required value="{{ date('Y-m-d') }}">
                </div>

                <div class="form-group full">
                    <label class="form-label">Observación</label>
                    <input type="text" name="Observacion" class="form-control" placeholder="Escriba los detalles de la justificación...">
                </div>
            </div>

            <div class="form-actions">
                <button type="submit" class="btn btn-primary">＋ Registrar Justificación</button>
                <a class="btn btn-secondary" href="{{ url('movimiento/justificaciones') }}">Cancelar</a>
            </div>
        </form>
    </div>
@endsection
