@extends('layouts.app', ['title' => 'Asignar Horario'])

@section('topbar-title', 'Asignar Horario')
@section('breadcrumb', 'Movimientos / Horarios / Asignar')

@section('content')
    <div class="page-header">
        <div>
            <h1>📅 Asignar Horario</h1>
            <p class="subtitle">Asigne horario ideal de entrada y salida a un trabajador</p>
        </div>
        <a class="btn btn-secondary" href="{{ url('movimiento/horarios') }}">← Volver al listado</a>
    </div>

    <div class="card form-card">
        <form method="POST" action="{{ url('movimiento/horarios') }}">
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
                    <label class="form-label">Fecha <span class="required">*</span></label>
                    <input type="date" name="Fecha" class="form-control" required value="{{ date('Y-m-d') }}">
                </div>

                <div class="form-group">
                    <label class="form-label">Hora Ingreso Ideal <span class="required">*</span></label>
                    <input type="time" name="HoraIngresoIdeal" class="form-control" required value="08:00">
                </div>

                <div class="form-group">
                    <label class="form-label">Hora Salida Ideal <span class="required">*</span></label>
                    <input type="time" name="HoraSalidaIdeal" class="form-control" required value="17:00">
                </div>
            </div>

            <div class="form-actions">
                <button type="submit" class="btn btn-primary">＋ Asignar Horario</button>
                <a class="btn btn-secondary" href="{{ url('movimiento/horarios') }}">Cancelar</a>
            </div>
        </form>
    </div>
@endsection
