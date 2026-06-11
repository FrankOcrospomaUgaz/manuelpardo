@extends('layouts.app', ['title' => 'Marcaciones'])
@section('topbar-title', 'Marcaciones')
@section('breadcrumb', 'Movimientos / Marcaciones')

@section('content')
    <div class="page-header">
        <div>
            <h1>⏰ Marcaciones de Asistencia</h1>
            <p class="subtitle">Visualización y corrección manual de horas de entrada y salida</p>
        </div>
    </div>

    {{-- Filtros --}}
    <div class="card" style="padding:18px;margin-bottom:20px;border-top:3px solid var(--accent)">
        <form method="GET" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:14px;align-items:end">
            <input type="hidden" name="buscar" value="1">
            <div class="form-group">
                <label class="form-label">Trabajador</label>
                <select name="id_trabajador" class="form-control">
                    <option value="0">Todos</option>
                    @foreach ($trabajadores as $t)
                        <option value="{{ $t['value'] }}" @selected($filters['id_trabajador'] == $t['value'])>{{ $t['label'] }}</option>
                    @endforeach
                </select>
            </div>
            <div class="form-group">
                <label class="form-label">Fecha Inicio</label>
                <input type="date" name="fecha_inicio" value="{{ $filters['fecha_inicio'] }}" class="form-control">
            </div>
            <div class="form-group">
                <label class="form-label">Fecha Fin</label>
                <input type="date" name="fecha_fin" value="{{ $filters['fecha_fin'] }}" class="form-control">
            </div>
            <div><button type="submit" class="btn btn-primary" style="width:100%">🔍 Buscar</button></div>
        </form>
    </div>

    {{-- Resultados --}}
    <div class="card table-card">
        <div class="table-toolbar">
            <div class="table-search"><input type="text" placeholder="Filtrar marcaciones..."></div>
            <span class="table-count">{{ count($rows) }} marcaciones</span>
        </div>
        @if (empty($rows))
            <div class="empty-state"><div class="icon">⏰</div><p>Ejecute una búsqueda para visualizar marcaciones.</p></div>
        @else
            <div style="overflow-x:auto">
                <table class="data-table">
                    <thead><tr>
                        <th>#</th>
                        <th>Fecha</th>
                        <th>DNI / Tarjeta</th>
                        <th>Trabajador</th>
                        <th>Hora Ingreso Ideal</th>
                        <th>Hora Salida Ideal</th>
                        <th>Hora Ingreso Real</th>
                        <th>Hora Salida Real</th>
                        <th>Tardanza (Min)</th>
                        <th>Acción</th>
                    </tr></thead>
                    <tbody>
                        @foreach ($rows as $i => $row)
                            @php
                                $mid = $row['IdMarcacion'] ?? $row['idmarcacion'] ?? 0;
                                $fecha = $row['Fecha'] ?? $row['fecha'] ?? '';
                                try { $fecha = \Carbon\Carbon::parse($fecha)->format('d/m/Y'); } catch(\Throwable $e){}
                                $nombre = trim(($row['Apellido'] ?? $row['apellido'] ?? '') . ' ' . ($row['Nombre'] ?? $row['nombre'] ?? ''));
                                if ($nombre === '') $nombre = $row['Trabajador'] ?? $row['trabajador'] ?? '';
                            @endphp
                            <tr>
                                <td style="color:var(--muted);font-size:12px">{{ $i + 1 }}</td>
                                <td>{{ $fecha }}</td>
                                <td>{{ $row['Dni'] ?? $row['dni'] ?? $row['Tarjeta'] ?? $row['tarjeta'] ?? '—' }}</td>
                                <strong><td>{{ $nombre }}</td></strong>
                                <td style="color:var(--muted)">{{ $row['HoraIngresoIdeal'] ?? $row['horaingresoideal'] ?? '—' }}</td>
                                <td style="color:var(--muted)">{{ $row['HoraSalidaIdeal'] ?? $row['horasalidaideal'] ?? '—' }}</td>
                                <form method="POST" action="{{ url("movimiento/marcacion/{$mid}") }}">
                                    @csrf
                                    <td>
                                        <input type="text" name="HoraIngresoReal" class="form-control" style="width:90px;height:32px;padding:4px"
                                               value="{{ $row['HoraIngresoReal'] ?? $row['horaingresoreal'] ?? '' }}">
                                    </td>
                                    <td>
                                        <input type="text" name="HoraSalidaReal" class="form-control" style="width:90px;height:32px;padding:4px"
                                               value="{{ $row['HoraSalidaReal'] ?? $row['horasalidareal'] ?? '' }}">
                                    </td>
                                    <td>
                                        @php $tard = (int)($row['Tardanza'] ?? $row['tardanza'] ?? 0); @endphp
                                        <span class="badge {{ $tard > 0 ? 'badge-danger' : 'badge-success' }}">{{ $tard }} min</span>
                                    </td>
                                    <td>
                                        <button type="submit" class="btn btn-primary btn-sm">💾 Guardar</button>
                                    </td>
                                </form>
                            </tr>
                        @endforeach
                    </tbody>
                </table>
            </div>
        @endif
    </div>
@endsection
