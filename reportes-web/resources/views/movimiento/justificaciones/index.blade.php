@extends('layouts.app', ['title' => 'Justificaciones'])
@section('topbar-title', 'Justificaciones')
@section('breadcrumb', 'Movimientos / Justificaciones')

@section('content')
    <div class="page-header">
        <div>
            <h1>📝 Justificaciones</h1>
            <p class="subtitle">Justificación de inasistencias y tardanzas del personal</p>
        </div>
        <a class="btn btn-primary" href="{{ url('movimiento/justificaciones/crear') }}">＋ Nueva Justificación</a>
    </div>

    {{-- Filtros --}}
    <div class="card" style="padding:18px;margin-bottom:20px;border-top:3px solid var(--accent)">
        <form method="GET" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:14px;align-items:end">
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
                <label class="form-label">Grupo</label>
                <select name="id_grupo" class="form-control">
                    <option value="0">Todos</option>
                    @foreach ($grupos as $g)
                        <option value="{{ $g['value'] }}" @selected($filters['id_grupo'] == $g['value'])>{{ $g['label'] }}</option>
                    @endforeach
                </select>
            </div>
            <div class="form-group">
                <label class="form-label">Motivo</label>
                <select name="id_motivo" class="form-control">
                    <option value="0">Todos</option>
                    @foreach ($motivos as $m)
                        <option value="{{ $m['value'] }}" @selected($filters['id_motivo'] == $m['value'])>{{ $m['label'] }}</option>
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
            <div class="table-search"><input type="text" placeholder="Filtrar resultados..."></div>
            <span class="table-count">{{ count($rows) }} registros</span>
        </div>
        @if (empty($rows))
            <div class="empty-state"><div class="icon">📭</div><p>Ejecute una búsqueda para ver justificaciones.</p></div>
        @else
            <div style="overflow-x:auto">
                <table class="data-table">
                    <thead><tr>
                        <th>#</th>
                        @foreach (array_keys($rows[0]) as $col)
                            <th>{{ $col }}</th>
                        @endforeach
                        <th>Acciones</th>
                    </tr></thead>
                    <tbody>
                        @foreach ($rows as $i => $row)
                            <tr>
                                <td style="color:var(--muted);font-size:12px">{{ $i + 1 }}</td>
                                @foreach ($row as $val)
                                    <td>{{ is_scalar($val) ? $val : '' }}</td>
                                @endforeach
                                <td>
                                    <div class="table-actions">
                                        @php $rid = $row['IdJustificacion'] ?? $row['idjustificacion'] ?? 0; @endphp
                                        <button class="btn btn-danger btn-sm" onclick="confirmDelete('{{ url("movimiento/justificaciones/{$rid}") }}')">🗑️ Eliminar</button>
                                    </div>
                                </td>
                            </tr>
                        @endforeach
                    </tbody>
                </table>
            </div>
        @endif
    </div>
@endsection
