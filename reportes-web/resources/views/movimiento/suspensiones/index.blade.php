@extends('layouts.app', ['title' => 'Suspensiones'])
@section('topbar-title', 'Suspensiones')
@section('breadcrumb', 'Movimientos / Suspensiones')

@section('content')
    <div class="page-header">
        <div>
            <h1>⛔ Suspensiones</h1>
            <p class="subtitle">Gestión de suspensiones de trabajadores</p>
        </div>
        <a class="btn btn-primary" href="{{ url('movimiento/suspensiones/crear') }}">＋ Nueva Suspensión</a>
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
                        <option value="{{ $t['value'] }}" @selected($idTrabajador == $t['value'])>{{ $t['label'] }}</option>
                    @endforeach
                </select>
            </div>
            <div class="form-group">
                <label class="form-label">Grupo</label>
                <select name="id_grupo" class="form-control">
                    <option value="0">Todos</option>
                    @foreach ($grupos as $g)
                        <option value="{{ $g['value'] }}" @selected($idGrupo == $g['value'])>{{ $g['label'] }}</option>
                    @endforeach
                </select>
            </div>
            <div class="form-group">
                <label class="form-label">Fecha Inicio</label>
                <input type="date" name="fecha_inicio" value="{{ $fechaInicio }}" class="form-control">
            </div>
            <div class="form-group">
                <label class="form-label">Fecha Fin</label>
                <input type="date" name="fecha_fin" value="{{ $fechaFin }}" class="form-control">
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
            <div class="empty-state"><div class="icon">📭</div><p>Ejecute una búsqueda para ver suspensiones.</p></div>
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
                                        @php $rid = $row['IdSuspension'] ?? $row['idsuspension'] ?? 0; @endphp
                                        <a class="btn btn-secondary btn-sm" href="{{ url("movimiento/suspensiones/{$rid}/editar") }}">✏️</a>
                                        <button class="btn btn-danger btn-sm" onclick="confirmDelete('{{ url("movimiento/suspensiones/{$rid}") }}')">🗑️</button>
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
