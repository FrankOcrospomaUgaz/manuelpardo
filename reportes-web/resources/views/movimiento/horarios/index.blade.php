@extends('layouts.app', ['title' => 'Horarios por Trabajador'])
@section('topbar-title', 'Horarios por Trabajador')
@section('breadcrumb', 'Movimientos / Horarios')

@section('content')
    <div class="page-header">
        <div>
            <h1>📅 Horarios por Trabajador</h1>
            <p class="subtitle">Asignación de horarios ideales para el control de asistencia</p>
        </div>
        <a class="btn btn-primary" href="{{ url('movimiento/horarios/crear') }}">＋ Asignar Horario</a>
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
                <label class="form-label">Grupo</label>
                <select name="id_grupo" class="form-control">
                    <option value="0">Todos</option>
                    @foreach ($grupos as $g)
                        <option value="{{ $g['value'] }}" @selected($filters['id_grupo'] == $g['value'])>{{ $g['label'] }}</option>
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
            <div class="table-search"><input type="text" placeholder="Filtrar horarios..."></div>
            <span class="table-count">{{ count($rows) }} horarios</span>
        </div>
        @if (empty($rows))
            <div class="empty-state"><div class="icon">📅</div><p>Ejecute una búsqueda para visualizar horarios.</p></div>
        @else
            <div style="overflow-x:auto">
                <table class="data-table">
                    <thead><tr>
                        <th>#</th>
                        @foreach (array_keys($rows[0]) as $col)
                            <th>{{ $col }}</th>
                        @endforeach
                        <th>Acción</th>
                    </tr></thead>
                    <tbody>
                        @foreach ($rows as $i => $row)
                            <tr>
                                <td style="color:var(--muted);font-size:12px">{{ $i + 1 }}</td>
                                @foreach ($row as $val)
                                    <td>{{ is_scalar($val) ? $val : '' }}</td>
                                @endforeach
                                <td>
                                    @php $hid = $row['IdHorario'] ?? $row['idhorario'] ?? 0; @endphp
                                    @if ($hid > 0)
                                        <button class="btn btn-danger btn-sm" onclick="confirmDelete('{{ url("movimiento/horarios/{$hid}") }}')">🗑️</button>
                                    @else
                                        —
                                    @endif
                                </td>
                            </tr>
                        @endforeach
                    </tbody>
                </table>
            </div>
        @endif
    </div>
@endsection
