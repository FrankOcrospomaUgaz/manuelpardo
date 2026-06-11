@extends('layouts.app', ['title' => 'Consolidado Mensual'])
@section('topbar-title', 'Consolidado Mensual')
@section('breadcrumb', 'Movimientos / Consolidado')

@section('content')
    <div class="page-header">
        <div>
            <h1>📊 Consolidado Mensual de Asistencia</h1>
            <p class="subtitle">Resumen consolidado mensual de horas trabajadas, faltas y tardanzas</p>
        </div>
    </div>

    {{-- Filtros --}}
    <div class="card" style="padding:18px;margin-bottom:20px;border-top:3px solid var(--accent)">
        <form method="GET" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:14px;align-items:end">
            <input type="hidden" name="buscar" value="1">
            <div class="form-group">
                <label class="form-label">Mes</label>
                <select name="mes" class="form-control">
                    @for ($m = 1; $m <= 12; $m++)
                        <option value="{{ $m }}" @selected($mes == $m)>
                            {{ Str::title(\Carbon\Carbon::create(2000, $m, 1)->translatedFormat('F')) }}
                        </option>
                    @endfor
                </select>
            </div>
            <div class="form-group">
                <label class="form-label">Año</label>
                <select name="anio" class="form-control">
                    @for ($y = date('Y') - 5; $y <= date('Y') + 2; $y++)
                        <option value="{{ $y }}" @selected($anio == $y)>{{ $y }}</option>
                    @endfor
                </select>
            </div>
            <div><button type="submit" class="btn btn-primary" style="width:100%">📊 Procesar / Visualizar</button></div>
        </form>
    </div>

    {{-- Resultados --}}
    <div class="card table-card">
        <div class="table-toolbar">
            <div class="table-search"><input type="text" placeholder="Filtrar trabajadores..."></div>
            <span class="table-count">{{ count($rows) }} trabajadores consolidados</span>
        </div>
        @if (empty($rows))
            <div class="empty-state"><div class="icon">📊</div><p>Seleccione mes y año y presione Procesar para visualizar consolidado.</p></div>
        @else
            <div style="overflow-x:auto">
                <table class="data-table">
                    <thead><tr>
                        <th>#</th>
                        @foreach (array_keys($rows[0]) as $col)
                            <th>{{ $col }}</th>
                        @endforeach
                    </tr></thead>
                    <tbody>
                        @foreach ($rows as $i => $row)
                            <tr>
                                <td style="color:var(--muted);font-size:12px">{{ $i + 1 }}</td>
                                @foreach ($row as $val)
                                    <td>{{ is_scalar($val) ? $val : '' }}</td>
                                @endforeach
                            </tr>
                        @endforeach
                    </tbody>
                </table>
            </div>
        @endif
    </div>
@endsection
