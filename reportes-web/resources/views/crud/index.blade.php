@extends('layouts.app', ['title' => $config['title']])

@section('topbar-title', $config['title'])
@section('breadcrumb', 'Mantenimiento / ' . $config['title'])

@section('content')
    <div class="page-header">
        <div>
            <h1>{{ $config['icon'] ?? '📋' }} {{ $config['title'] }}</h1>
            <p class="subtitle">Gestión completa de {{ strtolower($config['title']) }}</p>
        </div>
        <div>
            @if (isset($config['sp_agregar']))
                <a class="btn btn-primary" href="{{ url("mantenimiento/{$entity}/crear") }}">
                    ＋ Nuevo {{ $config['singular'] }}
                </a>
            @endif
        </div>
    </div>

    <div class="card table-card">
        <div class="table-toolbar">
            <div class="table-search">
                <input type="text" placeholder="Buscar en {{ strtolower($config['title']) }}..." id="searchInput">
            </div>
            <span class="table-count">{{ count($rows) }} registros</span>
        </div>

        @if (empty($rows))
            <div class="empty-state">
                <div class="icon">📭</div>
                <p>No se encontraron registros de {{ strtolower($config['title']) }}.</p>
            </div>
        @else
            <div style="overflow-x:auto">
                <table class="data-table">
                    <thead>
                        <tr>
                            <th style="width:50px">#</th>
                            @foreach ($config['columns'] as $col => $label)
                                <th>{{ $label }}</th>
                            @endforeach
                            <th style="width:120px">Acciones</th>
                        </tr>
                    </thead>
                    <tbody>
                        @foreach ($rows as $i => $row)
                            @php
                                $idCol = $config['id_column'];
                                $recordId = $row[$idCol] ?? $row[strtolower($idCol)] ?? $row[strtoupper($idCol)] ?? 0;
                            @endphp
                            <tr>
                                <td style="color:var(--muted);font-size:12px">{{ $i + 1 }}</td>
                                @foreach ($config['columns'] as $col => $label)
                                    @php
                                        $val = $row[$col] ?? $row[strtolower($col)] ?? $row[strtoupper($col)] ?? '';
                                    @endphp
                                    <td>
                                        @if ($col === 'Estado')
                                            @if ($val === 'N' || $val === 'Normal')
                                                <span class="badge badge-success">Normal</span>
                                            @elseif ($val === 'S' || $val === 'Suspendido')
                                                <span class="badge badge-danger">Suspendido</span>
                                            @else
                                                {{ $val }}
                                            @endif
                                        @elseif ($col === 'Remunerado')
                                            @if ($val)
                                                <span class="badge badge-success">Sí</span>
                                            @else
                                                <span class="badge badge-warn">No</span>
                                            @endif
                                        @elseif (str_contains(strtolower($col), 'fecha') && $val)
                                            @php
                                                try { $val = \Carbon\Carbon::parse($val)->format('d/m/Y'); } catch (\Throwable) {}
                                            @endphp
                                            {{ $val }}
                                        @else
                                            {{ $val }}
                                        @endif
                                    </td>
                                @endforeach
                                <td>
                                    <div class="table-actions">
                                        @if (isset($config['sp_modificar']))
                                            <a class="btn btn-secondary btn-sm" href="{{ url("mantenimiento/{$entity}/{$recordId}/editar") }}" title="Editar">✏️</a>
                                        @endif
                                        @if (isset($config['sp_eliminar']))
                                            <button class="btn btn-danger btn-sm" onclick="confirmDelete('{{ url("mantenimiento/{$entity}/{$recordId}") }}')" title="Eliminar">🗑️</button>
                                        @endif
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
