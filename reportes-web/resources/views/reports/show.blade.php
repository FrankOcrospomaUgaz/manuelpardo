@extends('reports.layout', ['title' => $report['title']])

@section('content')
    <div class="page-head">
        <div>
            <h1>{{ $report['title'] }}</h1>
        </div>
        @if ($rows !== [])
            <div class="actions">
                <a class="button secondary" href="{{ route('reports.excel', array_merge(['slug' => $slug], $params)) }}"><span class="file-icon excel">X</span> Excel</a>
                <a class="button secondary" href="{{ route('reports.pdf', array_merge(['slug' => $slug], $params)) }}"><span class="file-icon pdf">P</span> PDF</a>
                <a class="button secondary" href="{{ route('reports.export', array_merge(['slug' => $slug], $params)) }}"><span class="file-icon csv">C</span> CSV</a>
            </div>
        @endif
    </div>

    @if ($error)
        <div class="alert">
            No se pudo ejecutar la consulta. Verifique la conexion SQL Server, el driver pdo_sqlsrv y que existan los procedimientos/tables legacy.
            <br>{{ $error }}
        </div>
    @endif

    <section class="panel filter-panel">
        <form class="filters" action="{{ route('reports.run', $slug) }}" method="post">
            @csrf
            @foreach (($report['params'] ?? []) as $name => $field)
                @php
                    $value = old($name, $params[$name] ?? ($field['default'] ?? ''));
                    $type = $field['type'] ?? 'text';
                @endphp

                @if ($type === 'select')
                    <label>
                        {{ $field['label'] }}
                        <select name="{{ $name }}">
                            @foreach (($field['options'] ?? []) as $optionValue => $optionLabel)
                                <option value="{{ $optionValue }}" @selected((string) $value === (string) $optionValue)>{{ $optionLabel }}</option>
                            @endforeach
                        </select>
                    </label>
                @elseif ($type === 'lookupselect')
                    @php
                        $options = $lookups[$field['lookup'] ?? ''] ?? [];
                    @endphp
                    <label>
                        {{ $field['label'] }}
                        <select name="{{ $name }}" @required($field['required'] ?? false)>
                            <option value="0">{{ ($field['required'] ?? false) ? 'Seleccione...' : 'Todos' }}</option>
                            @foreach ($options as $option)
                                <option value="{{ $option['value'] }}" @selected((string) $value === (string) $option['value'])>
                                    {{ $option['label'] }}
                                </option>
                            @endforeach
                        </select>
                        @if ($options === [])
                            <span class="field-hint">No se pudo cargar este catalogo desde SQL Server.</span>
                        @endif
                    </label>
                @elseif ($type === 'multiselect')
                    @php
                        $selected = collect(is_array($value) ? $value : explode(',', (string) $value))->map(fn ($item) => (string) trim($item))->filter()->all();
                        $options = $lookups[$field['lookup'] ?? ''] ?? [];
                    @endphp
                    <label>
                        {{ $field['label'] }}
                        <select name="{{ $name }}[]" multiple>
                            @forelse ($options as $option)
                                <option value="{{ $option['value'] }}" @selected(in_array((string) $option['value'], $selected, true))>
                                    {{ $option['label'] }}
                                </option>
                            @empty
                                <option disabled>No se pudieron cargar grupos</option>
                            @endforelse
                        </select>
                        @if (! empty($field['help']))
                            <span class="field-hint">{{ $field['help'] }}</span>
                        @endif
                    </label>
                @elseif ($type === 'checkbox')
                    <label>
                        {{ $field['label'] }}
                        <span class="checkbox">
                            <input type="hidden" name="{{ $name }}" value="0">
                            <input type="checkbox" name="{{ $name }}" value="1" @checked((bool) $value)>
                            Activo
                        </span>
                    </label>
                @else
                    <label>
                        {{ $field['label'] }}
                        <input
                            type="{{ $type }}"
                            name="{{ $name }}"
                            value="{{ $value }}"
                            @required($field['required'] ?? false)
                        >
                        @if (! empty($field['help']))
                            <span class="field-hint">{{ $field['help'] }}</span>
                        @endif
                    </label>
                @endif
            @endforeach

            <div class="actions">
                <button type="submit">Consultar</button>
            </div>
        </form>
    </section>

    <section>
        @if ($rows === [])
            <div class="empty-state">Ejecute una consulta para visualizar la informacion y habilitar las exportaciones.</div>
        @else
            <section class="metric-strip">
                <div class="metric">
                    <span>Registros encontrados</span>
                    <strong>{{ count($rows) }}</strong>
                </div>
                <div class="metric">
                    <span>Periodo</span>
                    <strong>{{ ($params['fecha_inicio'] ?? '-') }} / {{ ($params['fecha_fin'] ?? '-') }}</strong>
                </div>
                <div class="metric">
                    <span>Formato</span>
                    <strong>Excel / PDF</strong>
                </div>
            </section>
            <div class="table-wrap">
                <table>
                    <thead>
                        <tr>
                            @foreach (array_keys($rows[0]) as $column)
                                <th>{{ $column }}</th>
                            @endforeach
                        </tr>
                    </thead>
                    <tbody>
                        @foreach ($rows as $row)
                            <tr>
                                @foreach ($row as $value)
                                    <td>{{ is_scalar($value) || $value === null ? $value : json_encode($value) }}</td>
                                @endforeach
                            </tr>
                        @endforeach
                    </tbody>
                </table>
            </div>
        @endif
    </section>
@endsection
