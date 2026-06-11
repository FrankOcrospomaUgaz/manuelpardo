@extends('layouts.app', ['title' => ($mode === 'create' ? 'Nuevo' : 'Editar') . ' ' . $config['singular']])

@section('topbar-title', ($mode === 'create' ? 'Nuevo' : 'Editar') . ' ' . $config['singular'])
@section('breadcrumb', 'Mantenimiento / ' . $config['title'] . ' / ' . ($mode === 'create' ? 'Nuevo' : 'Editar'))

@section('content')
    <div class="page-header">
        <div>
            <h1>{{ $config['icon'] ?? '📋' }} {{ $mode === 'create' ? 'Nuevo' : 'Editar' }} {{ $config['singular'] }}</h1>
            <p class="subtitle">{{ $mode === 'create' ? 'Complete los datos para registrar' : 'Modifique los datos del registro' }}</p>
        </div>
        <a class="btn btn-secondary" href="{{ url("mantenimiento/{$entity}") }}">← Volver al listado</a>
    </div>

    <div class="card form-card">
        <form method="POST"
              action="{{ $mode === 'create' ? url("mantenimiento/{$entity}") : url("mantenimiento/{$entity}/{$id}") }}">
            @csrf
            @if ($mode === 'edit')
                @method('PUT')
            @endif

            <div class="form-grid">
                @foreach ($config['fields'] as $name => $field)
                    @php
                        // Skip fields not applicable to current mode
                        if (($field['exclude_on'] ?? null) === $mode) continue;
                        if (($field['only_on'] ?? null) && ($field['only_on'] !== $mode)) continue;
                        if ($field['hidden'] ?? false) continue;

                        $type = $field['type'] ?? 'text';
                        $required = $field['required'] ?? false;

                        // Get current value
                        $value = old($name);
                        if ($value === null && $record) {
                            $value = $record[$name] ?? $record[strtolower($name)] ?? $record[strtoupper($name)] ?? '';
                        }
                        if ($value === null) {
                            $value = $field['default'] ?? '';
                        }

                        // Handle dates
                        if ($type === 'date' && $value && !preg_match('/^\d{4}-\d{2}-\d{2}$/', (string)$value)) {
                            try { $value = \Carbon\Carbon::parse($value)->format('Y-m-d'); } catch (\Throwable) {}
                        }
                    @endphp

                    <div class="form-group">
                        <label class="form-label">
                            {{ $field['label'] }}
                            @if ($required) <span class="required">*</span> @endif
                        </label>

                        @if ($type === 'select' && isset($field['lookup']))
                            {{-- Dynamic lookup select --}}
                            <select name="{{ $name }}" class="form-control" @if($required) required @endif>
                                <option value="0">— Seleccione —</option>
                                @foreach ($lookups[$field['lookup']] ?? [] as $option)
                                    <option value="{{ $option['value'] }}" @selected((string)$value === (string)$option['value'])>
                                        {{ $option['label'] }}
                                    </option>
                                @endforeach
                            </select>

                        @elseif ($type === 'select' && isset($field['options']))
                            {{-- Static options select --}}
                            <select name="{{ $name }}" class="form-control" @if($required) required @endif>
                                @foreach ($field['options'] as $optVal => $optLabel)
                                    <option value="{{ $optVal }}" @selected((string)$value === (string)$optVal)>
                                        {{ $optLabel }}
                                    </option>
                                @endforeach
                            </select>

                        @elseif ($type === 'boolean')
                            <select name="{{ $name }}" class="form-control">
                                <option value="1" @selected((int)$value === 1)>Sí</option>
                                <option value="0" @selected((int)$value === 0)>No</option>
                            </select>

                        @elseif ($type === 'date')
                            <input type="date" name="{{ $name }}" value="{{ $value }}"
                                   class="form-control" @if($required) required @endif>

                        @elseif ($type === 'number' || $type === 'integer')
                            <input type="number" name="{{ $name }}" value="{{ $value }}"
                                   class="form-control" @if($required) required @endif>

                        @else
                            <input type="text" name="{{ $name }}" value="{{ $value }}"
                                   class="form-control"
                                   @if($required) required @endif
                                   @if(isset($field['max'])) maxlength="{{ $field['max'] }}" @endif
                                   placeholder="{{ $field['label'] }}">
                        @endif

                        @if (isset($field['hint']))
                            <span class="form-hint">{{ $field['hint'] }}</span>
                        @endif
                    </div>
                @endforeach
            </div>

            <div class="form-actions">
                <button type="submit" class="btn btn-primary">
                    {{ $mode === 'create' ? '＋ Crear ' . $config['singular'] : '💾 Guardar Cambios' }}
                </button>
                <a class="btn btn-secondary" href="{{ url("mantenimiento/{$entity}") }}">Cancelar</a>
            </div>
        </form>
    </div>
@endsection
