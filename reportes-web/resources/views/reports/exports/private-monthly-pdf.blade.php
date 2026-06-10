<!doctype html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <style>
        @page { margin: 14px; }
        body { font-family: Arial, Helvetica, sans-serif; font-size: 7px; color: #000; }
        h1 { font-size: 16px; margin: 0 0 4px; }
        table { border-collapse: collapse; width: 100%; table-layout: fixed; }
        th, td { border: 1px solid #000; padding: 2px; text-align: center; vertical-align: middle; }
        .green { background: #70ad47; font-weight: bold; }
        .yellow { background: #ffe699; font-weight: bold; }
        .blue { background: #d9eaf7; font-weight: bold; }
        .weekend { background: #d7dde5; }
        .name { text-align: left; width: 120px; }
        .base { width: 42px; font-style: italic; font-weight: bold; }
        .day { width: 13px; }
        .total { width: 42px; font-weight: bold; font-style: italic; }
        .obs { text-align: left; width: 160px; }
        .late { color: #f00; font-weight: bold; }
        .legend { margin-top: 18px; width: 260px; }
        .legend th { background: #bfbfbf; font-size: 11px; }
    </style>
</head>
<body>
    <h1>DATOS PERSONAL PRIVADO</h1>

    <table>
        <thead>
            <tr>
                <th class="green" rowspan="3" style="width:22px;">N</th>
                <th class="green" rowspan="3">APELLIDOS Y NOMBRES</th>
                <th class="blue" rowspan="3">INGRESO BASE<br>CONTRATO</th>
                <th class="yellow" colspan="{{ count($payload['days']) }}">{{ mb_strtoupper(\Carbon\Carbon::parse($payload['start'])->translatedFormat('F')) }}</th>
                <th class="green" rowspan="3">TOTAL<br>MINUTOS<br>ACUMULADO</th>
                <th class="yellow" rowspan="3">OBSERVACIONES</th>
            </tr>
            <tr>
                @foreach ($payload['days'] as $day)
                    <th class="day {{ $day['weekend'] ? 'weekend' : '' }}">{{ $day['weekday'] }}</th>
                @endforeach
            </tr>
            <tr>
                @foreach ($payload['days'] as $day)
                    <th class="day {{ $day['weekend'] ? 'weekend' : '' }}">{{ $day['day'] }}</th>
                @endforeach
            </tr>
        </thead>
        <tbody>
            @foreach ($payload['rows'] as $row)
                <tr>
                    <td>{{ $row['N'] }}</td>
                    <td class="name">{{ $row['Apellidos y nombres'] }}</td>
                    <td class="base">{{ $row['Ingreso base contrato'] }}</td>
                    @foreach ($payload['days'] as $day)
                        @php $value = $row[$day['weekday'].' '.$day['day']] ?? ''; @endphp
                        <td class="{{ $day['weekend'] ? 'weekend' : '' }} {{ is_numeric($value) && (int) $value > 0 ? 'late' : '' }}">{{ $value }}</td>
                    @endforeach
                    <td class="total">{{ $row['Total minutos acumulado'] }}</td>
                    <td class="obs">{{ $row['Observaciones'] }}</td>
                </tr>
            @endforeach
        </tbody>
    </table>

    <table class="legend">
        <tr><th colspan="2">LEYENDA</th></tr>
        <tr><td class="obs">ASISTENCIA NORMAL</td><td>&bull;</td></tr>
        <tr><td class="obs">DESCANSO MEDICO</td><td>DM</td></tr>
        <tr><td class="obs">PERMISOS</td><td>P</td></tr>
        <tr><td class="obs">TARDANZAS EN MINUTOS</td><td>1, 2, 3 ...</td></tr>
        <tr><td class="obs">LICENCIA PATERNIDAD LP</td><td>LP</td></tr>
        <tr><td class="obs">VACACIONES</td><td>V</td></tr>
        <tr><td class="obs">SUSPENSION SIN GOCE DE HABER</td><td>SSGH</td></tr>
        <tr><td class="obs">FALTAS</td><td>F</td></tr>
    </table>
</body>
</html>
