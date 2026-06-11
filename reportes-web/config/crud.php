<?php

/*
|--------------------------------------------------------------------------
| CRUD Entity Definitions
|--------------------------------------------------------------------------
| Each key maps to a URL segment. Fields are ordered as the SP expects them.
| 'type' => text|integer|number|date|select|boolean
| 'lookup' => key for LegacyCrudService::getLookup()
| 'exclude_on' => 'create'|'update' to skip that field for that action
| 'only_on' => 'update' to include only on update
*/

return [
    'areas' => [
        'title' => 'Áreas',
        'singular' => 'Área',
        'icon' => '🏢',
        'id_column' => 'IdArea',
        'sp_buscar' => 'up_BuscarArea',
        'sp_agregar' => 'up_AgregarArea',
        'sp_modificar' => 'up_ModificarArea',
        'sp_eliminar' => 'up_EliminarArea',
        'buscar_default' => [0],
        'columns' => ['IdArea' => 'ID', 'Descripcion' => 'Descripción'],
        'fields' => [
            'Descripcion' => ['label' => 'Descripción', 'type' => 'text', 'required' => true, 'max' => 100],
        ],
    ],

    'cargos' => [
        'title' => 'Cargos',
        'singular' => 'Cargo',
        'icon' => '💼',
        'id_column' => 'IdCargo',
        'sp_buscar' => 'up_BuscarCargo',
        'sp_agregar' => 'up_AgregarCargo',
        'sp_modificar' => 'up_ModificarCargo',
        'sp_eliminar' => 'up_EliminarCargo',
        'buscar_default' => [0],
        'columns' => ['IdCargo' => 'ID', 'Descripcion' => 'Descripción'],
        'fields' => [
            'Descripcion' => ['label' => 'Descripción', 'type' => 'text', 'required' => true, 'max' => 100],
        ],
    ],

    'categorias' => [
        'title' => 'Categorías',
        'singular' => 'Categoría',
        'icon' => '📂',
        'id_column' => 'IdCategoria',
        'sp_buscar' => 'up_BuscarCategoria',
        'sp_agregar' => 'up_AgregarCategoria',
        'sp_modificar' => 'up_ModificarCategoria',
        'sp_eliminar' => 'up_EliminarCategoria',
        'buscar_default' => [0],
        'columns' => ['IdCategoria' => 'ID', 'Descripcion' => 'Descripción'],
        'fields' => [
            'Descripcion' => ['label' => 'Descripción', 'type' => 'text', 'required' => true, 'max' => 100],
        ],
    ],

    'tipos-hora' => [
        'title' => 'Tipos de Hora',
        'singular' => 'Tipo de Hora',
        'icon' => '🕐',
        'id_column' => 'IdTipoHora',
        'sp_buscar' => 'up_BuscarTipoHora',
        'sp_agregar' => 'up_AgregarTipoHora',
        'sp_modificar' => 'up_ModificarTipoHora',
        'sp_eliminar' => 'up_EliminarTipoHora',
        'buscar_default' => [0, ''],
        'columns' => ['IdTipoHora' => 'ID', 'Descripcion' => 'Descripción', 'Simbolo' => 'Símbolo'],
        'fields' => [
            'Descripcion' => ['label' => 'Descripción', 'type' => 'text', 'required' => true, 'max' => 100],
            'Simbolo' => ['label' => 'Símbolo', 'type' => 'text', 'required' => false, 'max' => 10],
        ],
    ],

    'motivos' => [
        'title' => 'Motivos de Justificación',
        'singular' => 'Motivo de Justificación',
        'icon' => '📝',
        'id_column' => 'IdMotivoJustificacion',
        'sp_buscar' => 'up_BuscarMotivo',
        'sp_agregar' => 'up_AgregarMotivoJustificacion',
        'sp_modificar' => 'up_ModificarMotivoJustificacion',
        'sp_eliminar' => 'up_EliminarMotivoJustificacion',
        'buscar_default' => [0],
        'columns' => ['IdMotivoJustificacion' => 'ID', 'Descripcion' => 'Descripción'],
        'fields' => [
            'Descripcion' => ['label' => 'Descripción', 'type' => 'text', 'required' => true, 'max' => 200],
        ],
    ],

    'feriados' => [
        'title' => 'Feriados',
        'singular' => 'Feriado',
        'icon' => '🎉',
        'id_column' => 'IdFeriado',
        'sp_buscar' => 'up_BuscarFeriado',
        'sp_agregar' => 'up_AgregarFeriado',
        'sp_modificar' => 'up_ModificarFeriado',
        'sp_eliminar' => 'up_EliminarFeriado',
        'buscar_default' => [0, '01/01/1900', '31/12/2099'],
        'columns' => ['IdFeriado' => 'ID', 'FechaFeriado' => 'Fecha', 'Descripcion' => 'Descripción', 'Remunerado' => 'Remunerado'],
        'fields' => [
            'FechaFeriado' => ['label' => 'Fecha del Feriado', 'type' => 'date', 'required' => true],
            'Descripcion' => ['label' => 'Descripción', 'type' => 'text', 'required' => true, 'max' => 200],
            'Remunerado' => ['label' => 'Remunerado', 'type' => 'boolean', 'default' => 0],
        ],
    ],

    'grupos' => [
        'title' => 'Grupos de Trabajador',
        'singular' => 'Grupo',
        'icon' => '👥',
        'id_column' => 'IdGrupo',
        'sp_buscar' => 'up_BuscarGrupo',
        'sp_agregar' => 'up_AgregarGrupo',
        'sp_modificar' => 'up_ModificarGrupo',
        'sp_eliminar' => 'up_EliminarGrupo',
        'buscar_default' => [0],
        'columns' => ['IdGrupo' => 'ID', 'Descripcion' => 'Descripción'],
        'fields' => [
            'Descripcion' => ['label' => 'Descripción', 'type' => 'text', 'required' => true, 'max' => 100],
        ],
    ],

    'trabajadores' => [
        'title' => 'Trabajadores',
        'singular' => 'Trabajador',
        'icon' => '👤',
        'id_column' => 'IdTrabajador',
        'sp_buscar' => 'up_BuscarTrabajador',
        'sp_agregar' => 'up_AgregarTrabajador',
        'sp_modificar' => 'up_ModificarTrabajador',
        'sp_eliminar' => 'up_EliminarTrabajador',
        'buscar_default' => [0, '', 0, 0, 0],
        'columns' => [
            'IdTrabajador' => 'ID',
            'Tarjeta' => 'Tarjeta',
            'Apellido' => 'Apellido',
            'Nombre' => 'Nombre',
            'Dni' => 'DNI',
            'Descripcion' => 'Categoría',
            'Estado' => 'Estado',
        ],
        'fields' => [
            'IdGrupo' => ['label' => 'Grupo', 'type' => 'select', 'lookup' => 'grupos', 'required' => true, 'sp_order' => 1],
            'IdCategoria' => ['label' => 'Categoría', 'type' => 'select', 'lookup' => 'categorias', 'required' => true, 'sp_order' => 2],
            'IdArea' => ['label' => 'Área', 'type' => 'select', 'lookup' => 'areas', 'required' => true, 'sp_order' => 3],
            'IdCargo' => ['label' => 'Cargo', 'type' => 'select', 'lookup' => 'cargos', 'required' => true, 'sp_order' => 4],
            'Tarjeta' => ['label' => 'Nro. Tarjeta', 'type' => 'text', 'max' => 20, 'sp_order' => 5],
            'Apellido' => ['label' => 'Apellido', 'type' => 'text', 'required' => true, 'max' => 100, 'sp_order' => 6],
            'Nombre' => ['label' => 'Nombre', 'type' => 'text', 'required' => true, 'max' => 100, 'sp_order' => 7],
            'Dni' => ['label' => 'DNI', 'type' => 'text', 'max' => 15, 'sp_order' => 8],
            'Telefono' => ['label' => 'Teléfono', 'type' => 'text', 'max' => 20, 'sp_order' => 9],
            'Direccion' => ['label' => 'Dirección', 'type' => 'text', 'max' => 200, 'sp_order' => 10],
            'FechaNacimiento' => ['label' => 'Fecha Nacimiento', 'type' => 'date', 'sp_order' => 11],
            'FechaIngreso' => ['label' => 'Fecha Ingreso', 'type' => 'date', 'sp_order' => 12],
            'Foto' => ['label' => 'Foto', 'type' => 'text', 'max' => 200, 'sp_order' => 13],
            'Huella' => ['label' => 'Huella', 'type' => 'text', 'max' => 500, 'sp_order' => 14, 'hidden' => true],
            'IndTarjeta' => ['label' => 'Ind. Tarjeta', 'type' => 'text', 'max' => 5, 'sp_order' => 15, 'hidden' => true],
            'Estado' => ['label' => 'Estado', 'type' => 'select', 'options' => ['N' => 'Normal', 'S' => 'Suspendido'], 'only_on' => 'update', 'default' => 'N'],
        ],
    ],

    'usuarios' => [
        'title' => 'Usuarios',
        'singular' => 'Usuario',
        'icon' => '🔑',
        'id_column' => 'IdUsuario',
        'sp_buscar' => 'SELECT IdUsuario, NombreUsuario as Nombre, Clave FROM Usuario WHERE 1=1',
        'sp_agregar' => 'up_AgregarUsuario',
        'sp_modificar' => 'up_ModificarUsuario',
        'sp_eliminar' => 'up_EliminarUsuario',
        'columns' => ['IdUsuario' => 'ID', 'Nombre' => 'Nombre'],
        'fields' => [
            'Nombre' => ['label' => 'Nombre de Usuario', 'type' => 'text', 'required' => true, 'max' => 50],
            'Clave' => ['label' => 'Contraseña', 'type' => 'text', 'required' => true, 'max' => 50],
        ],
    ],
];
