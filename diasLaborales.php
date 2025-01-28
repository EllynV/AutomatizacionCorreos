<?php
function calcularTiempoRespuesta($fechaInicio, $fechaFin)
{
    // Obtener el arreglo de festivos
    $festivos = obtenerFestivos();

    // Convertir las fechas a objetos DateTime
    $inicio = new DateTime($fechaInicio);
    $fin = new DateTime($fechaFin);

    // Ajustar las horas para que no se cuente fuera del horario laboral
    ajustarHoraLaboral($inicio);
    ajustarHoraLaboral($fin);

    // Si la fecha de inicio o fin está fuera del horario laboral, no se cuenta
    if ($inicio > $fin) {
        return '0 días 0 horas 0 minutos 0 segundos'; // El correo no se recibió dentro del horario laboral
    }

    // Calcular el tiempo de respuesta en segundos
    $segundosTotales = 0;
    $interval = new DateInterval('P1D');
    $periodo = new DatePeriod($inicio, $interval, $fin->add($interval)); // Incluye fecha de fin

    foreach ($periodo as $fecha) {
        $diaSemana = $fecha->format('N'); // 1 = Lunes, 7 = Domingo
        if ($diaSemana < 6 && !in_array($fecha->format('Y-m-d'), $festivos)) {
            // Verificar si está dentro del horario laboral del día
            if ($fecha == $inicio && $fecha->format('H:i') >= '07:30') {
                $segundosTotales += (strtotime('17:30') - strtotime($fecha->format('H:i')));
            } elseif ($fecha == $fin && $fecha->format('H:i') <= '17:30') {
                $segundosTotales += (strtotime($fecha->format('Y-m-d') . ' 17:30') - strtotime($fecha->format('Y-m-d') . ' 00:00:00'));
            } elseif ($fecha != $inicio && $fecha != $fin) {
                $segundosTotales += (strtotime('17:30') - strtotime('07:30'));
            }
        }
    }

    // Convertir los segundos totales en días, horas, minutos y segundos
    $días = floor($segundosTotales / 86400);
    $segundosRestantes = $segundosTotales % 86400;
    $horas = floor($segundosRestantes / 3600);
    $segundosRestantes %= 3600;
    $minutos = floor($segundosRestantes / 60);
    $segundos = $segundosRestantes % 60;

    return "{$días} días {$horas} horas {$minutos} minutos {$segundos} segundos";
}

// Función para ajustar la hora de inicio y fin dentro del horario laboral (7:30 a 17:30)
function ajustarHoraLaboral(&$fecha)
{
    $horaInicio = '07:30';
    $horaFin = '17:30';

    // Si la hora es antes de las 7:30, ajusta a las 7:30
    if ($fecha->format('H:i') < $horaInicio) {
        $fecha->setTime(7, 30);
    }

    // Si la hora es después de las 17:30, ajusta a las 17:30
    if ($fecha->format('H:i') > $horaFin) {
        $fecha->setTime(17, 30);
    }
}

// Función para obtener los festivos en Colombia (día y mes)
function obtenerFestivos()
{
    return [
        // Fijos
        '01-01', // Año Nuevo
        '19-03', // Día de San José
        '01-05', // Día del Trabajo
        '20-07', // Día de la Independencia
        '07-08', // Batalla de Boyacá
        '12-10', // Día de la Raza
        '01-11', // Día de Todos los Santos
        '08-12', // Inmaculada Concepción
        '25-12', // Navidad

        // Movibles (según la ley de festivos)
        date('d-m', strtotime('third monday of january')), // Día de los Reyes Magos
        date('d-m', strtotime('third monday of febrero')), // Día de la Marmolera
        date('d-m', strtotime('monday after easter')), // Lunes de Pascua
        date('d-m', strtotime('Ascensión: monday after 40 days of Easter')),
        date('d-m', strtotime('Corpus Christi: sunday after 60 days of Easter')),
        date('d-m', strtotime('Sacred Heart of Jesus: friday after Corpus Christi')),
        date('d-m', strtotime('third monday of agosto')), // La Asunción
        date('d-m', strtotime('second monday of octubre')), // Día de la Raza
        date('d-m', strtotime('first monday of noviembre')), // Día de Todos los Santos
        date('d-m', strtotime('second monday of diciembre')), // Día de la Inmaculada Concepción
    ];
}
?>
