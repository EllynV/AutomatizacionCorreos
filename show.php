<a href="index.php">Volver a cargar</a>
<?php
require 'vendor/autoload.php';

use Google\Client;
use Google\Service\Gmail;
use Google\Service\Sheets;
use Google\Service\Sheets as Google_Service_Sheets;
use Google\Service\Sheets\ValueRange;
// Iniciar sesión
session_start();

function getGmailService($client)
{

    $client->setApplicationName('Gmail API PHP Script');

    // Verificar si hay un código en la URL, intercambiarlo por un token
    if (isset($_GET['code'])) {
        $code = $_GET['code'];
        $accessToken = $client->fetchAccessTokenWithAuthCode($code);

        // Verificar si se obtuvo un token válido
        if (isset($accessToken['access_token'])) {
            $client->setAccessToken($accessToken);
        } else {
            echo json_encode($accessToken);
            exit;
        }
    } else {
        echo 'No se ha encontrado un código de autorización en la URL.';
        exit;
    }

    // Si el token ha caducado o no existe, redirigir a la autenticación
    if ($client->isAccessTokenExpired()) {
        if ($client->getRefreshToken()) {
            $client->fetchAccessTokenWithRefreshToken($client->getRefreshToken());
        } else {
            header('Location: index.php');
            exit;
        }
    }

    return new Gmail($client);
}

function getSheetsService($client)
{
    $client->setApplicationName('Google Sheets API PHP');

    return new Google_Service_Sheets($client);
}

function getInterestedLabel($emailLabels, $labelName)
{
    $response = "";
    foreach ($emailLabels as $label) {
        if (str_starts_with($label, $labelName . "/")) {
            // Prioridad: Retorna la parte después de $labelName . "/"
            $response = substr($label, strlen($labelName . "/"));
        }
    }

    // Si no se encontró con "/", verificamos si existe una coincidencia exacta
    foreach ($emailLabels as $label) {
        if ($label === $labelName) {
            $response = $label;
        }
    }

    return $response;
}

function extractEmail($fromHeader)
{
    // Usar expresión regular para encontrar la dirección de correo
    preg_match('/<([^>]+)>/', $fromHeader, $matches);
    return isset($matches[1]) ? $matches[1] : $fromHeader; // Retorna solo el correo o la cadena original si no hay coincidencias
}



// Configuración
$credentialsPath = 'credentials.json'; // Ruta al archivo JSON de credenciales
$spreadsheetId = '1xuQVC9zlxLRJriWlPG3Y7py_O_0gfO5RMVF1U0XdTZc'; //Contacto
$spreadsheetId = '13c8EjNZYI4Sq5wlQbvOTKVL4rLH_jXtsOxfrQrSAVMQ'; //ellyn.vasquez
$startDate = '2025/01/24'; //YYYY/MM/DD
$endDate = '2025/01/24'; //YYYY/MM/DD
$miEmail = 'ellyn.vasquez@iudigital.edu.co';
$canalesIngresoArray = ['contacto@iudigital.edu.co', 'soportecampus@iudigital.edu.co','atencionalciudadano@iudigital.edu.co'];
$client = new Google_Client();
$client->setAuthConfig($credentialsPath);
$client->setAccessType('offline');
$client->setPrompt('select_account consent');
$client->addScope('https://www.googleapis.com/auth/gmail.readonly');
$client->addScope('https://www.googleapis.com/auth/spreadsheets');


try {
    // Obtener el servicio de Gmail
    $gmailService = getGmailService($client);
    // Obtener el servicio de Sheets
    $sheetsService = getSheetsService($client);

    /* echo '<pre style="font-family: calibri; font-size: 11px">';
    print_r($client->getAccessToken());
    echo '</pre>'; */

    // Construir la consulta para mensajes leídos entre las fechas
    $query = sprintf('label:read after:%s before:%s', $startDate, $endDate);

    // Obtener los mensajes que coincidan con la consulta
    $messages = $gmailService->users_messages->listUsersMessages('me', ['maxResults' => 2]);


    // Obtener los mensajes
    //$messages = $gmailService->users_messages->listUsersMessages('me', ['maxResults' => 5]);

    $data = []; // Array para almacenar los datos extraídos

    $labels = $gmailService->users_labels->listUsersLabels('me');
    $userLabels = [];
    foreach ($labels->getLabels() as $label) {
        $userLabels[$label->getId()] = $label->getName();
    }


    foreach ($messages->getMessages() as $message) {
        $msg = $gmailService->users_messages->get('me', $message->getId());
        $headers = $msg->getPayload()->getHeaders();


        $messageId = $msg->getId(); // ID del mensaje
        $messageUrl = "https://mail.google.com/mail/u/0/#inbox/" . $messageId;


        $timestamp = $msg->getInternalDate();
        $from = $subject = $receivedDate = $receivedTime = "";

        $originalSender = ''; // Variable para almacenar el correo original

        foreach ($headers as $header) {
            // Obtener el remitente principal
            if ($header->getName() === 'From') {
                $from = $header->getValue();
            }
            if ($header->getName() == 'Subject') {
                $subject = $header->getValue();
            }

            // Intentar encontrar el correo original a través de "X-Forwarded-For"
            if ($header->getName() === 'X-Forwarded-For') {
                $originalSender = $header->getValue();
            }

            // O también verificar el campo "Received"
            if ($header->getName() === 'Received') {
                // Aquí puedes intentar extraer la dirección del remitente original si está presente
                if (preg_match('/from (.*?) /', $header->getValue(), $matches)) {
                    $originalSender = $matches[1]; // Usar la dirección que aparece en el encabezado "Received"
                }
            }
        }

        if(in_array($from, $canalesIngresoArray)){
            $canalIngreso = extractEmail($from);
        }

        $timestamp = $msg->getInternalDate(); // Marca de tiempo en milisegundos

        // Crear un objeto DateTime con la marca de tiempo
        $dateTime = new DateTime();
        $dateTime->setTimestamp($timestamp / 1000); // Convertir milisegundos a segundos

        // Establecer la zona horaria a la local (por ejemplo, Bogotá)
        $dateTime->setTimezone(new DateTimeZone('America/Bogota'));

        // Formatear la fecha y la hora
        $receivedDate = $dateTime->format('Y-m-d'); // Fecha en formato AAAA-MM-DD
        $receivedTime = $dateTime->format('H:i:s'); // Hora en formato HH:mm:ss

        /** INICIO:Ver el ultimo correo respondido */
        $threadId = $msg->getThreadId();
        $thread = $gmailService->users_threads->get('me', $threadId);

        foreach ($thread->getMessages() as $messageThread) {
            $fechaSolucion = '';
            $horaSolucion = '';
            $headersThread = $messageThread->getPayload()->getHeaders();
            $fromThread = '';
            $dateThread = '';

            // Extraer los encabezados relevantes
            foreach ($headersThread as $headerThread) {
                if ($headerThread->getName() === 'From') {
                    $fromThread = $headerThread->getValue();
                }
                if ($headerThread->getName() === 'Date') {
                    $dateThread = $headerThread->getValue();
                }
            }

            // Verifica si el mensaje fue enviado por ti
            if (strpos($fromThread, $miEmail) !== false) {
                // Ajustar la fecha y hora
                $dateTimeRe = new DateTime($dateThread);
                $dateTimeRe->setTimezone(new DateTimeZone('America/Bogota')); // Ajusta la zona horaria si es necesario
               
                $fechaSolucion = $dateTimeRe->format('Y-m-d');
                $horaSolucion = $dateTimeRe->format('H:i:s');

            }
        }

        /**FIN: */



        $labelIds = $msg->getLabelIds();

        // Mapear IDs de etiquetas con sus nombres
        //echo "<h1>Etiquetas del correo</h1>";
        $emailLabels = [];
        foreach ($labelIds as $labelId) {
            if (isset($userLabels[$labelId])) {
                $emailLabels[] = $userLabels[$labelId];
            }
        }

        // Crear fila
        $row = [
            $timestamp ?? "", //Marca temporal
            extractEmail($from) ?? "", //Correo electronico
            $receivedDate ?? "", //Fecha de recibido
            $receivedTime ?? "", //Hora de recibido
            getInterestedLabel($emailLabels, '*Agente') ?? "", //AGENTE //etiqueta
            getInterestedLabel($emailLabels, '*Prioridades') ?? "", //Prioridad //etiqueta
            getInterestedLabel($emailLabels, '*Nivel') ?? "", //Nivel //etiqueta
            $Nivelacademico  ?? "", //Nivel academico
            $ProgramaAcademico ?? "", //Programa academico
            $subject ?? "", //ASUNTO
            getInterestedLabel($emailLabels, '*Creación de Usuarios') ?? "", //CREACION DE USUARIOS //etiqueta
            getInterestedLabel($emailLabels, '*No Pertinente') ?? "", //NO PERTINENTE //etiqueta
            getInterestedLabel($emailLabels, '*Canvas') ?? "", //CANVAS //etiqueta
            getInterestedLabel($emailLabels, '*GOOGLE') ?? "", //GOOGLE //etiqueta
            getInterestedLabel($emailLabels, '*Extensión') ?? "", //EXTENSION//etiqueta
            getInterestedLabel($emailLabels, '*EducaTIC') ?? "", //EDUCATIC//etiqueta
            getInterestedLabel($emailLabels, '*Certificados') ?? "", //CERTIFICADOS//etiqueta
            $proveedores ?? "", //PROVEEDORES//etiqueta
            getInterestedLabel($emailLabels, '*Plataformas')  ?? "", //PLATAFORMAS//etiqueta
            getInterestedLabel($emailLabels, '*Atención al Ciudadano') ?? "", //ATENCION AL CIUDADANO
            $fechaSolucion ?? "", //Fecha de solucion
            $horaSolucion ?? "", //Hora de solucion
            $canalIngreso ?? "", //Canal de Ingreso
            $encuesta ?? "", //Se envio Encuesta
            $HoraRecibido ?? "", //Hora de recibido
            $FechaTotalInicio ?? "", //Fecha total de inicio
            $FechaTotalFin ?? "", //Fecha total de fin
            $tiempoSolucion ?? "", //Fecha total de fin
            $messageUrl
        ];

        // Validar si la fila no está vacía
        if (array_filter($row)) {
            $data['values'][] = $row;
        }
    }

    // Agregar encabezados (opcional)
    $headersTable = [
        'Marca temporal',
        'Correo electronico/Usuario Final',
        'Fecha de recibido',
        'Hora de recibido',
        'AGENTE',
        'Prioridad',
        'Nivel',
        'Nivel academico',
        'Programa academico',
        'ASUNTO',
        'CREACION DE USUARIOS',
        'NO PERTINENTE',
        'CANVAS',
        'GOOGLE',
        'EXTENSION',
        'EDUCATIC',
        'CERTIFICADOS',
        'PROVEEDORES',
        'PLATAFORMAS',
        'ATENCION AL CIUDADANO',
        'Fecha de solucion',
        'Hora de solucion',
        'Canal de Ingreso',
        'Se envio Encuesta',
        'Hora de recibido',
        'Fecha total de inicio',
        'Fecha total de fin',
        'Total tiempo de Solucion',
        'Url del correo'
    ];
    $data['values'] = array_merge([$headersTable], $data['values'] ?? []);


    // Crear el rango en Sheets
    $range = 'Hoja 1!A1';
    $body = new ValueRange(['values' => $data['values']]);

    $params = ['valueInputOption' => 'RAW'];
    $sheetsService->spreadsheets_values->update($spreadsheetId, $range, $body, $params);


    // Imprimir los datos
    /* echo '<pre style="font-family: calibri; font-size: 11px">';
    print_r($data);
    echo '</pre>'; */

    // Crear el rango en el que insertar los datos
    $range = 'Hoja 1!A2';

    $body = new ValueRange([
        'values' => $data
    ]);

    // Escribir los datos en la hoja de cálculo
    $params = ['valueInputOption' => 'RAW'];
    $sheetsService->spreadsheets_values->update($spreadsheetId, $range, $body, $params);

    echo "Datos insertados correctamente en Google Sheets.";
} catch (Exception $e) {
    echo '<pre style="font-family: calibri; font-size: 11px">';
    print_r(json_decode($e->getMessage(), true));
    echo '</pre>';
}
