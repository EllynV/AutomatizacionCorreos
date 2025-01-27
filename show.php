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
//$spreadsheetId = '13c8EjNZYI4Sq5wlQbvOTKVL4rLH_jXtsOxfrQrSAVMQ'; //ellyn.vasquez
$startDate = '2025/01/24'; //YYYY/MM/DD
$endDate = '2025/01/25'; //YYYY/MM/DD
$miEmail = 'ellyn.vasquez@iudigital.edu.co';
$canalesIngresoArray = ['contacto@iudigital.edu.co', 'soportecampus@iudigital.edu.co', 'atencionalciudadano@iudigital.edu.co'];
$client = new Google_Client();
$client->setAuthConfig($credentialsPath);
$client->setAccessType('offline');
$client->setPrompt('select_account consent');
$client->addScope('https://www.googleapis.com/auth/gmail.readonly');
$client->addScope('https://www.googleapis.com/auth/spreadsheets');
$data = []; // Array para almacenar los datos extraídos
try {
    $gmailService = getGmailService($client);
    $sheetsService = getSheetsService($client);
    $processedMessages = [];
    // Construir la consulta para mensajes leídos entre las fechas
    $query = sprintf('label:read');
    //$query = sprintf('label:read after:%s before:%s', $startDate, $endDate);

    // Obtener los mensajes que coincidan con la consulta
    $messages = $gmailService->users_messages->listUsersMessages('me', [
        'q' => $query,    // Aplicar la consulta
        'maxResults' => 200 // Número máximo de mensajes a obtener
    ]);

    $labels = $gmailService->users_labels->listUsersLabels('me');
    $userLabels = [];
    foreach ($labels->getLabels() as $label) {
        $userLabels[$label->getId()] = $label->getName();
    }

    foreach ($messages->getMessages() as $cont => $message) {
        echo "Procesando mensaje " . ($cont + 1) . " de " . count($messages->getMessages()) . ' <br>';
        $msg = $gmailService->users_messages->get('me', $message->getId());
        // Obtener el ID del mensaje
        //$messageId = $msg->getId();




        $threadId = $message->getThreadId();


        // Obtener todos los mensajes del hilo utilizando el threadId
        $thread = $gmailService->users_threads->get('me', $threadId);
        $firstMessage = null;
        $from = $subject = $receivedDate = $receivedTime = "";

        // Iterar sobre los mensajes del hilo para encontrar el primero
        foreach ($thread->getMessages() as $msgth) {
            // Obtener el timestamp (fecha interna) del mensaje
            $timestamp = $msgth->getInternalDate();

            // Agregar cada mensaje con su fecha al array
            $threadMessages[] = [
                'message' => $msgth,
                'timestamp' => $timestamp
            ];
        }
        // Ordenar los mensajes por fecha (internalDate) de manera ascendente
        usort($threadMessages, function ($a, $b) {
            return $a['timestamp'] <=> $b['timestamp']; // Ordenar de menor a mayor
        });

        $firstMessage = $threadMessages[0];
        // Ahora $firstMessage contiene el primer correo del hilo
        if ($firstMessage) {
            // Extraer los detalles del primer correo
            $msgth = $firstMessage['message'];
            $headers = $msgth->getPayload()->getHeaders();
            $messageId = $msgth->getId(); // ID del primer mensaje
            $messageUrl = "https://mail.google.com/mail/u/0/#inbox/" . $messageId;
            $specificUrl = 'https://docs.google.com/forms/d/e/1FAIpQLSc9Nl_BHQ9JdKF6W8LxJcGN9dk3WB7uD7aTtT6-2xFsV4wT7g/viewform';
            // Verificar si el mensaje ya ha sido procesado (comprobando el messageId)
            if (in_array($messageId, $processedMessages)) {
                continue; // Si ya procesamos este mensaje, saltamos al siguiente
            }

            $processedMessages[] = $messageId;
            echo '<pre>';
            print_r($processedMessages);
            echo '</pre>';

            // Obtener los encabezados del primer mensaje
            foreach ($headers as $header) {
               
                // Buscar el campo "From"
                if ($header->getName() == 'From') {
                    $from = extractEmail($header->getValue());  // Obtiene el valor del encabezado "From"
                }
                // Buscar el campo "Subject"
                if ($header->getName() == 'Subject') {
                    $subject = $header->getValue();
                }
            }

            $timestamp = $msgth->getInternalDate(); // Marca de tiempo en milisegundos
            // Crear un objeto DateTime con la marca de tiempo
            $dateTime = new DateTime();
            $dateTime->setTimestamp($timestamp / 1000); // Convertir milisegundos a segundos

            // Establecer la zona horaria a la local (por ejemplo, Bogotá)
            $dateTime->setTimezone(new DateTimeZone('America/Bogota'));

            // Formatear la fecha y la hora
            $receivedDate = $dateTime->format('Y-m-d'); // Fecha en formato AAAA-MM-DD
            $receivedTime = $dateTime->format('H:i:s'); // Hora en formato HH:mm:ss
            $FechaTotalInicio = $dateTime->format('Y-m-d H:i:s'); // Obtener solo la hora


            if (in_array($from, $canalesIngresoArray)) {
                $canalIngreso = $from;
            } else {
                $canalIngreso = $canalesIngresoArray[0];
            }
            echo $from;
            if ($from == 'atencionalciudadano@iudigital.edu.co') {
               echo $originalBody = $msgth->getPayload()->getBody()->getData();

               
            }
        }


        // El último mensaje será el último en el array ordenado
        $lastMessage = $threadMessages[count($threadMessages) - 1];

        // Obtener la fecha y hora del último mensaje
        $lastTimestamp = $lastMessage['timestamp'];
        $lastDateTime = new DateTime();
        $lastDateTime->setTimestamp($lastTimestamp / 1000); // Convertir milisegundos a segundos
        $lastDateTime->setTimezone(new DateTimeZone('America/Bogota')); // Ajustar a tu zona horaria
        $FechaTotalFin = $lastDateTime->format('Y-m-d H:i:s'); // Obtener solo la hora
        $fechaInicio = new DateTime($FechaTotalInicio);  // Convertir string a DateTime
        $fechaFin = new DateTime($FechaTotalFin);        // Convertir string a DateTime

        // Calcular la diferencia entre las fechas
        $intervalo = $fechaInicio->diff($fechaFin);
        $tiempoSolucion = $intervalo->format('%d días, %h:%i:%s');

        $labelIds = $msg->getLabelIds();

        // Mapear IDs de etiquetas con sus nombres
        //echo "<h1>Etiquetas del correo</h1>";
        $emailLabels = [];
        foreach ($labelIds as $labelId) {
            if (isset($userLabels[$labelId])) {
                $emailLabels[] = $userLabels[$labelId];
            }
        }

        $currentTime = new DateTime();
        $currentTime->setTimezone(new DateTimeZone('America/Bogota'));
        $formattedDate = $currentTime->format('Y-m-d H:i:s');
        // Crear fila
        $row = [
            $formattedDate ?? "", //Current datetime
            $from ?? "", //Correo electronico
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
            $receivedTime ?? "", //Hora de recibido
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


    echo "Datos insertados correctamente en Google Sheets.";
} catch (Exception $e) {
    echo '<pre style="font-family: calibri; font-size: 11px">';
    print_r(json_decode($e->getMessage(), true));
    echo '</pre>';
}
