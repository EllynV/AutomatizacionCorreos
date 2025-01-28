<a href="index.php">Volver a cargar</a>

<?php
require 'vendor/autoload.php';
require 'diasLaborales.php';

use Google\Client;
use Google\Service\Gmail as Google_Service_Gmail;
use Google\Service\Sheets\ValueRange;
use Google\Service\Sheets as Google_Service_Sheets;


function extractEmail($fromHeader)
{
    // Usar expresión regular para encontrar la dirección de correo
    preg_match('/<([^>]+)>/', $fromHeader, $matches);
    return isset($matches[1]) ? $matches[1] : $fromHeader; // Retorna solo el correo o la cadena original si no hay coincidencias
}


function getInterestedLabel($emailLabels, $labelName)
{
    $response = "";
    foreach ($emailLabels as $label) {
        if (str_starts_with($label, $labelName . "/")) {
            // Prioridad: Retorna la parte después de $labelName . "/"
            if( $response == "")    
                $response = substr($label, strlen($labelName . "/"));
            else
                $response = $response . ', ' . substr($label, strlen($labelName . "/"));
            
                
           // $response = substr($label, strlen($labelName . "/"));
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


// Función para decodificar base64url
function base64_url_decode($input)
{
    // Reemplazar los caracteres específicos para base64url
    $base64 = strtr($input, '-_', '+/');
    // Asegurarse de que el tamaño sea múltiplo de 4
    $padding = strlen($base64) % 4;
    if ($padding) {
        $base64 .= str_repeat('=', 4 - $padding);
    }
    return base64_decode($base64);
}

function extractForwardedEmail($parts)
{
    foreach ($parts as $part) {
        // Si la parte tiene un cuerpo, analizarlo
        if (isset($part['body']) && isset($part['body']['data'])) {
            $body = base64_url_decode($part['body']['data']);

            // Buscar el correo original en un mensaje reenviado
            $pattern = '/(De:|From:)\s+([^\r\n<]+)\s+<([^>]+)>/i';
            if (preg_match($pattern, $body, $matches)) {
                return trim($matches[3]); // Devuelve el correo del remitente original
            }
        }

        // Si hay subpartes, analizarlas recursivamente
        if (isset($part['parts'])) {
            $email = extractForwardedEmail($part['parts']);
            if ($email) {
                return $email; // Devuelve el correo encontrado en subpartes
            }
        }
    }
    return null; // Si no se encontró ningún correo
}


function listMessagesGroupedByThread($service)
{
    $messages = [];
    $userLabels = [];
    $emailLabels = [];

    $date = '2025/01/28'; //YYYY/MM/DD

    // Ajustar la consulta para incluir correos del día completo
    $query = sprintf('label:read label:inbox after:%s before:%s', $date, date('Y/m/d', strtotime($date . ' +1 day')));
    $threads = $service->users_threads->listUsersThreads('me', ['q' =>  $query, 'maxResults' => 400])->getThreads();

    $labels = $service->users_labels->listUsersLabels('me');

    foreach ($labels->getLabels() as $label) {
        $userLabels[$label->getId()] = $label->getName();
    }



    if (empty($threads)) {
        echo "No se encontraron mensajes.";
        return;
    }

    foreach ($threads as $thread) {
        $threadId = $thread->getId();
        $threadDetails = $service->users_threads->get('me', $threadId);



        $threadMessages = [];
        foreach ($threadDetails->getMessages() as $message) {
            $headers = $message->getPayload()->getHeaders();
            $snippet = $message->getSnippet();


            // Obtener etiquetas del mensaje actual
            $labelIds = $message->getLabelIds();
            $emailLabels = []; // Reinicia las etiquetas para este mensaje

            if (!empty($labelIds)) {
                foreach ($labelIds as $labelId) {
                    if (isset($userLabels[$labelId])) {
                        $emailLabels[] = $userLabels[$labelId];
                    }
                }
            }

            $messageInfo = [
                'id' => $message->getId(),
                'snippet' => $snippet,
                'labels' => $emailLabels,
                'headers' => [],
                'messageUrl' => "https://mail.google.com/mail/u/0/#inbox/" . $message->getId(),
                'encuesta' => 'No'
            ];

            foreach ($headers as $header) {
                if (in_array($header->getName(), ['Subject', 'From', 'Date'])) {
                    $messageInfo['headers'][$header->getName()] = $header->getValue();
                }
            }

            $dateTime = new DateTime($messageInfo['headers']['Date']); // Analizar la fecha y hora

            $messageInfo['headers']['receivedDate'] = $dateTime->format('Y-m-d'); // Formatear la fecha
            $messageInfo['headers']['receivedTime']  = $dateTime->format('H:i:s'); // Formatear la hora

            $messageInfo['headers']['From'] = extractEmail($messageInfo['headers']['From']);
            $canalesIngresoArray = ['contacto@iudigital.edu.co', 'soportecampus@iudigital.edu.co', 'atencionalciudadano@iudigital.edu.co'];

            $messageInfo['headers']['canalIngreso']  = $canalesIngresoArray[0];



            foreach ($message->getPayload()->getParts() as $part) {
                if (isset($part['mimeType']) && ($part['mimeType'] == 'text/plain' || $part['mimeType'] == 'text/html')) {
                    // Si es texto o HTML, decodificar el contenido
                    $body = $part['body']['data'];
                    $bodyDecoded = base64_url_decode($body);

                    // Buscar si contiene la URL de Google Forms
                    if (strpos($bodyDecoded, 'https://docs.google.com/forms') !== false) {
                        // Si la URL se encuentra, hacer algo
                        $messageInfo['encuesta'] = 'Sí';
                        break;
                        // echo "Se encontró una URL de Google Forms en el mensaje: $bodyDecoded<br/><br/>";
                    }
                }
            }



            if ($messageInfo['headers']['From'] == 'atencionalciudadano@iudigital.edu.co') {
                echo $messageInfo['headers']['From'] . '---';

                // Función recursiva para analizar todas las partes del mensaje

                // Obtener las partes del mensaje y buscar el correo reenviado
                $parts = $message->getPayload()->getParts();
                $correoOriginal = extractForwardedEmail($parts);

                if ($correoOriginal) {
                    echo '**Correo Original: ' . $correoOriginal . '**<br>';
                    $messageInfo['headers']['From'] = $correoOriginal;
                } else {
                    echo 'No se encontró el correo original en el cuerpo del mensaje.<br>';
                }

                echo $messageInfo['headers']['From'] . '<br>';
            }


            $threadMessages[] = $messageInfo;
        }

        $messages[] = [
            'threadId' => $threadId,
            'messages' => $threadMessages,
        ];
    }

    return $messages;
}


function getSheetsService($client)
{
    $client->setApplicationName('Google Sheets API PHP');

    return new Google_Service_Sheets($client);
}




$client = new Google_Client();
$client->setAuthConfig('credentials.json');
$client->addScope(Google_Service_Gmail::GMAIL_READONLY);
$client->addScope('https://www.googleapis.com/auth/spreadsheets');

$spreadsheetId = '1xuQVC9zlxLRJriWlPG3Y7py_O_0gfO5RMVF1U0XdTZc'; //Contacto
$client->setAccessType('offline');
$client->setPrompt('select_account consent');
$myEmail = 'contacto@iudigital.edu.co';
$canalesIngresoArray = ['contacto@iudigital.edu.co', 'soportecampus@iudigital.edu.co', 'atencionalciudadano@iudigital.edu.co'];
$sheetsService = getSheetsService($client);


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


$service = new Google_Service_Gmail($client);
$threads = listMessagesGroupedByThread($service);

// Mostrar los mensajes
if ($threads) {
    foreach ($threads as $thread) {

        $encuesta = 'No';

        echo "<h3>Hilo ID: {$thread['threadId']}</h3>";

        foreach ($thread['messages'] as $message) {
            if ($message['encuesta'] == 'Sí') {
                $encuesta = 'Sí';
                break;
            }
        }

        $currentTime = new DateTime();
        $currentTime->setTimezone(new DateTimeZone('America/Bogota'));
        $formattedDate = $currentTime->format('Y-m-d H:i:s');


        $lastSentMessage = null;

        foreach (array_reverse($thread['messages']) as $message) {
            $headers = $message['headers'];

            if (isset($headers['From']) && strpos($headers['From'], $myEmail) !== false) {
                $lastSentMessage = $message;
                break; // Sal del bucle en cuanto encuentres el último mensaje enviado por ti
            }
        }


        // Obtener la fecha y hora del último mensaje
        $lastDateTime = new DateTime($lastSentMessage['headers']['Date']); // Crear DateTime desde la cadena de fecha
        $lastDateTime->setTimezone(new DateTimeZone('America/Bogota')); // Ajustar a tu zona horaria

        $FechaTotalFin = $lastDateTime->format('Y-m-d H:i:s'); // Obtener solo la hora
        $fechaInicio = new DateTime($thread['messages'][0]['headers']['Date']);  // Convertir string a DateTime
        $fechaFin = new DateTime($FechaTotalFin); // Convertir string a DateTime

        $fechaSolucion = $lastDateTime->format('Y-m-d'); // Formatear la fecha
        $horaSolucion  = $lastDateTime->format('H:i:s'); // Formatear la hora

        // Calcular la diferencia entre las fechas
        $intervalo = $fechaInicio->diff($fechaFin);
        $tiempoSolucion = calcularTiempoRespuesta($FechaTotalFin, $FechaTotalFin);


        $formattedDateInicio = $fechaInicio->format('Y-m-d H:i:s');
        echo $thread['messages'][0]['headers']['From'];
        echo '<pre>';
        print_r($thread['messages'][0]['labels']);
        echo '</pre>';
        // Crear fila
        $row = [
            $formattedDate ?? "", //Current datetime
            $thread['messages'][0]['headers']['From'] ?? "", //Correo electronico
            $thread['messages'][0]['headers']['receivedDate'] ?? "", //Fecha de recibido
            $thread['messages'][0]['headers']['receivedTime'] ?? "", //Hora de recibido
            getInterestedLabel($thread['messages'][0]['labels'], '*Agente') ?? "", //AGENTE //etiqueta
            getInterestedLabel($thread['messages'][0]['labels'], '*Prioridades') ?? "", //Prioridad //etiqueta
            getInterestedLabel($thread['messages'][0]['labels'], '*Nivel') ?? "", //Nivel //etiqueta
            $Nivelacademico  ?? "", //Nivel academico
            $ProgramaAcademico ?? "", //Programa academico
            $thread['messages'][0]['headers']['Subject']  ?? "", //ASUNTO
            getInterestedLabel($thread['messages'][0]['labels'], '*Creación de Usuarios') ?? "", //CREACION DE USUARIOS //etiqueta
            getInterestedLabel($thread['messages'][0]['labels'], '*No Pertinente') ?? "", //NO PERTINENTE //etiqueta
            getInterestedLabel($thread['messages'][0]['labels'], '*Canvas') ?? "", //CANVAS //etiqueta
            getInterestedLabel($thread['messages'][0]['labels'], '*GOOGLE') ?? "", //GOOGLE //etiqueta
            getInterestedLabel($thread['messages'][0]['labels'], '*Extensión') ?? "", //EXTENSION//etiqueta
            getInterestedLabel($thread['messages'][0]['labels'], '*EducaTIC') ?? "", //EDUCATIC//etiqueta
            getInterestedLabel($thread['messages'][0]['labels'], '*Certificados') ?? "", //CERTIFICADOS//etiqueta
            $proveedores ?? "", //PROVEEDORES//etiqueta
            getInterestedLabel($thread['messages'][0]['labels'], '*Plataformas')  ?? "", //PLATAFORMAS//etiqueta
            getInterestedLabel($thread['messages'][0]['labels'], '*Atención al Ciudadano') ?? "", //ATENCION AL CIUDADANO
            $fechaSolucion ?? "", //Fecha de solucion
            $horaSolucion ?? "", //Hora de solucion
            $thread['messages'][0]['headers']['canalIngreso']  ?? "", //Canal de Ingreso
            $encuesta ?? "", //Se envio Encuesta
            $thread['messages'][0]['headers']['receivedTime'] ?? "", //Hora de recibido
            $formattedDateInicio  ?? "", //Fecha total de inicio
            $FechaTotalFin ?? "", //Fecha total de fin
            $tiempoSolucion ?? "", //Fecha total de fin
            $message['messageUrl']
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
}
