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


// Configuración
$credentialsPath = 'credentials.json'; // Ruta al archivo JSON de credenciales
$spreadsheetId = '13c8EjNZYI4Sq5wlQbvOTKVL4rLH_jXtsOxfrQrSAVMQ';

$client = new Google_Client();
$client->setAuthConfig($credentialsPath);
$client->setAccessType('offline');
$client->setPrompt('select_account consent');
$client->addScope('https://www.googleapis.com/auth/gmail.readonly');
$client->addScope('https://www.googleapis.com/auth/spreadsheets');


try {
    // Obtener el servicio de Gmail
    $gmailService = getGmailService($client, $credentialsPath);
    // Obtener el servicio de Sheets
    $sheetsService = getSheetsService($client, $credentialsPath);

    print_r($client->getAccessToken());
    // Obtener los mensajes
    $messages = $gmailService->users_messages->listUsersMessages('me', ['maxResults' => 2]);

    $data = []; // Array para almacenar los datos extraídos

    $labels = $gmailService->users_labels->listUsersLabels('me');
    foreach ($labels->getLabels() as $label) {
        $labels[$label->getId()] = $label->getName();
        echo "Etiqueta: " . $label->getName() . " (ID: " . $label->getId() . ")<br>";
    }


    echo "<h1>Correos electrónicos</h1>";
    foreach ($messages->getMessages() as $message) {
        $msg = $gmailService->users_messages->get('me', $message->getId());
        $headers = $msg->getPayload()->getHeaders();

        $timestamp = $msg->getInternalDate(); // Marca temporal
        $from = '';
        $subject = '';
        $receivedDate = '';
        $receivedTime = '';

        $labelIds = $msg->getLabelIds();


        // Mapear IDs de etiquetas con sus nombres
        echo "<h1>Etiquetas del correo</h1>";
        foreach ($labelIds as $labelId) {
            if (isset($labels[$labelId])) {
                echo "Etiqueta: " . $labels[$labelId] . "<br>";

                

            } else {
                echo "Etiqueta desconocida (ID: $labelId)<br>";
            }
        }


        // Procesar los encabezados para obtener los detalles
        foreach ($headers as $header) {
            if ($header->getName() == 'From') {
                $from = $header->getValue();
            }
            if ($header->getName() == 'Subject') {
                $subject = $header->getValue();
            }
            if ($header->getName() == 'Date') {
                $date = new DateTime($header->getValue());
                $receivedDate = $date->format('Y-m-d');
                $receivedTime = $date->format('H:i:s');
            }
        }

        // Agregar los datos al arreglo
        $data[] = [
            'Marca temporal' => $timestamp,
            'Correo electronico' => $from,
            'Usuario Final' => '',
            'Fecha de recibido' => $receivedDate,
            'Hora de recibido' => $receivedTime,
            'AGENTE' => '',
            'Prioridad' => '',
            'Nivel' => '',
            'Nivel académico' => '',
            'Programa académico' => '',
            'ASUNTO' => $subject,
            'CREACIÓN DE USUARIOS' => '',
            'NO PERTINENTE' => '',
            'CANVAS	GOOGLE' => '',
            'EXTENSIÓN' => '',
            'EDUCATIC' => '',
            'CERTIFICADOS' => '',
            'PROVEEDORES' => '',
            'PLATAFORMAS' => '',
            'ATENCIÓN AL CIUDADANO' => '',
            'Fecha de solución' => '',
            'Hora de solución' => '',
            'Canal de Ingreso' => '',
            '¿Se envio Encuesta ?' => '',
            'Hora de recibido' => '',
            'Fecha total de inicio' => '',
            'Fecha total de fin' => '',
            'Total tiempo de Solución' => ''
        ];
    }

    // Imprimir los datos
    echo '<pre>';
    print_r($data);
    echo '</pre>';



    // Crear el rango en el que insertar los datos
    $range = 'Hoja1!A1'; // Cambia "Hoja1" al nombre de tu hoja
    $body = new ValueRange([
        'values' => $data
    ]);

    // Escribir los datos en la hoja de cálculo
    $params = ['valueInputOption' => 'RAW'];
    $sheetsService->spreadsheets_values->update($spreadsheetId, $range, $body, $params);

    echo "Datos insertados correctamente en Google Sheets.";
} catch (Exception $e) {
    echo 'Error: ' . $e->getMessage();
}
?>