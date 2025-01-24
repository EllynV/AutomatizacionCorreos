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
    foreach ($emailLabels as $label) {
        if (str_starts_with($label, $labelName . "/")) {
            // Prioridad: Retorna la parte después de $labelName . "/"
            return substr($label, strlen($labelName . "/"));
        }
    }
    
    // Si no se encontró con "/", verificamos si existe una coincidencia exacta
    foreach ($emailLabels as $label) {
        if ($label === $labelName) {
            return $label;
        }
    }
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
    $gmailService = getGmailService($client);
    // Obtener el servicio de Sheets
    $sheetsService = getSheetsService($client);

    /* echo '<pre style="font-family: calibri; font-size: 11px">';
    print_r($client->getAccessToken());
    echo '</pre>'; */
    
    // Obtener los mensajes
    $messages = $gmailService->users_messages->listUsersMessages('me', ['maxResults' => 5]);

    $data = []; // Array para almacenar los datos extraídos

    $labels = $gmailService->users_labels->listUsersLabels('me');
    $userLabels = [];
    foreach ($labels->getLabels() as $label) {
        $userLabels[$label->getId()] = $label->getName();
    }
    /* echo '<pre style="font-family: calibri; font-size: 11px">';
    print_r($userLabels);
    echo '</pre>'; */

    /* $interestedLabels = ['*Agentes', '*Prioridades'];
    echo '<pre style="font-family: calibri; font-size: 11px">';
    print_r($interestedLabels);
    echo '</pre>'; */

    //echo "<h1>Correos electrónicos</h1>";
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
        //echo "<h1>Etiquetas del correo</h1>";
        $emailLabels = [];
        foreach ($labelIds as $labelId) {
            if (isset($userLabels[$labelId])) {
                $emailLabels[] = $userLabels[$labelId];
            }
        }
        /* echo '<pre style="font-family: calibri; font-size: 11px">';
        print_r($emailLabels);
        echo '</pre>'; */

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
                $receivedDate = $date->format('m-d-Y');
                $receivedTime = $date->format('H:i:s');
            }
        }

        // Agregar los datos al arreglo
        $data['values'][] = [
            'Marca temporal' => $timestamp,
            'Correo electronico' => $from,
            'Usuario Final' => NULL,
            'Fecha de recibido' => $receivedDate,
            'Hora de recibido' => $receivedTime,
            'AGENTE' => getInterestedLabel($emailLabels, '*Agentes'), //etiqueta
            'Prioridad' => getInterestedLabel($emailLabels, '*Prioridades'), //etiqueta
            'Nivel' => NULL, //etiqueta
            'Nivel academico' => NULL,
            'Programa academico' => NULL,
            'ASUNTO' => $subject,
            'CREACION DE USUARIOS' => NULL, //etiqueta
            'NO PERTINENTE' => NULL, //etiqueta
            'CANVAS' => NULL, //etiqueta
            'GOOGLE' => NULL, //etiqueta
            'EXTENSION' => NULL,//etiqueta
            'EDUCATIC' => NULL,//etiqueta
            'CERTIFICADOS' => NULL,//etiqueta
            'PROVEEDORES' => NULL,//etiqueta
            'PLATAFORMAS' => NULL,//etiqueta
            'ATENCION AL CIUDADANO' => NULL,
            'Fecha de solucion' => NULL,
            'Hora de solucion' => NULL,
            'Canal de Ingreso' => NULL,
            'Se envio Encuesta' => NULL,
            'Hora de recibido' => NULL,
            'Fecha total de inicio' => NULL,
            'Fecha total de fin' => NULL,
            'Total tiempo de Solucion' => NULL
        ];
    }

    // Imprimir los datos
    /* echo '<pre style="font-family: calibri; font-size: 11px">';
    print_r($data);
    echo '</pre>'; */

    // Crear el rango en el que insertar los datos
    $range = 'Hoja1!A2';

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
