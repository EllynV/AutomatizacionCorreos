<?php
require 'vendor/autoload.php';

use Google\Client;
use Google\Service\Gmail;

// Iniciar sesión
session_start();

function authenticateGoogleClient($credentialsPath, $redirectUri) {
    $client = new Client();
    $client->setApplicationName('Gmail API PHP Script');
    $client->setScopes(Gmail::GMAIL_READONLY);
    $client->setAuthConfig($credentialsPath);
    $client->setAccessType('offline');
    $client->setPrompt('select_account consent');
    $client->setRedirectUri($redirectUri);

    // Redirigir al usuario para autenticarse
    $authUrl = $client->createAuthUrl();
    header('Location: ' . $authUrl);
    exit;
}

// Configuración
$credentialsPath = 'credentials.json'; // Ruta al archivo JSON de credenciales
$redirectUri = 'http://localhost/get-emails/show.php'; // URI configurada en Google Cloud Console

// Iniciar el flujo de autenticación
authenticateGoogleClient($credentialsPath, $redirectUri);
?>