<?php
require_once __DIR__ . '/vendor/autoload.php';

// Obtener los datos del formulario
$nombre = $_POST['nombre'];
$email = $_POST['email'];
$mensaje = $_POST['mensaje'];

// Cargar la plantilla de Word
$template = new \PhpOffice\PhpWord\TemplateProcessor('plantilla.docx');

// Reemplazar los campos de reemplazo en la plantilla con los datos del formulario
$template->setValue('nombre', $nombre);
$template->setValue('email', $email);
$template->setValue('mensaje', $mensaje);

// Generar un nombre de archivo único para el archivo de Word
$filename = $nombre . date('YmdHis') . '.docx';

// Descargar el archivo de Word generado
header('Content-Type: application/octet-stream');
header('Content-Disposition: attachment; filename="' . $filename . '"');

$template->saveAs('php://output');

?>
