<?php
require_once __DIR__ . '/vendor/autoload.php';

// Obtener los datos del formulario
$apellido = $_POST['apellido'];
$nombre = $_POST['nombres'];
$dni = $_POST['dni'];
$curso = $_POST['curso'];
$dia = $_POST['dia'];
$mes = $_POST['mes'];
$ano = $_POST['ano'];
$autoridad = $_POST['autoridad'];

// Cargar la plantilla de Word
$template = new \PhpOffice\PhpWord\TemplateProcessor('plantilla/CONSTANCIAGENERAL2023.docx');

// Reemplazar los campos de reemplazo en la plantilla con los datos del formulario
$template->setValue('apellido', $apellido);
$template->setValue('nombres', $nombre);
$template->setValue('dni', $dni);
$template->setValue('curso', $curso);
$template->setValue('dia', $dia);
$template->setValue('mes', $mes);
$template->setValue('ano', $ano);
$template->setValue('autoridad', $autoridad);

// Generar un nombre de archivo único para el archivo de Word
$filename = $apellido . date('YmdHis') . '.docx';

// Descargar el archivo de Word generado
header('Content-Type: application/octet-stream');
header('Content-Disposition: attachment; filename="' . $filename . '"');

$template->saveAs('php://output');

?>
