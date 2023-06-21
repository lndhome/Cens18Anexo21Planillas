<?php
require_once __DIR__ . '/vendor/autoload.php';

// Obtener los datos del formulario
$apellido = $_POST['apellido'];
$nombre = $_POST['nombres'];
$dni = $_POST['dni'];

$NacDia = $_POST['fecha-nacimiento-dia'];
$NacMes = $_POST['fecha-nacimiento-mes'];
$NacAno = $_POST['fecha-nacimiento-ano'];

$CurPrimero = $_POST['CursadaPrimero'];
$CurSegundo = $_POST['CursadaSegundo'];
$CurTercero = $_POST['CursadaTercero'];

$curso = $_POST['curso'];
$CursoCompleto = $_POST['CursoCompleto'];
$InstPase = $_POST['InstPase'];
$titulo = $_POST['titulo'];
$dia = $_POST['dia'];
$mes = $_POST['mes'];
$ano = $_POST['ano'];


// Cargar la plantilla de Word
$template = new \PhpOffice\PhpWord\TemplateProcessor('plantilla/PASE.INT.docx');

// Reemplazar los campos de reemplazo en la plantilla con los datos del formulario
$template->setValue('apellido', $apellido);
$template->setValue('nombres', $nombre);
$template->setValue('dni', $dni);

$template->setValue('NacDia', $NacDia);
$template->setValue('NacMes', $Nacmes);
$template->setValue('NacAno', $NacAno);

$template->setValue('CursadaPrimero', $CurPrimero);
$template->setValue('CursadaSegundo', $CurSegundo);
$template->setValue('CursadaTercero', $CurTercero);

$template->setValue('curso', $curso);
$template->setValue('CursoCompleto', $CursoCompleto);
$template->setValue('InstPase', $InstPase);
$template->setValue('titulo', $titulo);
$template->setValue('dia', $dia);
$template->setValue('mes', $mes);
$template->setValue('ano', $ano);


// Generar un nombre de archivo único para el archivo de Word
$filename = 'Pase Intearno ' . $apellido . $nombre . date('YmdHis') . '.docx';

// Descargar el archivo de Word generado
header('Content-Type: application/octet-stream');
header('Content-Disposition: attachment; filename="' . $filename . '"');

$template->saveAs('php://output');

?>
