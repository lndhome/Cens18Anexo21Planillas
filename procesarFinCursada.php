<?php
require_once __DIR__ . '/vendor/autoload.php';

// Obtener los datos del formulario
$apellido = $_POST['apellido'];
$nombre = $_POST['nombres'];
$dni = $_POST['dni'];
$FechaNacDia = $_POST['fecha-nacimiento-dia'];
$FechaNacMes = $_POST['fecha-nacimiento-mes'];
$FechaNacAno = $_POST['fecha-nacimiento-ano'];
$LugarNac = $_POST['LugarNac'];
$MateriaAd = $_POST['MateriaAd'];
$modalidad = $_POST['modalidad'];
$titulo = $_POST['titulo'];
$dia = $_POST['dia'];
$mes = $_POST['mes'];
$ano = $_POST['ano'];

$fechaFormateada = date("d \\de F \\del Y", strtotime($fechaNac));
// Cargar la plantilla de Word
$template = new \PhpOffice\PhpWord\TemplateProcessor('plantilla/CERTI.FINALIZACIÓN.CURSADA.docx');

// Reemplazar los campos de reemplazo en la plantilla con los datos del formulario
$template->setValue('apellido', $apellido);
$template->setValue('nombres', $nombre);
$template->setValue('dni', $dni);
$template->setValue('FechaNacDia', $FechaNacDia);
$template->setValue('FechaNacMes', $FechaNacMes);
$template->setValue('FechaNacAno', $FechaNacAno);
$template->setValue('LugarNac', $LugarNac);
$template->setValue('MateriaAd', $MateriaAd);
$template->setValue('modalidad', $modalidad);
$template->setValue('titulo', $titulo);
$template->setValue('dia', $dia);
$template->setValue('mes', $mes);
$template->setValue('ano', $ano);


// Generar un nombre de archivo único para el archivo de Word
$filename = 'Cons. finalizacion cursada ' . $apellido . date('YmdHis') . '.docx';

// Descargar el archivo de Word generado
header('Content-Type: application/octet-stream');
header('Content-Disposition: attachment; filename="' . $filename . '"');

$template->saveAs('php://output');

?>
