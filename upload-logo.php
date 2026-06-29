<?php
// Crée le dossier logos/ si besoin
$dir = __DIR__ . '/logos';
if (!is_dir($dir)) mkdir($dir, 0755, true);

header('Content-Type: application/json; charset=utf-8');

if ($_SERVER['REQUEST_METHOD'] !== 'POST' || empty($_FILES['logo'])) {
    http_response_code(400);
    echo json_encode(['error' => 'Aucun fichier reçu']);
    exit;
}

$file = $_FILES['logo'];

// Vérification type MIME
$allowed = ['image/jpeg', 'image/png', 'image/gif', 'image/webp', 'image/svg+xml'];
$finfo = finfo_open(FILEINFO_MIME_TYPE);
$mime  = finfo_file($finfo, $file['tmp_name']);
finfo_close($finfo);

if (!in_array($mime, $allowed)) {
    http_response_code(400);
    echo json_encode(['error' => 'Type de fichier non autorisé']);
    exit;
}

// Limite 2 Mo
if ($file['size'] > 2 * 1024 * 1024) {
    http_response_code(400);
    echo json_encode(['error' => 'Fichier trop grand (max 2 Mo)']);
    exit;
}

// Nom unique basé sur le hash du contenu
$ext      = pathinfo($file['name'], PATHINFO_EXTENSION) ?: 'png';
$hash     = md5_file($file['tmp_name']);
$filename = $hash . '.' . strtolower($ext);
$dest     = $dir . '/' . $filename;

if (!file_exists($dest)) {
    move_uploaded_file($file['tmp_name'], $dest);
}

// Chemin relatif utilisable comme src dans le navigateur
$url = 'logos/' . $filename;
echo json_encode(['url' => $url]);
