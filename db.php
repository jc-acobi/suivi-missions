<?php
require_once __DIR__ . '/migrate.php';

/**
 * Connexion à la base de données de l'application « Suivi des missions ».
 *
 * Les identifiants viennent d'un fichier serveur (/var/www/config/suivi-missions.php),
 * jamais du dépôt. La base `suivi_missions` est créée automatiquement si besoin, et les
 * migrations en attente (dossier migrations/) sont appliquées.
 *
 * Utilisation :
 *   require_once __DIR__ . '/db.php';
 *   $bdd = db();
 *   $lignes = $bdd->query("SELECT * FROM ma_table")->fetchAll();
 */
function db(): PDO
{
    static $pdo = null;
    if ($pdo !== null) {
        return $pdo;
    }

    $configPath = '/var/www/config/suivi-missions.php';
    if (!is_file($configPath)) {
        throw new RuntimeException("Configuration de la base introuvable sur le serveur.");
    }
    $cfg = require $configPath;
    $schema = 'suivi_missions';

    $pdo = new PDO(
        "mysql:host={$cfg['host']};charset=utf8mb4",
        $cfg['user'],
        $cfg['pass'],
        [
            PDO::ATTR_ERRMODE            => PDO::ERRMODE_EXCEPTION,
            PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
        ]
    );
    $pdo->exec("CREATE DATABASE IF NOT EXISTS `{$schema}` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci");
    $pdo->exec("USE `{$schema}`");

    migrate($pdo, __DIR__ . '/migrations');

    return $pdo;
}
