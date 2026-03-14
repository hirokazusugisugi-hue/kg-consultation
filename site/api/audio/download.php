<?php
/**
 * 音声ファイルダウンロード/再生API
 *
 * GET /site/api/audio/download.php?file={path}&token={APIトークン}
 *
 * トークン検証後にファイルを配信。
 * ブラウザでの音声再生にも対応（Content-Typeヘッダー設定）。
 */

require_once __DIR__ . '/config.php';

// CORS対応
$origin = isset($_SERVER['HTTP_ORIGIN']) ? $_SERVER['HTTP_ORIGIN'] : '';
if (in_array($origin, ALLOWED_ORIGINS)) {
    header('Access-Control-Allow-Origin: ' . $origin);
}
header('Access-Control-Allow-Methods: GET, OPTIONS');
header('Access-Control-Allow-Headers: X-Audio-Token');

// preflight
if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    http_response_code(204);
    exit;
}

// トークン検証（ヘッダーまたはクエリパラメータ）
$token = isset($_SERVER['HTTP_X_AUDIO_TOKEN'])
    ? $_SERVER['HTTP_X_AUDIO_TOKEN']
    : (isset($_GET['token']) ? $_GET['token'] : '');

if ($token !== AUDIO_API_TOKEN) {
    http_response_code(403);
    header('Content-Type: application/json; charset=utf-8');
    echo json_encode(['success' => false, 'message' => '認証エラー']);
    exit;
}

// ファイルパス検証
$filePath = isset($_GET['file']) ? $_GET['file'] : '';
if (empty($filePath)) {
    http_response_code(400);
    header('Content-Type: application/json; charset=utf-8');
    echo json_encode(['success' => false, 'message' => 'fileパラメータが必要です']);
    exit;
}

// ディレクトリトラバーサル防止
$filePath = str_replace(['..', "\0"], '', $filePath);
$fullPath = AUDIO_FILES_DIR . $filePath;
$realPath = realpath($fullPath);
$realFilesDir = realpath(AUDIO_FILES_DIR);

if ($realPath === false || strpos($realPath, $realFilesDir) !== 0) {
    http_response_code(404);
    header('Content-Type: application/json; charset=utf-8');
    echo json_encode(['success' => false, 'message' => 'ファイルが見つかりません']);
    exit;
}

if (!is_file($realPath)) {
    http_response_code(404);
    header('Content-Type: application/json; charset=utf-8');
    echo json_encode(['success' => false, 'message' => 'ファイルが見つかりません']);
    exit;
}

// MIMEタイプ判定
$ext = strtolower(pathinfo($realPath, PATHINFO_EXTENSION));
$mimeType = isset(MIME_TYPES[$ext]) ? MIME_TYPES[$ext] : 'application/octet-stream';

// ファイル配信
$fileSize = filesize($realPath);
header('Content-Type: ' . $mimeType);
header('Content-Length: ' . $fileSize);
header('Accept-Ranges: bytes');
header('Cache-Control: private, max-age=3600');

// Range対応（音声のシーク再生用）
if (isset($_SERVER['HTTP_RANGE'])) {
    preg_match('/bytes=(\d+)-(\d*)/', $_SERVER['HTTP_RANGE'], $matches);
    $start = intval($matches[1]);
    $end = !empty($matches[2]) ? intval($matches[2]) : ($fileSize - 1);

    if ($start > $end || $start >= $fileSize) {
        http_response_code(416);
        header('Content-Range: bytes */' . $fileSize);
        exit;
    }

    http_response_code(206);
    header('Content-Range: bytes ' . $start . '-' . $end . '/' . $fileSize);
    header('Content-Length: ' . ($end - $start + 1));

    $fp = fopen($realPath, 'rb');
    fseek($fp, $start);
    $remaining = $end - $start + 1;
    while ($remaining > 0 && !feof($fp)) {
        $chunk = min(8192, $remaining);
        echo fread($fp, $chunk);
        $remaining -= $chunk;
    }
    fclose($fp);
} else {
    readfile($realPath);
}
