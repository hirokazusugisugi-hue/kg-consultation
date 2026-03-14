<?php
/**
 * 音声ファイルアップロードAPI
 *
 * POST /site/api/audio/upload.php
 * Headers: X-Audio-Token: {APIトークン}
 * Body: multipart/form-data
 *   - file: 音声ファイル
 *   - row: スプレッドシート行番号
 */

require_once __DIR__ . '/config.php';

// CORS対応
$origin = isset($_SERVER['HTTP_ORIGIN']) ? $_SERVER['HTTP_ORIGIN'] : '';
if (in_array($origin, ALLOWED_ORIGINS)) {
    header('Access-Control-Allow-Origin: ' . $origin);
}
header('Access-Control-Allow-Methods: POST, OPTIONS');
header('Access-Control-Allow-Headers: X-Audio-Token, Content-Type');
header('Content-Type: application/json; charset=utf-8');

// preflight
if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    http_response_code(204);
    exit;
}

// POSTのみ
if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    http_response_code(405);
    echo json_encode(['success' => false, 'message' => 'Method Not Allowed']);
    exit;
}

// トークン検証
$token = isset($_SERVER['HTTP_X_AUDIO_TOKEN']) ? $_SERVER['HTTP_X_AUDIO_TOKEN'] : '';
if ($token !== AUDIO_API_TOKEN) {
    http_response_code(403);
    echo json_encode(['success' => false, 'message' => '認証エラー']);
    exit;
}

// ファイル検証
if (!isset($_FILES['file']) || $_FILES['file']['error'] !== UPLOAD_ERR_OK) {
    $errCode = isset($_FILES['file']) ? $_FILES['file']['error'] : 'no file';
    http_response_code(400);
    echo json_encode(['success' => false, 'message' => 'ファイルのアップロードに失敗しました (code: ' . $errCode . ')']);
    exit;
}

$file = $_FILES['file'];
$row = isset($_POST['row']) ? intval($_POST['row']) : 0;

if ($row < 2) {
    http_response_code(400);
    echo json_encode(['success' => false, 'message' => '行番号が不正です']);
    exit;
}

// ファイルサイズチェック
if ($file['size'] > MAX_FILE_SIZE) {
    http_response_code(400);
    echo json_encode(['success' => false, 'message' => 'ファイルサイズが上限（100MB）を超えています']);
    exit;
}

// 拡張子チェック
$originalName = $file['name'];
$ext = strtolower(pathinfo($originalName, PATHINFO_EXTENSION));
if (!in_array($ext, ALLOWED_EXTENSIONS)) {
    http_response_code(400);
    echo json_encode(['success' => false, 'message' => '許可されていないファイル形式です。対応形式: ' . implode(', ', ALLOWED_EXTENSIONS)]);
    exit;
}

// 保存ディレクトリ作成（年月サブディレクトリ）
$yearMonth = date('Ym');
$saveDir = AUDIO_FILES_DIR . $yearMonth . '/';
if (!is_dir($saveDir)) {
    if (!mkdir($saveDir, 0755, true)) {
        http_response_code(500);
        echo json_encode(['success' => false, 'message' => '保存ディレクトリの作成に失敗しました']);
        exit;
    }
}

// ファイル名生成: {row}_{タイムスタンプ}.{ext}
$timestamp = date('YmdHis');
$fileName = $row . '_' . $timestamp . '.' . $ext;
$savePath = $saveDir . $fileName;

// ファイル移動
if (!move_uploaded_file($file['tmp_name'], $savePath)) {
    http_response_code(500);
    echo json_encode(['success' => false, 'message' => 'ファイルの保存に失敗しました']);
    exit;
}

// ダウンロードURL生成（download.php経由）
$downloadUrl = AUDIO_BASE_URL . 'download.php?file=' . urlencode($yearMonth . '/' . $fileName);

echo json_encode([
    'success' => true,
    'url' => $downloadUrl,
    'fileName' => $fileName,
    'size' => $file['size'],
    'message' => '音声ファイルをアップロードしました'
]);
