<?php
/**
 * 画像ファイルアップロードAPI
 *
 * POST /site/api/image/upload.php
 * Headers: X-Image-Token: {APIトークン}
 *
 * multipart/form-data:
 *   - file: 画像ファイル
 */

require_once __DIR__ . '/config.php';

// CORS対応
$origin = isset($_SERVER['HTTP_ORIGIN']) ? $_SERVER['HTTP_ORIGIN'] : '';
if (in_array($origin, ALLOWED_ORIGINS)) {
    header('Access-Control-Allow-Origin: ' . $origin);
}
header('Access-Control-Allow-Methods: POST, OPTIONS');
header('Access-Control-Allow-Headers: X-Image-Token, Content-Type');
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
$token = isset($_SERVER['HTTP_X_IMAGE_TOKEN']) ? $_SERVER['HTTP_X_IMAGE_TOKEN'] : '';
if ($token !== IMAGE_API_TOKEN) {
    http_response_code(403);
    echo json_encode(['success' => false, 'message' => '認証エラー']);
    exit;
}

// ファイルチェック
if (!isset($_FILES['file']) || $_FILES['file']['error'] !== UPLOAD_ERR_OK) {
    $errCode = isset($_FILES['file']) ? $_FILES['file']['error'] : 'no file';
    http_response_code(400);
    echo json_encode(['success' => false, 'message' => 'ファイルのアップロードに失敗しました (code: ' . $errCode . ')']);
    exit;
}

$file = $_FILES['file'];

// サイズチェック
if ($file['size'] > MAX_FILE_SIZE) {
    http_response_code(400);
    echo json_encode(['success' => false, 'message' => 'ファイルサイズが上限（10MB）を超えています']);
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

// 画像バリデーション（getimagesize で偽装防止）
$imageInfo = @getimagesize($file['tmp_name']);
if ($imageInfo === false) {
    http_response_code(400);
    echo json_encode(['success' => false, 'message' => '有効な画像ファイルではありません']);
    exit;
}

// 保存ディレクトリ（年月別）
$yearMonth = date('Ym');
$saveDir = IMAGE_FILES_DIR . $yearMonth . '/';
if (!is_dir($saveDir)) {
    if (!mkdir($saveDir, 0755, true)) {
        http_response_code(500);
        echo json_encode(['success' => false, 'message' => '保存ディレクトリの作成に失敗しました']);
        exit;
    }
}

// ファイル名生成: YYYYMMDDHHmmss_ユニークID.ext
$timestamp = date('YmdHis');
$uniqueId = substr(bin2hex(random_bytes(4)), 0, 8);
$fileName = $timestamp . '_' . $uniqueId . '.' . $ext;
$savePath = $saveDir . $fileName;

if (!move_uploaded_file($file['tmp_name'], $savePath)) {
    http_response_code(500);
    echo json_encode(['success' => false, 'message' => 'ファイルの保存に失敗しました']);
    exit;
}

// 公開URL（直接アクセス可能）
$publicUrl = IMAGE_BASE_URL . $yearMonth . '/' . $fileName;

echo json_encode([
    'success' => true,
    'url' => $publicUrl,
    'fileName' => $fileName,
    'size' => $file['size'],
    'width' => $imageInfo[0],
    'height' => $imageInfo[1],
    'message' => '画像をアップロードしました'
]);
