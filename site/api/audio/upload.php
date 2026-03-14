<?php
/**
 * 音声ファイルアップロードAPI
 *
 * POST /site/api/audio/upload.php
 * Headers: X-Audio-Token: {APIトークン}
 *
 * 通常アップロード (multipart/form-data):
 *   - file: 音声ファイル
 *   - row: スプレッドシート行番号
 *
 * チャンクアップロード (multipart/form-data):
 *   - chunk: チャンクデータ
 *   - row: スプレッドシート行番号
 *   - chunkIndex: チャンク番号 (0-based)
 *   - totalChunks: 合計チャンク数
 *   - fileName: 元ファイル名
 *   - uploadId: アップロード識別子
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

// チャンクアップロード判定
$isChunked = isset($_POST['chunkIndex']) && isset($_POST['totalChunks']);

if ($isChunked) {
    handleChunkedUpload();
} else {
    handleNormalUpload();
}

/**
 * 通常アップロード（小さいファイル向け）
 */
function handleNormalUpload() {
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

    if ($file['size'] > MAX_FILE_SIZE) {
        http_response_code(400);
        echo json_encode(['success' => false, 'message' => 'ファイルサイズが上限（100MB）を超えています']);
        exit;
    }

    $originalName = $file['name'];
    $ext = strtolower(pathinfo($originalName, PATHINFO_EXTENSION));
    if (!in_array($ext, ALLOWED_EXTENSIONS)) {
        http_response_code(400);
        echo json_encode(['success' => false, 'message' => '許可されていないファイル形式です。対応形式: ' . implode(', ', ALLOWED_EXTENSIONS)]);
        exit;
    }

    $yearMonth = date('Ym');
    $saveDir = AUDIO_FILES_DIR . $yearMonth . '/';
    if (!is_dir($saveDir)) {
        if (!mkdir($saveDir, 0755, true)) {
            http_response_code(500);
            echo json_encode(['success' => false, 'message' => '保存ディレクトリの作成に失敗しました']);
            exit;
        }
    }

    $timestamp = date('YmdHis');
    $fileName = $row . '_' . $timestamp . '.' . $ext;
    $savePath = $saveDir . $fileName;

    if (!move_uploaded_file($file['tmp_name'], $savePath)) {
        http_response_code(500);
        echo json_encode(['success' => false, 'message' => 'ファイルの保存に失敗しました']);
        exit;
    }

    $downloadUrl = AUDIO_BASE_URL . 'download.php?file=' . urlencode($yearMonth . '/' . $fileName);

    echo json_encode([
        'success' => true,
        'url' => $downloadUrl,
        'fileName' => $fileName,
        'size' => $file['size'],
        'message' => '音声ファイルをアップロードしました'
    ]);
}

/**
 * チャンクアップロード（大きいファイル対応）
 */
function handleChunkedUpload() {
    $row = isset($_POST['row']) ? intval($_POST['row']) : 0;
    $chunkIndex = intval($_POST['chunkIndex']);
    $totalChunks = intval($_POST['totalChunks']);
    $originalName = isset($_POST['fileName']) ? $_POST['fileName'] : '';
    $uploadId = isset($_POST['uploadId']) ? $_POST['uploadId'] : '';

    if ($row < 2 || $totalChunks < 1 || empty($originalName) || empty($uploadId)) {
        http_response_code(400);
        echo json_encode(['success' => false, 'message' => 'パラメータが不足しています']);
        exit;
    }

    // uploadId をサニタイズ（英数字とハイフンのみ許可）
    $uploadId = preg_replace('/[^a-zA-Z0-9\-]/', '', $uploadId);
    if (empty($uploadId)) {
        http_response_code(400);
        echo json_encode(['success' => false, 'message' => '不正なuploadId']);
        exit;
    }

    $ext = strtolower(pathinfo($originalName, PATHINFO_EXTENSION));
    if (!in_array($ext, ALLOWED_EXTENSIONS)) {
        http_response_code(400);
        echo json_encode(['success' => false, 'message' => '許可されていないファイル形式です']);
        exit;
    }

    // チャンク一時保存ディレクトリ
    $tmpDir = AUDIO_FILES_DIR . 'tmp/' . $uploadId . '/';
    if (!is_dir($tmpDir)) {
        if (!mkdir($tmpDir, 0755, true)) {
            http_response_code(500);
            echo json_encode(['success' => false, 'message' => '一時ディレクトリの作成に失敗しました']);
            exit;
        }
    }

    // チャンクデータ保存
    if (!isset($_FILES['chunk']) || $_FILES['chunk']['error'] !== UPLOAD_ERR_OK) {
        $errCode = isset($_FILES['chunk']) ? $_FILES['chunk']['error'] : 'no chunk';
        http_response_code(400);
        echo json_encode(['success' => false, 'message' => 'チャンクの受信に失敗しました (code: ' . $errCode . ')']);
        exit;
    }

    $chunkPath = $tmpDir . sprintf('%04d', $chunkIndex);
    if (!move_uploaded_file($_FILES['chunk']['tmp_name'], $chunkPath)) {
        http_response_code(500);
        echo json_encode(['success' => false, 'message' => 'チャンクの保存に失敗しました']);
        exit;
    }

    // 全チャンク揃ったか確認
    $receivedChunks = count(glob($tmpDir . '*'));

    if ($receivedChunks < $totalChunks) {
        // まだ残りのチャンクがある
        echo json_encode([
            'success' => true,
            'complete' => false,
            'received' => $receivedChunks,
            'total' => $totalChunks,
            'message' => "チャンク {$receivedChunks}/{$totalChunks} 受信完了"
        ]);
        return;
    }

    // 全チャンク受信完了 → 結合
    $yearMonth = date('Ym');
    $saveDir = AUDIO_FILES_DIR . $yearMonth . '/';
    if (!is_dir($saveDir)) {
        mkdir($saveDir, 0755, true);
    }

    $timestamp = date('YmdHis');
    $fileName = $row . '_' . $timestamp . '.' . $ext;
    $savePath = $saveDir . $fileName;

    // チャンクを結合
    $fp = fopen($savePath, 'wb');
    if (!$fp) {
        http_response_code(500);
        echo json_encode(['success' => false, 'message' => 'ファイル結合に失敗しました']);
        cleanupChunks($tmpDir);
        return;
    }

    for ($i = 0; $i < $totalChunks; $i++) {
        $chunkFile = $tmpDir . sprintf('%04d', $i);
        if (!file_exists($chunkFile)) {
            fclose($fp);
            unlink($savePath);
            http_response_code(500);
            echo json_encode(['success' => false, 'message' => "チャンク {$i} が見つかりません"]);
            cleanupChunks($tmpDir);
            return;
        }
        fwrite($fp, file_get_contents($chunkFile));
    }
    fclose($fp);

    // ファイルサイズ検証
    $finalSize = filesize($savePath);
    if ($finalSize > MAX_FILE_SIZE) {
        unlink($savePath);
        cleanupChunks($tmpDir);
        http_response_code(400);
        echo json_encode(['success' => false, 'message' => 'ファイルサイズが上限（100MB）を超えています']);
        return;
    }

    // 一時ファイル削除
    cleanupChunks($tmpDir);

    $downloadUrl = AUDIO_BASE_URL . 'download.php?file=' . urlencode($yearMonth . '/' . $fileName);

    echo json_encode([
        'success' => true,
        'complete' => true,
        'url' => $downloadUrl,
        'fileName' => $fileName,
        'size' => $finalSize,
        'message' => '音声ファイルをアップロードしました'
    ]);
}

/**
 * チャンク一時ファイルを削除
 */
function cleanupChunks($tmpDir) {
    if (!is_dir($tmpDir)) return;
    $files = glob($tmpDir . '*');
    foreach ($files as $f) {
        if (is_file($f)) unlink($f);
    }
    rmdir($tmpDir);
    // 親ディレクトリ（tmp/）も空なら削除
    $parentDir = dirname($tmpDir) . '/';
    if (is_dir($parentDir) && count(glob($parentDir . '*')) === 0) {
        rmdir($parentDir);
    }
}
