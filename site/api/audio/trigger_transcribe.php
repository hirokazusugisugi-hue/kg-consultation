<?php
/**
 * 文字起こしトリガーAPI
 *
 * POST /site/api/audio/trigger_transcribe.php
 * Body (JSON):
 *   - token: APIトークン
 *   - audioFile: 音声ファイルの相対パス（files/YYYYMM/xxx.mp3）
 *   - row: スプレッドシート行番号
 *   - webhookUrl: GAS webhook URL（transcribe-callback用）
 *
 * Pythonスクリプトをバックグラウンドで起動して文字起こしを実行する。
 */

require_once __DIR__ . '/config.php';

// CORS対応
$origin = isset($_SERVER['HTTP_ORIGIN']) ? $_SERVER['HTTP_ORIGIN'] : '';
if (in_array($origin, ALLOWED_ORIGINS)) {
    header('Access-Control-Allow-Origin: ' . $origin);
}
header('Access-Control-Allow-Methods: POST, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type');
header('Content-Type: application/json; charset=utf-8');

// preflight
if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    http_response_code(204);
    exit;
}

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    http_response_code(405);
    echo json_encode(['success' => false, 'message' => 'Method Not Allowed']);
    exit;
}

// リクエストボディ解析
$input = json_decode(file_get_contents('php://input'), true);
if (!$input) {
    http_response_code(400);
    echo json_encode(['success' => false, 'message' => 'Invalid JSON']);
    exit;
}

// トークン検証
$token = isset($input['token']) ? $input['token'] : '';
if ($token !== AUDIO_API_TOKEN) {
    http_response_code(403);
    echo json_encode(['success' => false, 'message' => '認証エラー']);
    exit;
}

// パラメータ検証
$audioFile = isset($input['audioFile']) ? $input['audioFile'] : '';
$row = isset($input['row']) ? intval($input['row']) : 0;
$webhookUrl = isset($input['webhookUrl']) ? $input['webhookUrl'] : '';

if (empty($audioFile) || $row < 2 || empty($webhookUrl)) {
    http_response_code(400);
    echo json_encode(['success' => false, 'message' => 'audioFile, row, webhookUrl は必須です']);
    exit;
}

// ファイルパス安全化
$audioFile = str_replace(['..', "\0"], '', $audioFile);
$fullPath = AUDIO_FILES_DIR . $audioFile;

if (!file_exists($fullPath)) {
    http_response_code(404);
    echo json_encode(['success' => false, 'message' => '音声ファイルが見つかりません']);
    exit;
}

// Pythonスクリプトのパス
$scriptPath = __DIR__ . '/transcribe.py';
if (!file_exists($scriptPath)) {
    http_response_code(500);
    echo json_encode(['success' => false, 'message' => '文字起こしスクリプトが見つかりません']);
    exit;
}

// ログファイル
$logDir = __DIR__ . '/logs/';
if (!is_dir($logDir)) {
    mkdir($logDir, 0755, true);
}
$logFile = $logDir . 'transcribe_' . $row . '_' . date('YmdHis') . '.log';

// Pythonスクリプトをバックグラウンド実行
$cmd = sprintf(
    'python3 %s %s %d %s %s > %s 2>&1 &',
    escapeshellarg($scriptPath),
    escapeshellarg($fullPath),
    $row,
    escapeshellarg($webhookUrl),
    escapeshellarg(AUDIO_API_TOKEN),
    escapeshellarg($logFile)
);

exec($cmd);

echo json_encode([
    'success' => true,
    'message' => '文字起こし処理を開始しました',
    'row' => $row
]);
