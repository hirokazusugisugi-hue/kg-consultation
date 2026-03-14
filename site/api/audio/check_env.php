<?php
/**
 * さくらサーバー環境チェックスクリプト
 *
 * アクセス: https://iba-consulting.jp/site/api/audio/check_env.php?token=YOUR_TOKEN
 * 確認後は削除してください。
 */

require_once __DIR__ . '/config.php';

$token = isset($_GET['token']) ? $_GET['token'] : '';
if ($token !== AUDIO_API_TOKEN) {
    http_response_code(403);
    echo 'Forbidden';
    exit;
}

header('Content-Type: text/plain; charset=utf-8');

echo "=== さくらサーバー環境チェック ===\n\n";

// 1. PHP設定
echo "--- PHP設定 ---\n";
echo "PHP Version: " . phpversion() . "\n";
echo "upload_max_filesize: " . ini_get('upload_max_filesize') . "\n";
echo "post_max_size: " . ini_get('post_max_size') . "\n";
echo "max_execution_time: " . ini_get('max_execution_time') . "\n";
echo "memory_limit: " . ini_get('memory_limit') . "\n\n";

// 2. Python3
echo "--- Python3 ---\n";
$python3 = trim(shell_exec('which python3 2>&1') ?? '');
echo "which python3: " . ($python3 ?: '未検出') . "\n";
if ($python3 && strpos($python3, 'not found') === false) {
    echo "python3 version: " . trim(shell_exec('python3 --version 2>&1') ?? '') . "\n";

    // google-cloud-speech チェック
    $check_speech = trim(shell_exec('python3 -c "import google.cloud.speech_v2; print(\'OK\')" 2>&1') ?? '');
    echo "google-cloud-speech: " . $check_speech . "\n";

    // google-cloud-storage チェック
    $check_storage = trim(shell_exec('python3 -c "import google.cloud.storage; print(\'OK\')" 2>&1') ?? '');
    echo "google-cloud-storage: " . $check_storage . "\n";
} else {
    echo "※ Python3が見つかりません\n";
}
echo "\n";

// 3. exec() 利用可否
echo "--- exec() ---\n";
$disabled = ini_get('disable_functions');
echo "disable_functions: " . ($disabled ?: '(なし)') . "\n";
$exec_disabled = (strpos($disabled, 'exec') !== false);
echo "exec() 利用: " . ($exec_disabled ? 'NG（無効化されています）' : 'OK') . "\n";
if (!$exec_disabled) {
    $test = shell_exec('echo "exec test OK" 2>&1');
    echo "exec テスト: " . trim($test ?? '(出力なし)') . "\n";
}
echo "\n";

// 4. ディレクトリ
echo "--- ディレクトリ ---\n";
$filesDir = AUDIO_FILES_DIR;
echo "files/ ディレクトリ: " . $filesDir . "\n";
echo "  存在: " . (is_dir($filesDir) ? 'OK' : 'NG') . "\n";
echo "  書込可能: " . (is_writable($filesDir) ? 'OK' : 'NG') . "\n";

$logsDir = __DIR__ . '/logs/';
echo "logs/ ディレクトリ: " . $logsDir . "\n";
echo "  存在: " . (is_dir($logsDir) ? 'OK' : 'NG') . "\n";
echo "  書込可能: " . (is_writable($logsDir) ? 'OK' : 'NG') . "\n";
echo "\n";

// 5. GCS鍵ファイル
echo "--- GCP ---\n";
$keyPath = getenv('GOOGLE_APPLICATION_CREDENTIALS') ?: ($_SERVER['HOME'] ?? '/home') . '/.gcp/speech-to-text-key.json';
echo "サービスアカウント鍵: " . $keyPath . "\n";
echo "  存在: " . (file_exists($keyPath) ? 'OK' : 'NG（未配置）') . "\n";
echo "\n";

// 6. .htaccess
echo "--- .htaccess ---\n";
echo "存在: " . (file_exists(__DIR__ . '/.htaccess') ? 'OK' : 'NG') . "\n";
echo ".user.ini 存在: " . (file_exists(__DIR__ . '/.user.ini') ? 'OK' : 'NG') . "\n";
echo "\n";

echo "=== チェック完了 ===\n";
echo "※ このファイルは確認後に削除してください。\n";
