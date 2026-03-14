<?php
/**
 * 音声アップロードAPI 設定ファイル
 */

// APIトークン（GAS ScriptProperties の AUDIO_API_TOKEN と同じ値を設定）
define('AUDIO_API_TOKEN', '3dc6abfb3b81bdaf7d21012a785889eda489580bacc7ca2709c2f770e98c3b72');

// 保存ディレクトリ（絶対パス）
// さくらサーバーの場合: /home/ユーザー名/www/site/api/audio/files/
define('AUDIO_FILES_DIR', __DIR__ . '/files/');

// 公開URL ベース
define('AUDIO_BASE_URL', 'https://iba-consulting.jp/site/api/audio/');

// 許可する音声ファイル形式
define('ALLOWED_EXTENSIONS', ['mp3', 'm4a', 'wav', 'ogg', 'webm']);

// MIMEタイプマッピング
define('MIME_TYPES', [
    'mp3'  => 'audio/mpeg',
    'm4a'  => 'audio/mp4',
    'wav'  => 'audio/wav',
    'ogg'  => 'audio/ogg',
    'webm' => 'audio/webm'
]);

// ファイルサイズ上限（バイト） — 100MB
define('MAX_FILE_SIZE', 100 * 1024 * 1024);

// CORSを許可するオリジン
define('ALLOWED_ORIGINS', [
    'https://iba-consulting.jp',
    'https://script.google.com'
]);
