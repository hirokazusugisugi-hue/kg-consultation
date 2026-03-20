<?php
/**
 * 画像アップロードAPI 設定ファイル
 */

// APIトークン（GAS ScriptProperties の IMAGE_API_TOKEN と同じ値を設定）
define('IMAGE_API_TOKEN', 'img_7f3a9c2e1b5d4f8e6a0c3b7d9e2f1a4b5c8d6e0f3a7b9c2d4e6f8a1b3c5d7e');

// 保存ディレクトリ（絶対パス）
define('IMAGE_FILES_DIR', __DIR__ . '/files/');

// 公開URL ベース
define('IMAGE_BASE_URL', 'https://iba-consulting.jp/site/api/image/files/');

// 許可する画像ファイル形式
define('ALLOWED_EXTENSIONS', ['jpg', 'jpeg', 'png', 'gif', 'webp']);

// ファイルサイズ上限（バイト） — 10MB
define('MAX_FILE_SIZE', 10 * 1024 * 1024);

// CORSを許可するオリジン
define('ALLOWED_ORIGINS', [
    'https://iba-consulting.jp',
    'https://script.google.com',
    'https://hirokazusugisugi-hue.github.io'
]);
