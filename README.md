# 写真・動画スライドショー自動作成ツール (`slideshow`)

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python: >=3.14](https://img.shields.io/badge/python-%3E%3D3.14-blue.svg)](https://www.python.org/)

指定したフォルダーやZIPファイル内の写真（画像）と動画を撮影日時順に自動ソートし、PowerPoint形式（`.pptx`）のスライドショーを自動生成するPythonツールです。

白を基調としたすっきりとしたデザインを採用しており、スライドショー開始時に自動再生され、最後まで行くと自動的に最初に戻るループ再生設定が組み込まれます。店頭ディスプレイ、展示会、家族の思い出の鑑賞などに最適です。

---

## 主な特徴

- **📅 撮影日時順での自動ソート**
  - 写真のEXIF情報（`DateTimeOriginal` / `DateTime`）を解析して古い順に並べ替えます。
  - EXIF情報が無い画像や動画は、ファイルの更新日時（`mtime`）を基準にソートします。
- **🔄 自動再生トランジションと無限ループ再生**
  - **画像スライド**: 指定した表示時間（既定値: 3秒）で自動的に次のスライドへ進みます。
  - **動画スライド**: 動画ファイルのバイナリを高速解析して秒数を動的に取得し、再生完了後（+0.5秒）に自動で次へ遷移します。
  - **無限ループ**: プレゼンテーションのループ再生設定が自動適用され、最後のスライドに達すると自動的に最初の表紙に戻ります（`Esc`キーで終了）。
- **🎬 動画の自動再生**
  - スライド切り替え時に、埋め込まれた動画が自動的に再生開始します。
- **📐 縦横アスペクト比と向きの自動調整**
  - EXIFの回転情報（Orientation）を自動認識して正立配置します。
  - スライド（16:9 ワイドスクリーン）の中央にアスペクト比を維持したまま最大サイズで配置されます。
- **📦 フォルダ＆ZIPファイル対応**
  - 画像・動画が入ったディレクトリはもちろん、ZIPアーカイブを直接指定して生成することも可能です。
- **🏷️ ファイル名キャプション**
  - 各スライドの下部にファイル名を表示可能（`--no-filename` で非表示にすることもできます）。

---

## 対応フォーマット

- **画像**: `.jpg`, `.jpeg`, `.png`, `.webp`, `.bmp`, `.gif`, `.tif`, `.tiff`
- **動画**: `.mp4`, `.mov`, `.avi`, `.wmv`, `.m4v`, `.mpg`, `.mpeg`

---

## 動作要件

- **Python**: 3.14 以上
- **パッケージマネージャー**: [uv](https://docs.astral.sh/uv/)（推奨）または `pip`
- **主要依存ライブラリ**:
  - `pillow` (画像の読み込み・EXIF解析・回転補正)
  - `python-pptx` (PowerPointファイルの生成・XML構造編集)

---

## セットアップ

### 1. リポジトリのクローン

```bash
git clone https://github.com/matsuda-works/slideshow.git
cd slideshow
```

### 2. 依存関係のインストール

[`uv`](https://docs.astral.sh/uv/) を利用する場合:
```bash
uv sync
```

または `pip` を利用する場合:
```bash
pip install -r pyproject.toml
```

---

## 使い方

### 基本的な実行

プロジェクト直下の `photo` フォルダに写真や動画を配置し、以下を実行します。
（既定値では、表示時間3秒、タイトル「フォトスライドショー」で `slideshow.pptx` が生成されます）

```bash
uv run python main.py
```

### オプションを指定して実行

```bash
uv run python main.py [入力パス] [オプション]
```

#### コマンドライン引数一覧

| 引数 | 短縮 | 既定値 | 説明 |
| :--- | :--- | :--- | :--- |
| `input` | - | `./photo` | 写真・動画が含まれるフォルダ、またはZIPファイルのパス |
| `--output` | `-o` | `slideshow.pptx` | 出力するPowerPointファイルパス（`.pptx`） |
| `--title` | `-t` | `フォトスライドショー` | 表紙スライドの中央に表示するタイトル |
| `--duration`| `-d` | `3.0` | 写真スライドの表示秒数（動画は動画の長さに連動） |
| `--no-filename` | - | `False` | スライド下部のファイル名キャプションを非表示にする |
| `--help` | `-h` | - | ヘルプメッセージを表示 |

#### 実行例

- **任意のフォルダを指定し、タイトルと表示秒数を変更する**:
  ```bash
  uv run python main.py ~/Pictures/Vacation2026 -o vacation.pptx -t "Summer Vacation 2026" -d 5.0
  ```

- **ZIPファイルを直接指定してファイル名キャプションを隠す**:
  ```bash
  uv run python main.py ./archive.zip -o presentation.pptx -t "Memory Collection" --no-filename
  ```

---

## トラブルシューティング

- **PermissionError（ファイル保存エラー）**:
  生成先の `.pptx` ファイルが PowerPoint や他のビューワーで開かれていると、ファイル書き込みがロックされてエラーになります。PowerPoint を閉じてから再実行してください。
- **動画の長さが取得できない場合**:
  非標準フォーマット等の理由で動画の長さを解析できない場合、既定のフォールバック時間（10秒）が設定されます。

---

## ライセンス

本プロジェクトは [MIT License](LICENSE) のもとで公開されています。
詳細は [LICENSE](LICENSE) ファイルをご確認ください。
