#!/usr/bin/env python3
"""
Kindle PDF変換ツール

Kindleアプリのスクリーンショットを自動撮影してPDFに変換する

使用方法:
    python kindle_to_pdf.py -n "書籍名"
"""

import subprocess
import time
import sys
import hashlib
import shutil
from pathlib import Path
from datetime import datetime
from typing import Optional


# =============================================================================
# 事前チェック
# =============================================================================

def check_kindle_running() -> bool:
    """Kindleアプリが起動しているか確認"""
    result = subprocess.run(
        ["osascript", "-e",
         'tell application "System Events" to (name of every process) contains "Kindle"'],
        capture_output=True, text=True
    )
    return result.stdout.strip() == "true"


def check_kindle_window_exists() -> bool:
    """Kindleウィンドウが存在するか確認"""
    result = subprocess.run(
        ["osascript", "-e",
         'tell application "System Events" to tell process "Kindle" to count of windows'],
        capture_output=True, text=True
    )
    try:
        return int(result.stdout.strip()) > 0
    except ValueError:
        return False


def get_kindle_window_bounds() -> Optional[tuple[int, int, int, int]]:
    """
    Kindleウィンドウの位置とサイズを取得

    Returns:
        (x, y, width, height) または None
    """
    # まず "Kindle" という名前のウィンドウを試す（メインウィンドウ）
    result = subprocess.run(
        ["osascript", "-e",
         'tell application "System Events" to tell process "Kindle" '
         'to get {position, size} of window "Kindle"'],
        capture_output=True, text=True
    )

    # 失敗したら最大サイズのウィンドウを探す
    if result.returncode != 0:
        result = subprocess.run(
            ["osascript", "-e",
             'tell application "System Events" to tell process "Kindle" '
             'to get {position, size} of window 1'],
            capture_output=True, text=True
        )

    try:
        # "x, y, width, height" の形式でパース
        values = [int(v.strip()) for v in result.stdout.strip().split(",")]
        if len(values) == 4:
            # 高さが100px未満はツールバーの可能性があるのでスキップ
            if values[3] < 100:
                # 別のウィンドウを試す
                result = subprocess.run(
                    ["osascript", "-e",
                     'tell application "System Events" to tell process "Kindle" '
                     'to get {position, size} of window 2'],
                    capture_output=True, text=True
                )
                values = [int(v.strip()) for v in result.stdout.strip().split(",")]
            return tuple(values)  # type: ignore
    except (ValueError, IndexError):
        pass

    return None


def activate_kindle() -> None:
    """Kindleアプリを最前面に持ってくる"""
    subprocess.run(
        ["osascript", "-e", 'tell application "Amazon Kindle" to activate'],
        capture_output=True
    )
    time.sleep(0.3)


def launch_kindle() -> bool:
    """Kindleアプリを起動する"""
    if check_kindle_running():
        activate_kindle()
        return True

    print("  Kindleアプリを起動中...")

    # openコマンドでKindleを起動（アプリ名は「Amazon Kindle」）
    subprocess.run(["open", "-a", "Amazon Kindle"], capture_output=True)
    time.sleep(5)

    # 起動を待つ
    for _ in range(15):
        if check_kindle_running():
            activate_kindle()
            time.sleep(2)
            if check_kindle_window_exists():
                return True
        time.sleep(1)

    return False


def setup_split_screen() -> bool:
    """
    画面を左右分割してターミナル（左）とKindle（右）を配置
    """
    print("  画面を左右分割中...")

    # 画面サイズを取得
    result = subprocess.run(
        ["osascript", "-e", 'tell application "Finder" to get bounds of window of desktop'],
        capture_output=True, text=True
    )

    try:
        bounds = [int(x.strip()) for x in result.stdout.strip().split(",")]
        screen_width = bounds[2]
        screen_height = bounds[3]
    except (ValueError, IndexError):
        # デフォルト値
        screen_width = 1800
        screen_height = 1169

    half_width = screen_width // 2
    menu_bar = 25  # メニューバーの高さ

    # ターミナルを左半分に配置
    terminal_script = f'''
    tell application "Terminal"
        activate
        if (count of windows) > 0 then
            set bounds of front window to {{0, {menu_bar}, {half_width}, {screen_height}}}
        end if
    end tell
    '''
    subprocess.run(["osascript", "-e", terminal_script], capture_output=True)

    # Kindleを右半分に配置
    kindle_script = f'''
    tell application "System Events"
        tell process "Kindle"
            set position of window 1 to {{{half_width}, {menu_bar}}}
            set size of window 1 to {{{half_width}, {screen_height - menu_bar}}}
        end tell
    end tell
    '''
    subprocess.run(["osascript", "-e", kindle_script], capture_output=True)

    time.sleep(0.5)
    print(f"  ✓ 画面分割完了 (左: ターミナル, 右: Kindle)")
    return True


def run_applescript(script: str) -> bool:
    """AppleScriptを実行"""
    result = subprocess.run(["osascript", "-e", script], capture_output=True)
    return result.returncode == 0


def go_to_library() -> bool:
    """Kindleライブラリに戻る"""
    activate_kindle()
    time.sleep(0.5)

    # Cmd+Shift+L でライブラリに戻る
    success = run_applescript('''
    tell application "System Events"
        tell process "Kindle"
            keystroke "l" using {command down, shift down}
        end tell
    end tell
    ''')
    time.sleep(2)
    return success


def search_and_open_book(book_name: str) -> bool:
    """
    Kindleライブラリで書籍を検索して開く

    Args:
        book_name: 検索する書籍名

    Returns:
        成功したらTrue
    """
    print(f"  書籍を検索中: {book_name}")

    # ライブラリに戻る
    go_to_library()
    time.sleep(1)

    # 検索ボックスにフォーカス（Cmd+F）
    run_applescript('''
    tell application "System Events"
        tell process "Kindle"
            keystroke "f" using command down
        end tell
    end tell
    ''')
    time.sleep(0.5)

    # 検索ボックスをクリア（Cmd+A で全選択してから入力）
    run_applescript('''
    tell application "System Events"
        tell process "Kindle"
            keystroke "a" using command down
        end tell
    end tell
    ''')
    time.sleep(0.2)

    # 書籍名を入力
    # 特殊文字をエスケープ
    escaped_name = book_name.replace('"', '\\"').replace("'", "\\'")
    run_applescript(f'''
    tell application "System Events"
        tell process "Kindle"
            keystroke "{escaped_name}"
        end tell
    end tell
    ''')
    time.sleep(1)

    # Enterで検索実行
    run_applescript('''
    tell application "System Events"
        tell process "Kindle"
            keystroke return
        end tell
    end tell
    ''')
    time.sleep(2)

    # 最初の検索結果をダブルクリックして開く
    bounds = get_kindle_window_bounds()
    if bounds:
        x, y, width, height = bounds
        # 検索結果の最初のアイテム位置（おおよそ）
        click_x = x + 150
        click_y = y + 250

        # ダブルクリックで開く
        subprocess.run(["cliclick", f"dc:{click_x},{click_y}"], capture_output=True)
        time.sleep(3)

    print(f"  書籍を開きました: {book_name}")
    return True


def go_to_first_page() -> bool:
    """書籍の最初のページに移動（左矢印キーを連打）"""
    print("  最初のページに移動中...")

    activate_kindle()
    time.sleep(0.3)

    # 左矢印キーを連打して最初のページへ
    for _ in range(20):
        prev_page()
        time.sleep(0.2)

    time.sleep(0.5)
    return True


# =============================================================================
# スクリーンショット
# =============================================================================

def capture_window(output_path: str, bounds: tuple[int, int, int, int]) -> bool:
    """
    指定した矩形領域のスクリーンショットを撮影

    Args:
        output_path: 出力ファイルパス
        bounds: (x, y, width, height)

    Returns:
        成功したらTrue
    """
    x, y, width, height = bounds

    # タイトルバー（約28px）を除外
    title_bar_height = 28
    y += title_bar_height
    height -= title_bar_height

    result = subprocess.run(
        ["screencapture", "-x", "-R", f"{x},{y},{width},{height}", output_path],
        capture_output=True
    )
    return result.returncode == 0


# =============================================================================
# ページ操作
# =============================================================================

def next_page() -> bool:
    """次のページに移動（右矢印キー）"""
    # Kindleをアクティブにしてから矢印キーを送信
    result = subprocess.run(
        ["osascript", "-e",
         'tell application "System Events" to tell process "Kindle" to key code 124'],
        capture_output=True
    )
    return result.returncode == 0


def prev_page() -> bool:
    """前のページに移動（左矢印キー）"""
    result = subprocess.run(
        ["osascript", "-e",
         'tell application "System Events" to tell process "Kindle" to key code 123'],
        capture_output=True
    )
    return result.returncode == 0


def verify_page_turned(old_hash: str, new_path: str) -> tuple[bool, str]:
    """
    ページがめくれたか検証

    Args:
        old_hash: 前のページの画像ハッシュ
        new_path: 新しいページの画像パス

    Returns:
        (ページが変わったか, 新しいハッシュ)
    """
    new_hash = get_image_hash(new_path)
    return old_hash != new_hash, new_hash


# =============================================================================
# 重複検出
# =============================================================================

def get_image_hash(filepath: str) -> str:
    """画像ファイルのMD5ハッシュを取得"""
    with open(filepath, 'rb') as f:
        return hashlib.md5(f.read()).hexdigest()


def is_same_page(path1: str, path2: str) -> bool:
    """2つの画像が同じかどうか判定"""
    return get_image_hash(path1) == get_image_hash(path2)


# =============================================================================
# 画像処理
# =============================================================================

def process_images(folder: Path) -> None:
    """
    画像の後処理（トリミング・最適化）

    Args:
        folder: 画像フォルダ
    """
    from PIL import Image

    print("\n画像を処理中...")

    image_files = sorted(folder.glob("page_*.png"))
    total = len(image_files)

    for i, img_path in enumerate(image_files, 1):
        print(f"  処理中: {i}/{total}", end="\r")

        img = Image.open(img_path)

        # Kindleのツールバー領域を除去（上下各50px程度）
        width, height = img.size
        top_crop = 50
        bottom_crop = 50
        img = img.crop((0, top_crop, width, height - bottom_crop))

        # 余白をトリミング（白い領域を検出）
        img = trim_whitespace(img)

        # 最適化して保存
        img.save(img_path, "PNG", optimize=True)

    print(f"  処理完了: {total}枚")


def trim_whitespace(img):
    """
    画像の余白（白い領域）をトリミング

    Args:
        img: PIL Image オブジェクト

    Returns:
        トリミングされた画像
    """
    from PIL import Image

    # グレースケールに変換して白い領域を検出
    gray = img.convert('L')
    bbox = gray.getbbox()

    if bbox:
        # 少しマージンを追加
        margin = 10
        x1 = max(0, bbox[0] - margin)
        y1 = max(0, bbox[1] - margin)
        x2 = min(img.width, bbox[2] + margin)
        y2 = min(img.height, bbox[3] + margin)
        return img.crop((x1, y1, x2, y2))

    return img


# =============================================================================
# PDF変換
# =============================================================================

def images_to_pdf(folder: Path, output_pdf: str) -> bool:
    """
    画像ファイルをPDFに結合

    Args:
        folder: 画像フォルダ
        output_pdf: 出力PDFパス

    Returns:
        成功したらTrue
    """
    from PIL import Image

    image_files = sorted(folder.glob("page_*.png"))

    if not image_files:
        print("エラー: 画像ファイルが見つかりません")
        return False

    print(f"\n{len(image_files)}枚の画像をPDFに変換中...")

    images = []
    first_image = None

    for img_path in image_files:
        img = Image.open(img_path)

        # RGBに変換
        if img.mode == 'RGBA':
            background = Image.new('RGB', img.size, (255, 255, 255))
            background.paste(img, mask=img.split()[3])
            img = background
        elif img.mode != 'RGB':
            img = img.convert('RGB')

        if first_image is None:
            first_image = img
        else:
            images.append(img)

    # PDFとして保存
    first_image.save(
        output_pdf,
        "PDF",
        resolution=150.0,
        save_all=True,
        append_images=images
    )

    return True


# =============================================================================
# メイン処理
# =============================================================================

def run_capture(book_name: str, wait_time: float = 0.5, max_pages: int = 2000, test_mode: bool = False) -> Optional[str]:
    """
    スクリーンショット撮影からPDF生成までを実行

    Args:
        book_name: 書籍名
        wait_time: ページめくり間隔（秒）
        max_pages: 最大ページ数
        test_mode: テストモード（5ページのみ撮影して検証）

    Returns:
        生成されたPDFのパス、失敗した場合はNone
    """
    if test_mode:
        max_pages = 5
        print("\n【テストモード】5ページのみ撮影して動作確認します")

    # 出力フォルダ作成
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_folder = Path(f"screenshots_{timestamp}")
    output_folder.mkdir(exist_ok=True)

    # ウィンドウ情報取得
    bounds = get_kindle_window_bounds()
    if not bounds:
        print("エラー: Kindleウィンドウの情報を取得できません")
        return None

    print(f"\n設定:")
    print(f"  書籍名: {book_name}")
    print(f"  ウィンドウ: {bounds[2]}x{bounds[3]} @ ({bounds[0]}, {bounds[1]})")
    print(f"  待機時間: {wait_time}秒")
    print(f"  ページめくり: 矢印キー方式")
    print()

    # Kindleを最前面に
    print("Kindleアプリをアクティブにしています...")
    activate_kindle()
    time.sleep(0.5)

    print("\n撮影を開始します...")
    print("※撮影中はKindleウィンドウを操作しないでください")
    print("※同じページが5回連続で検出されたら自動終了します")
    print("※停止するには Ctrl+C を押してください")
    print()

    # カウントダウン
    for i in range(3, 0, -1):
        print(f"  {i}...")
        time.sleep(1)

    print()

    # 撮影ループ
    last_hash = None
    same_count = 0
    page = 0
    failed_turns = 0  # ページめくり失敗カウント
    start_time = time.time()

    try:
        while page < max_pages:
            page += 1
            output_path = output_folder / f"page_{page:04d}.png"

            # Kindleを最前面に維持
            activate_kindle()
            time.sleep(0.1)

            # ウィンドウ位置を再取得
            current_bounds = get_kindle_window_bounds()
            if current_bounds:
                bounds = current_bounds

            # スクリーンショット撮影
            if not capture_window(str(output_path), bounds):
                print(f"\n警告: ページ {page} の撮影に失敗")
                continue

            # 重複チェック（ページがめくれたか検証）
            current_hash = get_image_hash(str(output_path))

            if current_hash == last_hash:
                same_count += 1
                output_path.unlink()  # 重複画像を削除
                page -= 1

                if same_count >= 5:
                    elapsed = time.time() - start_time
                    print(f"\n\n最後のページに到達しました")
                    print(f"  撮影ページ数: {page}")
                    print(f"  所要時間: {elapsed:.1f}秒")
                    break

                # ページめくりをリトライ
                failed_turns += 1
                if failed_turns >= 10:
                    print(f"\n警告: ページめくりが連続失敗しています")
            else:
                same_count = 0
                failed_turns = 0
                last_hash = current_hash

                # 進捗表示
                elapsed = time.time() - start_time
                rate = page / elapsed if elapsed > 0 else 0
                status = f"  ページ {page} 撮影完了 ({rate:.1f} ページ/秒)"
                print(status, end="\r")

            # ページめくり（矢印キー方式）
            time.sleep(wait_time)
            next_page()
            time.sleep(wait_time)

    except KeyboardInterrupt:
        print(f"\n\n中断されました（{page}ページまで撮影済み）")

    print(f"\n\n全{page}ページの撮影が完了しました！")

    # テストモードの検証結果
    if test_mode:
        print("\n【テスト結果】")
        captured_files = list(output_folder.glob("page_*.png"))
        print(f"  撮影成功: {len(captured_files)}ページ")
        if len(captured_files) >= 3:
            print("  ✓ ページめくりは正常に動作しています")
        else:
            print("  ✗ ページめくりに問題がある可能性があります")
            # テストモードでは一時ファイルを削除
            shutil.rmtree(output_folder)
            return None

    # 画像処理
    process_images(output_folder)

    # PDF変換
    safe_name = "".join(c for c in book_name if c.isalnum() or c in ' -_').strip()
    if not safe_name:
        safe_name = "kindle_book"
    output_pdf = f"{safe_name}_{timestamp}.pdf"

    if images_to_pdf(output_folder, output_pdf):
        print(f"\nPDFを保存しました: {output_pdf}")

        # クリーンアップ
        shutil.rmtree(output_folder)
        print("一時ファイルを削除しました")

        return output_pdf

    print("\nPDF変換に失敗しました")
    print(f"画像ファイルは {output_folder} に保存されています")
    return None


def run_capture_auto(book_name: str, wait_time: float = 0.8, max_pages: int = 2000) -> Optional[str]:
    """
    書籍を自動で検索・オープンしてPDF化

    Args:
        book_name: 書籍名
        wait_time: ページめくり間隔（秒）
        max_pages: 最大ページ数

    Returns:
        生成されたPDFのパス、失敗した場合はNone
    """
    print(f"\n{'='*60}")
    print(f"  処理開始: {book_name}")
    print(f"{'='*60}")

    # Kindleが起動していなければ起動
    if not check_kindle_running():
        if not launch_kindle():
            print("エラー: Kindleを起動できませんでした")
            return None
        time.sleep(2)

    # Kindleをアクティブに
    activate_kindle()
    time.sleep(0.5)

    # 書籍を検索して開く
    if not search_and_open_book(book_name):
        print(f"エラー: 書籍を開けませんでした: {book_name}")
        return None

    # 書籍が開くのを待つ
    time.sleep(3)

    # Kindleをアクティブにして画面分割を再適用
    activate_kindle()
    time.sleep(0.5)
    setup_split_screen()
    time.sleep(1)

    # ウィンドウ情報を確認（複数回リトライ）
    bounds = None
    for retry in range(5):
        activate_kindle()
        time.sleep(0.5)
        bounds = get_kindle_window_bounds()
        if bounds and bounds[3] > 100:  # 高さ100px以上のウィンドウ
            break
        print(f"  ウィンドウ情報取得リトライ... ({retry + 1}/5)")
        time.sleep(2)

    if not bounds:
        print("エラー: Kindleウィンドウの情報を取得できません")
        return None

    print(f"  ウィンドウ確認: {bounds[2]}x{bounds[3]}")

    # 最初のページに移動
    go_to_first_page()
    time.sleep(1)

    # 撮影実行
    result = run_capture(book_name, wait_time, max_pages)

    return result


def process_multiple_books(book_names: list[str], wait_time: float = 0.8) -> dict[str, Optional[str]]:
    """
    複数の書籍を連続でPDF化

    Args:
        book_names: 書籍名のリスト
        wait_time: ページめくり間隔（秒）

    Returns:
        {書籍名: PDFパス} の辞書
    """
    results = {}

    print("\n" + "=" * 60)
    print(f"  {len(book_names)}冊の書籍をPDF化します")
    print("=" * 60)

    # Kindleを起動
    print("\n  Kindleアプリを起動中...")
    if not launch_kindle():
        print("エラー: Kindleを起動できませんでした")
        for book_name in book_names:
            results[book_name] = None
        return results

    # 起動後の待機（ウィンドウ完全初期化待ち）
    time.sleep(3)

    # 画面を左右分割
    setup_split_screen()
    time.sleep(1)

    # ウィンドウが正常に取得できるか確認
    for retry in range(5):
        bounds = get_kindle_window_bounds()
        if bounds:
            print(f"  ✓ Kindleウィンドウ確認: {bounds[2]}x{bounds[3]}")
            break
        print(f"  ウィンドウ取得待機中... ({retry + 1}/5)")
        activate_kindle()
        time.sleep(2)
    else:
        print("エラー: Kindleウィンドウを取得できませんでした")
        for book_name in book_names:
            results[book_name] = None
        return results

    for i, book_name in enumerate(book_names, 1):
        print(f"\n[{i}/{len(book_names)}] {book_name}")
        result = run_capture_auto(book_name, wait_time)
        results[book_name] = result

        if result:
            print(f"  ✓ 完了: {result}")
        else:
            print(f"  ✗ 失敗")

        # 次の書籍処理前に少し待機
        if i < len(book_names):
            time.sleep(2)

    # 結果サマリー
    print("\n" + "=" * 60)
    print("  処理結果サマリー")
    print("=" * 60)

    success_count = sum(1 for v in results.values() if v)
    print(f"  成功: {success_count}/{len(book_names)}")

    for book_name, pdf_path in results.items():
        status = "✓" if pdf_path else "✗"
        print(f"  {status} {book_name}")
        if pdf_path:
            print(f"      → {pdf_path}")

    print("\n処理完了")
    return results


def main():
    """メイン関数"""
    import argparse

    parser = argparse.ArgumentParser(
        description='Kindle書籍をPDFに変換',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  # テストモード（5ページのみ撮影して動作確認）
  python kindle_to_pdf.py --test

  # 単一書籍（手動で開く）
  python kindle_to_pdf.py -n "ネットワークはなぜつながるのか"

  # 単一書籍（自動で検索・オープン）
  python kindle_to_pdf.py -n "書籍名" --auto

  # 複数書籍（自動で連続処理）
  python kindle_to_pdf.py --books "書籍1" "書籍2" "書籍3"

停止方法:
  Ctrl+C でいつでも中断できます
        """
    )
    parser.add_argument('-n', '--name', help='書籍名')
    parser.add_argument('-w', '--wait', type=float, default=0.8, help='ページめくり間隔（秒）')
    parser.add_argument('-m', '--max-pages', type=int, default=2000, help='最大ページ数')
    parser.add_argument('--auto', action='store_true', help='自動で書籍を検索・オープン')
    parser.add_argument('--books', nargs='+', help='複数書籍を連続処理')
    parser.add_argument('--test', action='store_true', help='テストモード（5ページのみ撮影して動作確認）')

    args = parser.parse_args()

    print("=" * 60)
    print("  Kindle PDF変換ツール")
    print("=" * 60)
    print()

    # 1. Kindleを起動（起動していなければ）
    print("準備中...")
    if not check_kindle_running():
        print("  Kindleアプリを起動中...")
        if not launch_kindle():
            print("エラー: Kindleを起動できませんでした")
            sys.exit(1)
        time.sleep(2)
    print("  ✓ Kindleアプリ起動確認")

    # 2. 画面を左右分割（ターミナル左、Kindle右）
    setup_split_screen()
    time.sleep(1)

    # 3. ウィンドウ確認（分割後に再確認）
    for retry in range(5):
        if check_kindle_window_exists():
            break
        print(f"  ウィンドウ待機中... ({retry + 1}/5)")
        activate_kindle()
        time.sleep(1)
    else:
        print("エラー: Kindleウィンドウが見つかりません")
        sys.exit(1)
    print("  ✓ Kindleウィンドウ確認")

    bounds = get_kindle_window_bounds()
    if not bounds:
        print("エラー: ウィンドウ情報の取得に失敗しました")
        sys.exit(1)
    print(f"  ✓ ウィンドウサイズ: {bounds[2]}x{bounds[3]}")

    # テストモード
    if args.test:
        print()
        print("【テストモード】5ページのみ撮影して動作確認します")
        print("  Kindleで任意の書籍を開いた状態でEnterを押してください...")
        input()
        result = run_capture("test", args.wait, test_mode=True)
        if result:
            print("\n✓ テスト成功！本番実行できます")
        else:
            print("\n✗ テスト失敗。設定を確認してください")
        sys.exit(0 if result else 1)

    # 複数書籍モード
    if args.books:
        results = process_multiple_books(args.books, args.wait)
        success = all(v for v in results.values())
        sys.exit(0 if success else 1)

    # 単一書籍モード
    book_name = args.name or 'kindle_book'

    if args.auto:
        # 自動モード
        result = run_capture_auto(book_name, args.wait, args.max_pages)
    else:
        # 手動モード
        print()
        print("【重要】アクセシビリティ権限が必要です")
        print("  システム設定 > プライバシーとセキュリティ > アクセシビリティ")
        print("  でターミナルを許可してください")
        print()

        input("Kindleで書籍を最初のページに開いてからEnterを押してください...")
        result = run_capture(book_name, args.wait, args.max_pages)

    if result:
        print()
        print("=" * 60)
        print(f"  完了！生成されたPDF: {result}")
        print("=" * 60)
    else:
        print()
        print("PDF生成に失敗しました")
        sys.exit(1)


if __name__ == "__main__":
    main()
