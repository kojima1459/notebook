# mybookshelf/build/ — 「マイ本棚AI」ビルド機構

このディレクトリは `mybookshelf/dist/MyBookshelf.xlsm`(または開発版
`MyBookshelf_dev.xlsm`)を生成するビルド機構一式です。**自己完結**しており、
`/home/user/notebook/build/`(V2チャットボットの資産)には一切依存しません
(コードは同一リポジトリ内から実証済み機構をコピーしていますが、実行時に
そのファイルを読み込むことはありません)。

## ファイル構成

| ファイル | 役割 |
|---|---|
| `build_mybookshelf.py` | ビルド本体。これを実行する |
| `ovba.py` | MS-OVBA圧縮/解凍・CFB(vbaProject.bin)最小限リーダー。純粋なバイト列操作のみ、他への依存なし |
| `modules.json` | VBAモジュール台帳(MASTER_SPEC §14準拠)。`{"name","path","role":"core\|opt\|test"}` |
| `template_skeleton.xlsm` | 本物のExcelで作られた.xlsmスケルトン(`/home/user/notebook/build/template_skeleton.xlsm` のコピー)。中の`vbaProject.bin`は本物なのでExcelに拒否されない |
| `README.md` | このファイル |

## 使い方

```bash
cd mybookshelf/build
pip install openpyxl olefile   # 未導入なら

# 開発ビルド (config mock_llm=TRUE。リボン無しでもmockで一通り動く)
python3 build_mybookshelf.py --dev

# 本番ビルド (config mock_llm=FALSE)
python3 build_mybookshelf.py --prod
```

既定の出力先: `mybookshelf/dist/MyBookshelf.xlsm`(`--prod`)/
`mybookshelf/dist/MyBookshelf_dev.xlsm`(`--dev`)。

### 主なオプション

- `--dev` / `--prod` — どちらか必須。config `mock_llm` の初期値を切り替える。
  どちらのビルドにも `role: test` のモジュール(テストランナー等)は同梱される
  (診断用。撤去したければ `modules.json` から該当行を削除する)。
- `--allow-missing` — `modules.json` に列挙されているのに実在しないファイルが
  あってもエラーで止めず、警告して該当モジュールを `vba_src` から除外して
  続行する。Wave2以前の段階でビルド機構だけを検証したいときに使う
  (本リポジトリのWave1時点ではこれが前提)。
- `--modules PATH` — `modules.json` の差し替え(既定: このディレクトリの
  `modules.json`)。
- `--root PATH` — `modules.json` 内 `path` の解決基準ディレクトリ(既定:
  `mybookshelf/`)。スモークテスト等で本物のソースツリーに触れずに検証したい
  時に使う。
- `--template PATH` — `template_skeleton.xlsm` の差し替え。
- `--out PATH` — 出力先.xlsmパスの上書き。

## opt機能の撤去

MASTER_SPEC §2/§7.7 の設計により、opt機能(`optTts`/`optVision`/
`optMarkdown`/`optDiffDoc`)はコアから `modFeatures.InvokeFeature` 経由の
遅延バインドでしか呼ばれません。**撤去は `modules.json` から該当行を1行
削除するだけ**でよく、コア機能は無傷のまま動きます(config の
`feature_<id>` フラグをFALSEにしておけば、モジュール自体が無くても
`modFeatures.ModulePresent` が False を返すため実行時エラーにはなりません)。

## 実装している機構(V2実証済みパターンの自己完結コピー)

1. `template_skeleton.xlsm` を `openpyxl(keep_vba=True)` で読み込み、
   MASTER_SPEC §4 の全13シートを生成する。この時点では `vbaProject.bin` は
   スケルトンのバイト列のまま一切変更しない。
2. `modules.json` に列挙され実在する標準モジュール(`*.bas`)のソースを
   `vba_src` シート(veryHidden)に1行1モジュールで格納する。
   **クラスモジュール(`*.cls`、例: `ThisWorkbook.cls`)は
   `VBComponents.Add(1)` では追加できない**(標準モジュール専用API)ため
   `vba_src` の対象外(`"type": "class"` または `"vba_src": false` を
   `modules.json` に指定)。
3. openpyxlで保存(`vbaProject.bin` はまだスケルトンのまま)。
4. 保存済み `.xlsm` から `vbaProject.bin` を取り出し、**外科的パッチ**を
   当てる:
   - `VBA/ThisWorkbook` ストリームを「自己インストーラ」のソースに差し替え。
     `Workbook_Open` → `vba_src` シートを読んで
     `VBProject.VBComponents.Add(1)` で標準モジュールを全部インストール →
     `Application.Run "modBoot.Boot"` を呼ぶ。
     インストーラのソース文字列は **ASCII のみ**(cp932エンコードだが
     英字に限定。MASTER_SPEC §14手順6)。
   - `VBA/dir` ストリームの `ThisWorkbook` の `MOFFSET` を `0` に書き換え、
     パフォーマンスキャッシュを無効化してバイト0からソース解凍させる。
   - どちらのストリームも書き換え後にOVBAの「空チャンク」でpaddingし、
     元のストリームとバイト長を完全一致させる(`olefile.write_stream` は
     ストリームサイズを変更できないため)。CFBヘッダ・FAT・
     `_VBA_PROJECT`・`PROJECT`/`PROJECTwm`・他の全ストリームは触れない。
   - 補足: 5バイトpadding変形チャンクは解凍すると実際には2バイトのゼロ値
     になる(3バイト変形は0バイト)。結果としてThisWorkbookストリームの
     解凍結果は「インストーラソース + 末尾に数バイトのNULパディング」に
     なるが、VBAコンパイラは末尾のNULを無視するため実害はない
     (`oletools.olevba` での独立検証でも問題なく抽出できることを確認済み)。
5. ビルド後、生成物を再オープンして自己検証する(次節)。

## ビルド後自己検証 (MASTER_SPEC §10)

`build_mybookshelf.py` は生成した `.xlsm` を書き出した直後に、同じ
プロセス内で以下を検証し、1つでも失敗すれば **exit code 1** で終了します
(dist に壊れたファイルを残したまま成功したように見せない):

1. `openpyxl(keep_vba=True)` で再オープンし、MASTER_SPEC §4 の全13シートが
   存在し、可視性(`visible`/`hidden`/`veryHidden`)が仕様通りであること。
2. `vba_src` シートの行数(モジュール数)が、実際に注入したモジュール集合と
   一致すること。
3. `vba_src` の各ソースセルが Excel のセル文字数上限(32,000字)を超えて
   いないこと(モジュール契約上限30,000字は注入前にも別途検査)。
4. `config` シートの `mock_llm` が `--dev`/`--prod` の期待値と一致し、
   キー数がMASTER_SPEC §5の29キーと一致すること。
5. `olefile` で `vbaProject.bin` の `VBA/ThisWorkbook` ストリームを
   読み出し・解凍し、自己インストーラのソースとして復元できること
   (dirストリームの `MOFFSET` が `0` になっていることも含む)。

## スモークテスト

Wave2(機能モジュール実装)が完了する前でも、ビルド機構そのものが正しいか
を検証できます。実ソース(`mybookshelf/src/`)や `dist/` には一切触れず、
`--root`/`--modules`/`--out` を一時ディレクトリに向けて実行します:

```bash
SCRATCH=/path/to/scratch
mkdir -p "$SCRATCH/smoke_root/src/ui"
cat > "$SCRATCH/smoke_root/src/ui/modBoot.bas" <<'EOF'
Attribute VB_Name = "modBoot"
Option Explicit
Public Sub Boot()
    MsgBox "Hello", vbInformation
End Sub
EOF
cat > "$SCRATCH/smoke_modules.json" <<'EOF'
{"modules": [{"name": "modBoot", "path": "src/ui/modBoot.bas", "role": "core", "type": "std"}]}
EOF

python3 build_mybookshelf.py --dev \
  --root "$SCRATCH/smoke_root" \
  --modules "$SCRATCH/smoke_modules.json" \
  --out "$SCRATCH/smoke_dist/Smoke_dev.xlsm"
```

`自己検証 OK` と表示され exit code 0 なら、ビルド機構(シート生成・
自己インストーラ注入・パッチ・自己検証)は健全です。実ファイルが揃った
Wave2完了後は `--allow-missing` なしの通常ビルド(`--dev`/`--prod` のみ)で
全モジュールの存在を強制検査してください。

## 注意事項

- `/home/user/notebook/build/`, `/home/user/notebook/src/`,
  `/home/user/notebook/docs/`(V2チャットボットの資産)は参照・コピー元
  であり、このビルド機構からは一切書き込みません。
- `template_skeleton.xlsm` は `/home/user/notebook/build/template_skeleton.xlsm`
  のバイト同一コピーです。テンプレート自体を更新したい場合は、
  元ファイルではなくこのコピーを直接差し替えてください(V2側は変更禁止)。
- VBAプロジェクトのパスワード保護(V2の `build_chatbot_v2.py` Stage 6相当)
  はこのスクリプトには未実装です。必要になった場合はWindows+pywin32環境で
  別途追加検討してください。
