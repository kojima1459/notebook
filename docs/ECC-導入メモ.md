# ECC 導入メモ（自分用の手順書）

ECC = エンジニア向けの大規模なClaude Code強化ツールキット。
エージェント64個・スキル262個・コマンド84個が入る、かなり本格的なもの。
公式：https://github.com/affaan-m/ECC

> ⚠️ 注意：これは「プロのエンジニアが性能を限界まで上げる」ための道具箱です。
> 初心者にはオーバースペック気味なので、**まずは入れてみて、合わなければ消す**くらいの気持ちでOK。

---

## なぜAI（Claude）が代わりに入れてくれないの？

インストールは `/plugin ...` という**スラッシュコマンド**で行います。
これはあなた（人間）がClaude Codeの入力欄に打ち込む専用コマンドで、
AIエージェントの側からは実行できない仕組みになっています。
なので**下の手順は自分の手で**実行してください。

また、Web版のClaude Code（ブラウザ）はセッションごとに環境がリセットされるため、
**入れるなら自分のパソコンにインストールしたClaude Code（ローカル版）で行うのがおすすめ**です。
そうすれば一度入れれば、毎回使えます。

---

## 事前準備

1. ターミナル（黒い画面）で、Claude Codeのバージョンを確認：
   ```bash
   claude --version
   ```
2. **v2.1.0 以上**であること。古ければアップデートする。

---

## インストール手順（Claude Codeの入力欄で実行）

```text
/plugin marketplace add https://github.com/affaan-m/ECC
/plugin install ecc@ecc
```

- 1行目：ECCの配布元（マーケットプレイス）を登録する
- 2行目：ECC本体をインストールする

---

## 仕上げ（ルールの追加）※任意

プラグインだけだと「ルール」が入らないので、必要なら手動でコピーする。
ターミナルで：

```bash
mkdir -p ~/.claude/rules/ecc
cp -r rules/common ~/.claude/rules/ecc/
cp -r rules/typescript ~/.claude/rules/ecc/
```

> ❗ `./install.sh --profile full` は**実行しないこと**。
> プラグイン版と二重インストールになって壊れる原因になります。

---

## 入ったか確認

```text
/plugin list ecc@ecc
```

エージェント・スキル・コマンドの一覧が出れば成功。

---

## 合わなかったら（アンインストール）

```text
/plugin uninstall ecc@ecc
```

気軽に消せるので、まずは試してみて大丈夫です。
