---
name: bssc-pptx-rules
description: BSSC社の PPTX（PowerPoint）作成における社内共通ルール。スライドサイズ（27.517cm×19.05cm）、フォント（ページ番号を除き全てメイリオ）、文字サイズ（原則11pt以上／図形内・表のセル内は6ptまで許容、6pt未満は不可）、著作権表示「© BSSC Co., Ltd.」（太字・表紙を含む全スライド）の依頼ごとの事前確認と配置、表紙を除いたページ番号（MS Pゴシック16pt太字・右下、スライドの追加・削除・並べ替えに自動連動するスライド番号フィールド方式）を規定する。ユーザーが .pptx / スライド / プレゼン資料 / 提案書 / 営業資料 / 研修資料 / 講義資料 / セミナー資料 などを「作成」「出力」「編集」する依頼をしたときは、内容やテーマを問わず必ずこのスキルを参照すること。ユーザーが明示的にルールに言及しなくても適用する。既存 pptx の修正・追記・差し替え・削除・並べ替え・体裁調整でも必ず参照する。
---

# BSSC PPTX 共通ルール

BSSC Co., Ltd. が配布・使用する資料の体裁を統一するためのルール。PPTX を新規作成する場合も、既存 PPTX を編集する場合も、最終的な成果物がこのルールを満たしている必要がある。

作成手順そのもの（pptxgenjs の使い方、QA、画像化など）は公開スキル `pptx`（`/mnt/skills/public/pptx/SKILL.md`）に従う。本スキルはその上に重ねる BSSC 固有の体裁要件を定める。

## 0. 作成前の必須確認（省略不可）

PPTX の作成・編集に着手する前に、ほかの作業より先に、必ずユーザーへ次を確認する。

> すべてのスライドに「© BSSC Co., Ltd.」を掲載しますか？

- **作成または編集の依頼ごとに、毎回確認する。** 同一会話内で以前に確認済みであっても、新しい資料の作成依頼や別ファイルの編集依頼を受けたときは、あらためて確認する。
- 過去の資料、過去の回答、メモリの内容から掲載可否を推測しない。
- 著作権表示は**原則として掲載する**。ユーザーが「不要」「入れない」と回答した場合のみ掲載しない。
- 確認を省略して着手しない。

なお、同一の資料に対する連続した微修正（「この2ページ目の文言だけ直して」等）であれば、直前の回答をそのまま引き継いでよい。

## 1. スライドサイズ

- **幅 27.517 cm × 高さ 19.05 cm** を明示的に設定する。
- これは PowerPoint の「A4 210×297mm・横」に相当する。用紙実寸の A4（29.7×21.0cm）ではないので、「A4横だから 29.7×21.0」と解釈しない。
- 16:9（既定の `LAYOUT_16x9` や `LAYOUT_WIDE`）にはしない。

pptxgenjs では、スライドを 1 枚も追加する前にレイアウトを定義して適用する（インチ換算：27.517cm = 10.8335in、19.05cm = 7.5in）。

```javascript
const pres = new pptxgen();
pres.defineLayout({ name: 'A4_LANDSCAPE', width: 10.8335, height: 7.5 });
pres.layout = 'A4_LANDSCAPE';   // ← addSlide() より前に設定する
```

出力後、EMU 単位で厳密に一致させる（27.517cm = 9,906,120 EMU、19.05cm = 6,858,000 EMU）。インチ換算の丸め誤差を消すため、この後処理は毎回行う。

```python
from pptx import Presentation
from pptx.util import Cm

prs = Presentation("output.pptx")
prs.slide_width  = Cm(27.517)
prs.slide_height = Cm(19.05)
prs.save("output.pptx")
```

## 2. フォントと文字サイズ

### 2-1. フォント

- 資料内のすべての文字は **メイリオ** とする。表紙・見出し・本文・箇条書き・表・図形内の文字・グラフ内の文字・注釈・出典を含む。
- **例外はページ番号のみ**。ページ番号は **MS Pゴシック** とする（後述 3.）。
- 著作権表示 `© BSSC Co., Ltd.` もメイリオとする。
- pptxgenjs の既定フォント（Arial 等）に任せず、`fontFace: 'メイリオ'` を毎回明示する。テーマや `defineSlideMaster` を使う場合もマスター側で既定を上書きする。

```javascript
const F = 'メイリオ';
slide.addText('見出し', { fontFace: F, fontSize: 24, bold: true, /* … */ });
slide.addText('本文',   { fontFace: F, fontSize: 14, /* … */ });
```

表（`addTable`）や図形（`addShape` + `addText`）は指定漏れが起きやすいので、セル・図形の一つひとつに `fontFace` を渡す。

### 2-2. 文字サイズ

- **原則として 11pt 未満の文字を作らない。** 10pt・9pt・8pt 等を指定した箇所はすべて 11pt に引き上げる。見出し・本文・箇条書き・注釈・脚注・出典が対象。
- **緩和されるのは「図形内の文字」と「表のセル内の文字」のみ。** ◯・矢印・四角などの図形（オートシェイプ）の中に入れる文字、および表（テーブル）のセル内の文字は、収まりを優先して 11pt 未満にしてよい。表のヘッダー行・データ行いずれも対象。
- **ただし、いかなる場合も 6pt 未満にはしない。** 図形内・表内であっても下限は 6pt。6pt でも収まらない場合は、文字量を減らす、図形や列幅を広げる、要素を分割するなどで対応する。
- 「収まらないから小さくする」は、図形内・表内以外では認めない。文字量を減らす、領域を広げる、行間を詰める、スライドを分けるなどで対応する。
- なお、フォント（メイリオ）は図形内・表内も含めて例外なく適用する。緩和されるのはサイズのみ。

| 対象 | 下限 |
|---|---|
| 見出し・本文・箇条書き・注釈・脚注・出典 | 11pt |
| 図形（オートシェイプ）内の文字 | 6pt |
| 表のセル内の文字 | 6pt |

### 2-3. 出力後の一括是正

指定漏れが残ることがあるため、出力後に python-pptx で機械的に是正する。フォントは全要素に適用し、11pt への引き上げは図形内・表内を除いた通常のテキストボックス／プレースホルダに、6pt への引き上げは図形内・表内に行う。

**ページ番号（3. のスクリプト）は、この是正処理を実行したあとに付与する。** ページ番号はスライド番号フィールド（`a:fld`）で入るため下記の run 処理の対象外だが、旧来の固定数字が残っている場合に備えて数字のみの run は除外している。

```python
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE

prs = Presentation("output.pptx")

def fix(tf, floor_pt):    # floor_pt: 11（通常）または 6（図形内・表内）
    for p in tf.paragraphs:
        for r in p.runs:
            if r.text.strip().isdigit():   # ページ番号は MS Pゴシックのまま
                continue
            r.font.name = "メイリオ"
            if r.font.size is not None and r.font.size < Pt(floor_pt):
                r.font.size = Pt(floor_pt)

for s in prs.slides:
    for sh in s.shapes:
        if sh.has_text_frame:
            is_shape = sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
            fix(sh.text_frame, floor_pt=6 if is_shape else 11)
        if sh.has_table:                    # 表のセルは下限 6pt
            for row in sh.table.rows:
                for c in row.cells:
                    fix(c.text_frame, floor_pt=6)

prs.save("output.pptx")
```

## 3. ページ番号（自動連動方式）

### 3-1. 表示ルール

- 表紙はスライド 1 とし、**表紙にはページ番号を表示しない**。表紙はページ数にも含めない。
- スライド 2 を「1 ページ目」として右下に `1` と表示する。以降、スライド 3 = `2`、スライド 4 = `3` と連番。
- 表示する番号は常に「実際のスライド番号 − 1」。
- 数字のみを表示する（「P.1」「1/10」等にしない）。

書式は全スライドで統一する。

| 項目 | 指定 |
|---|---|
| 位置 | 各スライドの右下（x=10.10in, y=6.95in, w=0.55in, h=0.35in、右揃え） |
| フォント | MS Pゴシック |
| サイズ | 16pt |
| スタイル | 太字 |

### 3-2. 自動連動の仕組み（必須）

**ページ番号は、スライドの追加・削除・並べ替えに自動で連動しなければならない。** 数字を固定テキストとして打ち込む方式（`addText(String(i))` など）は禁止する。次の 3 点で実現する。

| 設定 | 内容 | 効果 |
|---|---|---|
| ① スライド番号フィールド | 全スライドレイアウトに `<a:fld type="slidenum">` を含むテキストボックス（名前 `BSSC_PageNumber`）を 1 つ置く | PowerPoint が表示時に番号を自動計算。新規追加したスライドにも自動で番号が出る |
| ② 開始番号 0 | `presentation.xml` の `<p:presentation firstSlideNum="0">`（PowerPoint の［スライドのサイズ］→［ユーザー設定］→「スライド開始番号」＝0 に相当） | 表紙＝0、スライド 2＝1 となり「スライド番号 − 1」が自動で成立する |
| ③ 表紙で非表示 | 表紙スライドに `showMasterSp="0"`（［背景の書式設定］→「背景グラフィックを表示しない」に相当） | 表紙だけ番号（0）を出さない |

この結果、PowerPoint 上でユーザーがスライドを挿入・削除・複製・並べ替えしても、全ページの番号が即座に振り直される。

### 3-3. 付与手順

pptxgenjs 側ではページ番号を付けない（`slideNumber` オプションも使わない）。出力後、2-3. のフォント是正を実行したあとに、このスキル同梱の `scripts/bssc_page_numbers.py` を実行する。

```bash
python3 <このスキルのディレクトリ>/scripts/bssc_page_numbers.py output.pptx
```

スクリプトの処理内容：

1. `firstSlideNum="0"` を設定する。
2. 全レイアウトの `BSSC_PageNumber` を置き直す（再実行しても重複しない＝冪等）。
3. 各スライド上の旧ページ番号（固定数字の箱、`sldNum` プレースホルダ、右下の数字のみのテキスト）を削除し、二重表示を防ぐ。
4. 表紙に `showMasterSp="0"` を設定する。表紙のレイアウトにロゴ等の装飾図形がある場合は「注意」を表示するので、その場合は表紙の見た目を目視確認し、必要な装飾は表紙スライド上に直接配置する。

### 3-4. 既存 PPTX の編集時

- スライドの追加・削除・並べ替えを行った場合も、最後に必ず `bssc_page_numbers.py` を実行する（固定数字で番号が入っている既存資料は、このとき自動連動方式に置き換わる）。
- 表紙を先頭から移動しない。中表紙などで番号を消したいスライドがあれば、そのスライドにも `showMasterSp="0"` を設定する（番号は非表示になるが、ページ数には含まれる）。

### 3-5. 納品時にユーザーへ伝えること

資料を渡すときは、次の 1 文を添える。

> ページ番号は自動連動です。PowerPoint でスライドを追加・削除・並べ替えしても番号は自動で振り直されます（［スライドのサイズ］の「スライド開始番号」は 0 のまま変更しないでください）。

## 4. 著作権表示（0. の確認でユーザーが掲載を希望した場合）

- 表示内容は必ず `© BSSC Co., Ltd.` とする。会社名・記号・ピリオド・カンマ・大文字小文字を変更しない（表記ゆれ不可）。
- **太字**にする（`bold: true`）。
- フォントはメイリオ、サイズは 11pt 以上とする。
- **表紙を含むすべてのスライドに掲載する。**
- 位置・フォント・サイズ・書式を全スライドで統一する。ユーザーから位置の指定がなければ右下に配置する。
- 本文、図表、ページ番号と重ならないようにする。

## 5. 著作権表示とページ番号を併記する場合

- 右下の**最も右側にページ番号**を置く。
- `© BSSC Co., Ltd.` は**ページ番号の左側**に置く。
- 両者が重ならないよう十分な間隔（0.15in 以上を目安）を確保する。
- 全スライドで同じ位置関係を保つ。

配置イメージ（右下）：

```
              … © BSSC Co., Ltd.   │   3
                     ↑             ↑   ↑
                  左側に配置    十分な間隔  最も右
```

なお表紙にはページ番号がないため、著作権表示のみが右下に入る。位置は他のスライドと揃える。

## 6. 適用の優先順位

- ユーザーから個別に異なる書式指定があれば、その指定を優先する。
- 個別指定がない項目については、本ルールを必ず適用する。
- ただし **0. の著作権表示の確認だけは例外なく毎回行う**。個別指定の有無にかかわらず省略しない。

## 7. 出力前チェック（必須）

`pptx` スキルの QA に加えて、出力前に次を確認する。確認できない項目があれば、推測で「対応済み」と報告しない。

- [ ] スライドサイズが 27.517cm × 19.05cm（9,906,120 × 6,858,000 EMU）になっている
- [ ] ページ番号以外のすべての文字がメイリオになっている
- [ ] 図形内・表のセル内を除き、11pt 未満の文字が 1 つも残っていない
- [ ] 図形内・表のセル内も含め、6pt 未満の文字が 1 つも残っていない
- [ ] 表紙にページ番号がない（表紙に `showMasterSp="0"`）
- [ ] `firstSlideNum="0"` が設定され、スライド 2 以降の表示番号が「スライド番号 − 1」になる
- [ ] 全レイアウトに `BSSC_PageNumber`（`slidenum` フィールド、MS Pゴシック・16pt・太字・右下）が 1 つずつある
- [ ] スライド上に固定数字のページ番号が残っていない（自動連動が壊れ、二重表示になるため）
- [ ] 著作権表示の有無が、今回の依頼に対するユーザーの回答どおりである
- [ ] 著作権表示がある場合、表紙を含む全スライドにある
- [ ] 著作権表示がある場合、表記が `© BSSC Co., Ltd.` で太字、ページ番号の左側にあり重なっていない

サイズとページ番号・著作権表示は目視だけでなく機械的に検証する：

```python
from pptx import Presentation
from pptx.oxml.ns import qn
prs = Presentation("output.pptx")
print(prs.slide_width, prs.slide_height)   # 9906120 6858000 であること
first = prs.part._element.get("firstSlideNum")
print("firstSlideNum =", first)            # "0" であること
for m in prs.slide_masters:
    for l in m.slide_layouts:
        n = [s for s in l.shapes if s.name == "BSSC_PageNumber"]
        ok = len(n) == 1 and n[0]._element.find(".//" + qn("a:fld")).get("type") == "slidenum"
        print("レイアウト", l.name, "OK" if ok else "NG")
for i, s in enumerate(prs.slides, start=1):
    hidden = s._element.get("showMasterSp") == "0"
    static = [sh.text_frame.text for sh in s.shapes
              if sh.has_text_frame and sh.text_frame.text.strip().isdigit()]
    copy = [sh.text_frame.text for sh in s.shapes if sh.has_text_frame and "BSSC" in sh.text_frame.text]
    print(i, "番号:", "非表示" if hidden else i - 1, "固定数字:", static, "著作権:", copy)
    # 表紙のみ「非表示」、固定数字は空であること
```

フォントと文字サイズも機械的に洗い出す（下限は図形内・表内が 6pt、それ以外が 11pt）：

```python
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE

prs = Presentation("output.pptx")

def check(tf, i, floor_pt):
    for p in tf.paragraphs:
        for r in p.runs:
            if r.text.strip().isdigit():
                continue
            if r.font.name != "メイリオ":
                print("フォント違反", i, r.font.name, r.text[:20])
            if r.font.size and r.font.size < Pt(floor_pt):
                print(f"サイズ違反(下限{floor_pt}pt)", i, r.font.size.pt, r.text[:20])

for i, s in enumerate(prs.slides, start=1):
    for sh in s.shapes:
        if sh.has_text_frame:
            is_shape = sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
            check(sh.text_frame, i, floor_pt=6 if is_shape else 11)
        if sh.has_table:
            for row in sh.table.rows:
                for c in row.cells:
                    check(c.text_frame, i, floor_pt=6)
```

そのうえで、`pptx` スキルの手順で画像化し、右下の重なりや文字切れを目視確認する。
