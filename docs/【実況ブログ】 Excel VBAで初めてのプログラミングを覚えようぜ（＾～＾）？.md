# 📅2023-01-25 mon 19:15 start

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　Excel を購入しろだぜ」  

![kifuwarabe-futsu.png](https://crieit.now.sh/upload_images/beaf94b260ae2602ca8cf7f5bbc769c261daf8686dbda.png)  
「　してるぜ」  

![202301_excel_25-1920--sheet.png](https://crieit.now.sh/upload_images/7d2e6ce5359b68c2fbfc93128857ccd863d102e696019.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　これを使って、プログラミングを覚えてもらう」  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　現代の Excel は　チューリング完全らしいですしね」  

![kifuwarabe-futsu.png](https://crieit.now.sh/upload_images/beaf94b260ae2602ca8cf7f5bbc769c261daf8686dbda.png)  
「　じゃあ Excel で　チューリング・マシンを作るのか？  
無限のテープをどう表現する？」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　テープが足りなくなったら　強制終了すればいいだろ」  

📖 [Turing machine](https://en.wikipedia.org/wiki/Turing_machine)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　例題は Wikipedia を参考にしようぜ？」  

![202301_excel_25-2002--StateTable.png](https://crieit.now.sh/upload_images/a7bee6a2939b3d27802fa74666966e4e63d10ccb8eaa1.png)  

`[TuringMachineByExcelVBA.xlsm] file - [StateTable] sheet`:  

```csv
State,Read,Write,Move,Transition
A,White,Orange,>,B
A,Orange,Orange,<,C
B,White,Orange,<,A
B,Orange,Orange,>,B
C,White,Orange,<,B
C,Orange,Orange,>,HALT
```

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　とりあえず `StateMachine` シートを作れだぜ。  
セルに色も塗れだぜ」  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　このシートが何かの説明はしないの？」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　手を動かせば　あとで分かる」  

![202301_excel_25-2008--Tape.png](https://crieit.now.sh/upload_images/2bd8e45ab2cf441bd6fbd64b9fb1440863d10dc15985e.png)  

`[TuringMachineByExcelVBA.xlsm] file - [Tape] sheet`:  

```csv
A
```

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　もう１つ、 `Tape` シートを作れだぜ。  
`A` の１文字だけ入っている」  

![kifuwarabe-futsu.png](https://crieit.now.sh/upload_images/beaf94b260ae2602ca8cf7f5bbc769c261daf8686dbda.png)  
「　これは何なんだぜ？」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　あとで分かる」  

![202301_excel_25-2014--OpenCode-1.png](https://crieit.now.sh/upload_images/5be4add46fc7c8f5a3ff33285bc455d163d10f482920b.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　空っぽの `GUI` というシートを作って、  
メインメニューから `[開発] - [コードの表示]` を選べだぜ」  

![202301_excel_25-2017--VBAEditor.png](https://crieit.now.sh/upload_images/8e5c6787ff6884ae25742bf5316c659b63d10fc7a71bd.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　VBA のエディターが出てくるな」  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　あの　不便なやつね」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　ほんと　不便だな。　シートに戻るぜ」  

![202301_excel_25-2021--Button-1.png](https://crieit.now.sh/upload_images/e8e85792d662411ad13e424c8571f5b763d11114d4f20.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　先に　ボタンを置こうぜ？」  

![202301_excel_25-2024--PutButton-1.png](https://crieit.now.sh/upload_images/43e7f5d13408af482856786fd1e9570b63d111bad4d9e.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　何だか　よくわからないが　`ボタン1_Click` という名前はそのまんまで  
マクロの保存先を　今作業中のファイルに変えて、 `[新規作成(N)]` ボタンを押そうぜ？」  

![kifuwarabe-futsu.png](https://crieit.now.sh/upload_images/beaf94b260ae2602ca8cf7f5bbc769c261daf8686dbda.png)  
「　Excel の使い方を記憶してないの　わらう」  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　プログラマーは　記憶　ではなく、　**読み**　で進むのよ。  
その方が　応用が利くから」  

![202301_excel_25-2029--CreateButton-1.png](https://crieit.now.sh/upload_images/7fc9f87fe1c8e6ad26d268f24fc9812f63d112c4a2e30.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　ボタンがでけてる。  
当たった」  

![kifuwarabe-futsu.png](https://crieit.now.sh/upload_images/beaf94b260ae2602ca8cf7f5bbc769c261daf8686dbda.png)  
「　当たった　とか　講師から出てきたらおかしい言葉　わらう」  

![202301_excel_25-2032--RegisterMacro-1.png](https://crieit.now.sh/upload_images/a9682e7b76b5ec11dd96a83993301a8b63d11393ae7ea.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　ボタンを右クリックして　コンテキスト・メニューの `[マクロの登録(N)]` を  
クリックしてみようぜ？」  

![202301_excel_25-2035--Ok-1.png](https://crieit.now.sh/upload_images/ff5836d7acb5395226ed608bce8df05a63d114270be35.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　何をすればいいのか分からん。 `[OK]` ボタンを押してみようぜ？」  

![202301_excel_25-2037--Code.png](https://crieit.now.sh/upload_images/995318528394eec0b95b7eac406e1ace63d1149283db7.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　？  
何が起こった？」  

![202301_excel_25-2038--Skeleton-1.png](https://crieit.now.sh/upload_images/4127845072bc324880940de1cf50593763d114ef2b8ba.png)  

![kifuwarabe-futsu.png](https://crieit.now.sh/upload_images/beaf94b260ae2602ca8cf7f5bbc769c261daf8686dbda.png)  
「　👆　なんか　コードが増えてるのでは？」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　よっしゃ！  
じゃあ　そこに VBA Script （ぶい・びー・えー・すくりぷと）を書けばいいんだぜ」  

📅 2023-01-25 wed 20:41  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　🔍 `VBA セルに値を入れる`　で検索」  

📖 [セルに値を入れる：Excel VBA プログラミング入門](http://www.eurus.dti.ne.jp/~yoneyama/Excel/vba/prog/prog_atai.html)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　よし　分かったぜ」  

![202301_excel_25-2046--HelloWorld-1.png](https://crieit.now.sh/upload_images/8bedf851dc72c6806bd384e10b70508e63d116df021a7.png)  

```vba
Worksheets("GUI").Range("A1").Value = "Hello, world!!"
```

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　こんな感じに書けば　`GUI` シートの `A1` セルに　`Hello, world!!` という文字を  
入れてくれそうだな」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　ボタンを押そうぜ？」  

![202301_excel_25-2049--ShowHelloWorld-1.png](https://crieit.now.sh/upload_images/ea960374c712c2d05121e25ea27ea9cc63d11794cb716.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　ほら　出た」  

![kifuwarabe-futsu.png](https://crieit.now.sh/upload_images/beaf94b260ae2602ca8cf7f5bbc769c261daf8686dbda.png)  
「　`Excelでハローワールドを出力する` の実績を解除したな」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　じゃあ次。  
🔍 `VBA セルに色を付ける` で検索」  

📖 [セルに色を設定する](https://www.tipsfound.com/vba/07006)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　よし　分かったぜ」  

![202301_excel_25-2056--BackgroundColor-1.png](https://crieit.now.sh/upload_images/ab3f6398046ba32df211de0534339eb163d1190f3e6d6.png)  

```vba
Worksheets("GUI").Range("A1").Interior.ColorIndex = 45 ' オレンジ
```

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　こんな感じに書けば　`GUI` シートの `A1` セルの背景色をオレンジ色に  
してくれそうだな」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　ボタンを押そうぜ？」  

![202301_excel_25-2057--ShowBackgroundColor-1.png](https://crieit.now.sh/upload_images/ed751568c624bbe3c65a037eb56fccf663d11972c76a3.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　A1 セルに色が付いたな」  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　着色の自動化に使えそうねえ」  

![202301_excel_25-2108--GetColor-1.png](https://crieit.now.sh/upload_images/d3954444815536359e16f3966743a8c963d11c0290ab2.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　じゃあ、 `StateTable` シートの C2 セルの背景色が何色かとか、取得することはできるのかだぜ？」  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　ググりゃいいんじゃないの？」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　🔍 `VBA セルの色を取得` で検索」  

📖 [VBA セルの色を取得する (Interior.Color, ColorIndex)](https://www.tipsfound.com/vba/07005)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　よし　分かったぜ」  

![202301_excel_25-2117--CopyColor-1.png](https://crieit.now.sh/upload_images/54dd3efed1a31e73d04c267e53e17dac63d11e1182115.png)  

```vba
    Dim backgroundColor As Long
    backgroundColor = Worksheets("StateTable").Range("C2").Interior.color ' 背景色
    Debug.Print (backgroundColor)

    Worksheets("GUI").Range("A1").Interior.color = backgroundColor
```

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　こう書けば　`StateTable` シートの `C2` セルの背景色を、 `GUI` シートの `A1` セルへ　コピーできるはずだぜ！」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　ボタンを押そうぜ？」  

![202301_excel_25-2120--ExecuteCopyColor-1.png](https://crieit.now.sh/upload_images/1568e0c8ce0e64297f4f75d776947c2963d11eb23fb6a.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　背景色は　コピーでけたが……。  
`Debug.Print( ... )` って何だぜ？」  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　ググりゃいいんじゃないの？」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　🔍 `VBA デバッグプリント` で検索」  

📖 [【エクセルVBA】初心者のうちから知っておくべきDebug.Printの使い方](https://tonari-it.com/excel-vba-debug-print/)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　**イミディエイト・ウィンドウ** に値を表示するらしいぜ」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　イミディエイト・ウィンドウ　って何だぜ？」  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　ググりゃいいんじゃないの？」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　🔍 `VBA イミディエイトウィンドウ` で検索」  

📖 [イミディエイトウィンドウの使い方](https://www.kenschool.jp/blog/?p=3430)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　よし　分かったぜ」  

![202301_excel_25-2125--Immediate-1.png](https://crieit.now.sh/upload_images/e6b3acf6e5aca1fc55867d968f4f342d63d120141a3d7.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　`[Ctrl] + [G]` キーを打鍵すると　出てくるウィンドウだぜ」  

![kifuwarabe-futsu.png](https://crieit.now.sh/upload_images/beaf94b260ae2602ca8cf7f5bbc769c261daf8686dbda.png)  
「　そんなウィンドウの出し方、　画面のどこを探しても　無くね？」  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　コンピューター開発者は　`私には分かる。だからお前も分かるだろ` という脳をしてる人　多いのよ。  
伝え、継承する精神を持っている人は　リタイア組ぐらいよ」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　天才が　技術を継承するの　損だしな。  
ネット上で　記事書いてるの　リタイア組か、　業界が滅ぶ一歩手前で　しかたなく　天才のケツ掃除してる人たちの　どっちかだよな」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　しかし `49407` なんて数字出てきても　嬉しくないな」  

📅 2023-01-25 wed 21:36  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　じゃあ　次は長めの　コンボ（Combo；連続技）　やるから　よく聴けだぜ」  

![202301_excel_25-2140--a1-1.png](https://crieit.now.sh/upload_images/10533ade231776b7c34ab00bb9c960e063d12359726bf.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　`Tape` シートの `A1` セルに入っている値 `A` と、その背景色　白色　を取得して……」  

![202301_excel_25-2142--seek-1.png](https://crieit.now.sh/upload_images/21fadc22632f5e65a9da1a98a748e35563d123daa45b8.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　`StateTable` シートの `State` 列に `A` が、  
`Read` 列に　背景色が白色のセルが  
無いかなと探し……」  

![202301_excel_25-2142--seek-2.png](https://crieit.now.sh/upload_images/4f19e8287c402995d82464e71220ebe463d12458a2c1d.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　見つけたら……」  

![202301_excel_25-2142--seek-3.png](https://crieit.now.sh/upload_images/55977720b323939598c8a932fa58ba7763d124a97cfd2.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　そのまま　横へスライドし、
Write列 の背景色は　オレンジ色、
Move列 は `>` 、
Transition列 は `B`
と　いったん覚え……」  

![202301_excel_25-2149--output-1.png](https://crieit.now.sh/upload_images/9e508dc25de95f97970c4606d32134a463d1260a86ee0.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　`Tape` シートを開き、  
A1 セルから １行下にいったところを、　Write 列にあるように　オレンジ色　に塗り、  
Move列が `>` とあるように　その右側に対して、  
Transition列があるように　`B`　を書き込もうぜ？」  

![kifuwarabe-futsu.png](https://crieit.now.sh/upload_images/beaf94b260ae2602ca8cf7f5bbc769c261daf8686dbda.png)  
「　長い」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　いったん `TuringMachineByExcelVBA.xlsm` ファイルを保存して閉じるぜ。  
休憩だぜ」  

📅 2023-01-25 wed 22:10 stop  

📅 2023-01-25 wed 22:21 restart  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　よっしゃ　再開だぜ！  
`TuringMachineByExcelVBA.xlsm` ファイルを開けて、っと」  

![202301_excel_25-2222--clear.png](https://crieit.now.sh/upload_images/a5986e9999f216f44cf41643061082f963d12d3d41e93.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　消えてる！」  

![kifuwarabe-futsu.png](https://crieit.now.sh/upload_images/beaf94b260ae2602ca8cf7f5bbc769c261daf8686dbda.png)  
「　よく探せだぜ！」  

![202301_excel_25-2224--project-window-1.png](https://crieit.now.sh/upload_images/911081e91fc626fedc3222f61a52f6b063d12db90ab2f.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　無い、無い、無い、  
どこにも　無～い！」  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　ファイルを開けるには、ダブル・クリック　するんじゃないの？」  

![202301_excel_25-2226--double-click.png](https://crieit.now.sh/upload_images/51fb39e17099b3f9fea0320e7586266d63d12e1966994.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　あ、　有った」  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　じゃあ　`Tape` シートの `A1` セルの値 `A` と背景色　白色　を取得して、  
そのタプル（Tuple；組み）が `StateTable` シートの何行目にあるか探し出して  
イミディエイト・ウィンドウに　デバッグプリント　するところまで　やりましょう！」  

![202301_excel_25-2230--getValue-1.png](https://crieit.now.sh/upload_images/aa9fb664e006bfd66d5afdddfecf45c063d12f44ec29c.png)  

```vba
Sub ボタン1_Click()
    Dim text As String
    Dim backgroundColor As Long
    
    text = Worksheets("Tape").Range("A1").Value ' セルの値
    backgroundColor = Worksheets("Tape").Range("A1").Interior.color ' 背景色
    
    Debug.Print (text)
    Debug.Print (backgroundColor)

    Worksheets("GUI").Range("A1").Value = text
    Worksheets("GUI").Range("A1").Interior.color = backgroundColor
End Sub
```

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　値と　背景色のコピーは　できるようになったが、  
次は　探すというやつだな」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　🔍 `VBA For文` で検索」  

📖 [[Excel で VBA] For 文による繰り返し](https://brain.cc.kogakuin.ac.jp/~kanamaru/lecture/vba2013/04-for01.html)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　よし　分かったぜ」  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　`For文` が何かの説明は　しないのね」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　手を動かせば　分かるぜ」  

![202301_excel_25-2240--loop-1.png](https://crieit.now.sh/upload_images/5608b24b4b67d86711a4619d515748f563d1317277ca9.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　多分　`StateTable` シートを　１行目から　７行目まで読むのは　こんな雰囲気だろ。  
VBA の if 文ってどうやって書くんだったかな？」  

![kifuwarabe-futsu.png](https://crieit.now.sh/upload_images/beaf94b260ae2602ca8cf7f5bbc769c261daf8686dbda.png)  
「　記憶してないの　わらう」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　🔍 `VBA if文` で検索」  

📖 [ExcelのVBA（マクロ）でIf～Then～Elseを使って条件分岐する方法](https://office-hack.com/excel/if-vba/)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　よし　分かったぜ」  

![202301_excel_25-2249--if-then-else-1.png](https://crieit.now.sh/upload_images/55c377d1e72e453ee5954f4471aa3a9a63d133aea4cb7.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　多分　if文は　こんな雰囲気だろ。  
Forループを途中で抜けるの　VBAで　どうやって書くんだったかな？」  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　ググりゃいいんじゃないの？」  

![kifuwarabe-futsu.png](https://crieit.now.sh/upload_images/beaf94b260ae2602ca8cf7f5bbc769c261daf8686dbda.png)  
「　なんにも覚えてないんだな」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　🔍 `VBA break文` で検索」  

📖　[Excel VBAでFor文を途中で抜ける：Exit](https://uxmilk.jp/48591)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　よし　分かったぜ」  

```vba
Sub ボタン1_Click()
    Dim text As String
    Dim backgroundColor As Long
    
    text = Worksheets("Tape").Range("A1").Value ' セルの値
    backgroundColor = Worksheets("Tape").Range("A1").Interior.color ' 背景色
    
    Debug.Print (text)
    Debug.Print (backgroundColor)

    Worksheets("GUI").Range("A1").Value = text
    Worksheets("GUI").Range("A1").Interior.color = backgroundColor
    
    Dim i As Long
    Dim stateText As String
    Dim readBackgroundColor As Long
    Dim writeBackgroundColor As Long
    Dim moveText As String
    Dim transitionText As String
    
    For i = 2 To 7
        stateText = Worksheets("StateTable").Range("A" & i).Value ' セルの値
        readBackgroundColor = Worksheets("StateTable").Range("B" & i).Interior.color ' 背景色
        
        ' 一致するか？
        If text = stateText And backgroundColor = readBackgroundColor Then
            writeBackgroundColor = Worksheets("StateTable").Range("C" & i).Interior.color ' 背景色
            moveText = Worksheets("StateTable").Range("D" & i).Value ' セルの値
            transitionText = Worksheets("StateTable").Range("E" & i).Value ' セルの値
            
            Debug.Print (writeBackgroundColor)
            Debug.Print (moveText)
            Debug.Print (transitionText)

            ' TODO 次の処理へ

            Exit For
        End If
    Next i
End Sub
```

![202301_excel_25-2257--find-1.png](https://crieit.now.sh/upload_images/055ad87739413e7ce40cde35dc7e9bc363d135e9993df.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　取れてるぜ」  

📅 2023-01-25 wed 23:00  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　次は、  
`Tape` シートの A1 セルを　スタート地点として、
１行　下りたセルの背景色を　Write列のいう色に塗って、そこから  
Move 列が `>` だったら その右のセルへ、 Transition 列のいうテキストを入れましょう」  

![202301_excel_25-2310--write-1.png](https://crieit.now.sh/upload_images/6d424c48b1e3b97a5bdd13889698200363d1391f18557.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　こんな感じだろ」  

![202301_excel_25-2314--tape.png](https://crieit.now.sh/upload_images/c1ded23ca69e31848ab042beebc7d81a63d1395c1c5c8.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　`Tape` シートの２行目を　クリアーしておくぜ。
そして `GUI` シートのボタンを押すぜ」  

![202301_excel_25-2316--result-1.png](https://crieit.now.sh/upload_images/e1df2e8107ce252e9c3a3120d8fa8df663d139f139a7d.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　色も文字もコピーされたぜ」  

![kifuwarabe-futsu.png](https://crieit.now.sh/upload_images/beaf94b260ae2602ca8cf7f5bbc769c261daf8686dbda.png)  
「　おつ。  
長いコンボが決まったな」  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　これで　１クロック　よね」  

![kifuwarabe-futsu.png](https://crieit.now.sh/upload_images/beaf94b260ae2602ca8cf7f5bbc769c261daf8686dbda.png)  
「　何だぜ　それ」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　次は　２クロック目に　行ってみようぜ？」  

📅 2023-01-25 wed 23:21  

![202301_excel_25-2321--2th-clock-1.png](https://crieit.now.sh/upload_images/4a20812bd5c1a3fc61525e8d4e1e598463d13b29bab9f.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　B2 セルをスタート地点として、　同様に　さっきと同じことを　やればいいんだぜ」  

![kifuwarabe-futsu.png](https://crieit.now.sh/upload_images/beaf94b260ae2602ca8cf7f5bbc769c261daf8686dbda.png)  
「　嫌だぜ　なんでそんなことをするんだぜ？」  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　ラジオ体操　みたいなもんよ。　頭を　ほぐしてんのよ」  

![202301_excel_25-2325--b-white-1.png](https://crieit.now.sh/upload_images/f21bba2889296010e5e957b6cd6612a863d13c296c3bc.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　文字が `B` で、背景色が　白色　なのは　４行目だな。  
下にオレンジ塗って　左へ　A　を書けばよさそうだな」  

![202301_excel_25-2337--time-1.png](https://crieit.now.sh/upload_images/78e64f2f57b3fec52bb5592a96eff7c963d13eff66930.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　説明するのを忘れていたが、上の行の背景色を、下の行へ　引き継ぐぜ」  

![kifuwarabe-futsu.png](https://crieit.now.sh/upload_images/beaf94b260ae2602ca8cf7f5bbc769c261daf8686dbda.png)  
「　忘れないでくれだぜ」  

![202301_excel_25-2340--time-b-1.png](https://crieit.now.sh/upload_images/75687cd8dc9973fb5a9d1c9436f4002263d13f9b620f9.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　下にオレンジ塗るぜ」  

![202301_excel_25-2346--time-c-1.png](https://crieit.now.sh/upload_images/0190cff67c269feba781ed9ff9d4c4b563d140e08a2db.png)

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　左に　A　を書いたら　こんな感じだな」  

![202301_excel_25-2348--Code.png](https://crieit.now.sh/upload_images/e2e349677ccb4abecf4a080e9c770deb63d14154d23b7.png)  

```vba
Sub ボタン1_Click()
    Dim text As String
    Dim backgroundColor As Long
    Dim i As Long
    Dim stateText As String
    Dim readBackgroundColor As Long
    Dim writeBackgroundColor As Long
    Dim moveText As String
    Dim transitionText As String
    
    ' 1回目の処理
    text = Worksheets("Tape").Range("A1").Value ' セルの値
    backgroundColor = Worksheets("Tape").Range("A1").Interior.color ' 背景色
    
    For i = 2 To 7
        stateText = Worksheets("StateTable").Range("A" & i).Value ' セルの値
        readBackgroundColor = Worksheets("StateTable").Range("B" & i).Interior.color ' 背景色
        
        ' 一致するか？
        If text = stateText And backgroundColor = readBackgroundColor Then
            writeBackgroundColor = Worksheets("StateTable").Range("C" & i).Interior.color ' 背景色
            moveText = Worksheets("StateTable").Range("D" & i).Value ' セルの値
            transitionText = Worksheets("StateTable").Range("E" & i).Value ' セルの値
            
            ' `Tape` シートの A1 セルの下のセルの背景色を　Write列のいう色に塗る
            Worksheets("Tape").Range("A2").Interior.color = writeBackgroundColor
            
            ' Move 列が `>` だったら その右のセルへ、 Transition 列のいうテキストを入れる
            If moveText = ">" Then
                Worksheets("Tape").Range("B2").Value = transitionText
            End If

            Exit For
        End If
    Next i
    
    ' TODO ★ 同様の2回目の処理
    text = Worksheets("Tape").Range("B2").Value ' セルの値
    backgroundColor = Worksheets("Tape").Range("B2").Interior.color ' 背景色

    ' ★ 上の行の背景色は引き継ぐ
    Worksheets("Tape").Range("A3").Interior.color = Worksheets("Tape").Range("A2").Interior.color
    Worksheets("Tape").Range("B3").Interior.color = Worksheets("Tape").Range("B3").Interior.color

    For i = 2 To 7
        stateText = Worksheets("StateTable").Range("A" & i).Value ' セルの値
        readBackgroundColor = Worksheets("StateTable").Range("B" & i).Interior.color ' 背景色
        
        ' 一致するか？
        If text = stateText And backgroundColor = readBackgroundColor Then
            writeBackgroundColor = Worksheets("StateTable").Range("C" & i).Interior.color ' 背景色
            moveText = Worksheets("StateTable").Range("D" & i).Value ' セルの値
            transitionText = Worksheets("StateTable").Range("E" & i).Value ' セルの値

            ' `Tape` シートの A1 セルの下のセルの背景色を　Write列のいう色に塗る
            Worksheets("Tape").Range("B3").Interior.color = writeBackgroundColor
            
            ' ★ Move 列が `<` だったら その左のセルへ、 Transition 列のいうテキストを入れる
            If moveText = "<" Then
                Worksheets("Tape").Range("A3").Value = transitionText
            End If

            Exit For
        End If
    Next i
    
End Sub
```

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　それを　コードにしたら　こんな感じだぜ」  

![kifuwarabe-futsu.png](https://crieit.now.sh/upload_images/beaf94b260ae2602ca8cf7f5bbc769c261daf8686dbda.png)  
「　長い」  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　これで　２クロック　よね」  

![kifuwarabe-futsu.png](https://crieit.now.sh/upload_images/beaf94b260ae2602ca8cf7f5bbc769c261daf8686dbda.png)  
「　何だぜ　それ」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　次は　３クロック目に　行ってみようぜ？」  

![kifuwarabe-futsu.png](https://crieit.now.sh/upload_images/beaf94b260ae2602ca8cf7f5bbc769c261daf8686dbda.png)  
「　嫌だぜ」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　今日は　ここまでとするが、　３クロック目　行くからな」  

📅 2023-01-25 wed 23:51 end  

# 📅2023-01-26 thu 18:53 start

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　VBA でサブルーチンは　どうやって書いたらいいんだぜ？」  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　ググりゃいいんじゃないの？」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　🔍 `VBA サブルーチン` で検索」  

📖 [Excel VBA 処理の一部をサブルーチン化するCallステートメント](https://kosapi.com/post-5008/)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　よし　分かったぜ」  

![202301_excel_26-1901--Subroutine-1.png](https://crieit.now.sh/upload_images/bd596f09e3b962c6b9f6b5e2603dca2263d24fb97805b.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　スケルトン（Skeleton；穴埋めの穴じゃない方）を書こうぜ」  

![202301_excel_26-1906--MoveCode-1.png](https://crieit.now.sh/upload_images/aeed27da28a1e2e6fb9834509983fad163d2514d6e3aa.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　スケルトンの中へ　コードを　こうやって　入れたらいいんじゃないかだぜ？」  

![202301_excel_26-1911--MovedCode.png](https://crieit.now.sh/upload_images/9b465017915b949d3a4152c38134c39863d251d1a8f17.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　つまり　こう」  

![202301_excel_26-1912--Call-1.png](https://crieit.now.sh/upload_images/7f3a7089c2ea734bbb3f14dbe451e7d163d2523bbff86.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　移動した跡の所には　コール文（Call Statement）を置いておこうぜ？」  

```vba
Sub ボタン1_Click()
    
    ' 1回目の処理
    Call On1stClock
    
    ' 同様の2回目の処理
    Call On2ndClock
    
End Sub

Private Sub On1stClock()
    ' １回目のクロック
    Dim text As String
    Dim backgroundColor As Long
    Dim i As Long
    Dim stateText As String
    Dim readBackgroundColor As Long
    Dim writeBackgroundColor As Long
    Dim moveText As String
    Dim transitionText As String
    
    text = Worksheets("Tape").Range("A1").Value ' セルの値
    backgroundColor = Worksheets("Tape").Range("A1").Interior.color ' 背景色
    
    For i = 2 To 7
        stateText = Worksheets("StateTable").Range("A" & i).Value ' セルの値
        readBackgroundColor = Worksheets("StateTable").Range("B" & i).Interior.color ' 背景色
        
        ' 一致するか？
        If text = stateText And backgroundColor = readBackgroundColor Then
            writeBackgroundColor = Worksheets("StateTable").Range("C" & i).Interior.color ' 背景色
            moveText = Worksheets("StateTable").Range("D" & i).Value ' セルの値
            transitionText = Worksheets("StateTable").Range("E" & i).Value ' セルの値
            
            ' `Tape` シートの A1 セルの下のセルの背景色を　Write列のいう色に塗る
            Worksheets("Tape").Range("A2").Interior.color = writeBackgroundColor
            
            ' Move 列が `>` だったら その右のセルへ、 Transition 列のいうテキストを入れる
            If moveText = ">" Then
                Worksheets("Tape").Range("B2").Value = transitionText
            End If

            Exit For
        End If
    Next i

End Sub

Private Sub On2ndClock()
    ' ２回目のクロック
    Dim text As String
    Dim backgroundColor As Long
    Dim i As Long
    Dim stateText As String
    Dim readBackgroundColor As Long
    Dim writeBackgroundColor As Long
    Dim moveText As String
    Dim transitionText As String

    text = Worksheets("Tape").Range("B2").Value ' セルの値
    backgroundColor = Worksheets("Tape").Range("B2").Interior.color ' 背景色

    ' ★ 上の行の背景色は引き継ぐ
    Worksheets("Tape").Range("A3").Interior.color = Worksheets("Tape").Range("A2").Interior.color
    Worksheets("Tape").Range("B3").Interior.color = Worksheets("Tape").Range("B3").Interior.color

    For i = 2 To 7
        stateText = Worksheets("StateTable").Range("A" & i).Value ' セルの値
        readBackgroundColor = Worksheets("StateTable").Range("B" & i).Interior.color ' 背景色
        
        ' 一致するか？
        If text = stateText And backgroundColor = readBackgroundColor Then
            writeBackgroundColor = Worksheets("StateTable").Range("C" & i).Interior.color ' 背景色
            moveText = Worksheets("StateTable").Range("D" & i).Value ' セルの値
            transitionText = Worksheets("StateTable").Range("E" & i).Value ' セルの値

            ' `Tape` シートの A1 セルの下のセルの背景色を　Write列のいう色に塗る
            Worksheets("Tape").Range("B3").Interior.color = writeBackgroundColor
            
            ' ★ Move 列が `<` だったら その左のセルへ、 Transition 列のいうテキストを入れる
            If moveText = "<" Then
                Worksheets("Tape").Range("A3").Value = transitionText
            End If

            Exit For
        End If
    Next i
End Sub
```

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　じゃあ　`GUI` シートのボタンを押そうぜ？」  

![202301_excel_26-1915--Check.png](https://crieit.now.sh/upload_images/60d46eddc740de260994cbc83f82edc463d252c114f71.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　ちゃんと動いてるな」  

![kifuwarabe-futsu.png](https://crieit.now.sh/upload_images/beaf94b260ae2602ca8cf7f5bbc769c261daf8686dbda.png)  
「　場所を移しただけだしな」  

📅2023-01-26 thu 19:16  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　３クロック目も　コピー貼り付けして作んの？」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　嫌になるだろ」  

![202301_excel_26-1919--OnClock-1.png](https://crieit.now.sh/upload_images/3ef93b3d925a27efdeb7ebe3aa75cd5d63d253cfe6a02.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　何回目のクロックでも使えるジェネラル（General）なサブルーチンを作ろうぜ？」  

![202301_excel_26-1911--MovedCode-diff.png](https://crieit.now.sh/upload_images/5b0a749e37816dd597d3d8c6ca75b62763d255a98666e.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　違うところは５か所ぐらいなんだから、ここを違わないようにすればいいわけだぜ」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　Ａ列の右隣は B列 だが、  
`A` の右は何か尋ねたら `B` が返ってくるような方法って VBA にあるのかだぜ？」  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　ググりゃいいんじゃないの？」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　🔍 `VBA 列アルファベット変換` で検索」  

📖 [【ExcelVBA】列名のアルファベットと列番号の数字を相互変換する](https://qiita.com/11295/items/c26017eb21cb319fd29d)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　よし　分かったぜ」  

![202301_excel_26-2110--General.png](https://crieit.now.sh/upload_images/692b189f8ba2c55c3ffb94fda73ba9b963d26dd08ec89.png)  

```vba
Sub ボタン1_Click()
    
    ' 1回目の処理
    Call OnClock("A", 1)
    
    ' 同様の2回目の処理
    Call OnClock("B", 2)
    
End Sub

Private Sub OnClock(previousFileAlphabet As String, previousRank As Long)
    ' TODO 毎クロック（ｎ回目のクロック）
    Dim previousText As String
    Dim previousBackgroundColor As Long
    Dim previousCell As String
    Dim currentRank As Long
    Dim currentCell As String
    Dim stateText As String
    Dim readBackgroundColor As Long
    Dim writeBackgroundColor As Long
    Dim moveText As String
    Dim transitionText As String
    Dim i As Long
    
    previousCell = previousFileAlphabet & previousRank
    currentRank = previousRank + 1
    currentCell = previousFileAlphabet & currentRank
    Debug.Print ("--------")
    Debug.Print ("previousFileAlphabet:" & previousFileAlphabet)
    Debug.Print ("previousRank        :" & previousRank)
    Debug.Print ("previousCell        :" & previousCell)
    Debug.Print ("currentRank         :" & currentRank)
    Debug.Print ("currentCell         :" & currentCell)
        
    ' 開始行の背景色は、次行に引き継ぐ
    If 2 <= previousRank Then
        Dim aBackgroundColor As Long
        Dim bBackgroundColor As Long
        aBackgroundColor = Worksheets("Tape").Range("A" & previousRank).Interior.color
        bBackgroundColor = Worksheets("Tape").Range("B" & previousRank).Interior.color
        Worksheets("Tape").Range("A" & currentRank).Interior.color = aBackgroundColor
        Worksheets("Tape").Range("B" & currentRank).Interior.color = bBackgroundColor
        Debug.Print ("aBackgroundColor:" & aBackgroundColor)
        Debug.Print ("bBackgroundColor:" & bBackgroundColor)
    End If

    previousText = Worksheets("Tape").Range(previousCell).Value                             ' 開始セルの値
    previousBackgroundColor = Worksheets("Tape").Range(previousCell).Interior.color         ' 開始セルの背景色
    Debug.Print ("previousText           :" & previousText)
    Debug.Print ("previousBackgroundColor:" & previousBackgroundColor)

    For i = 2 To 7
        stateText = Worksheets("StateTable").Range("A" & i).Value                           ' 状態テーブルのState値
        readBackgroundColor = Worksheets("StateTable").Range("B" & i).Interior.color        ' 状態テーブルのRead列の背景色
        Debug.Print ("stateText           :" & stateText)
        Debug.Print ("readBackgroundColor :" & readBackgroundColor)
        
        ' 一致するか？
        If previousText = stateText And previousBackgroundColor = readBackgroundColor Then
            writeBackgroundColor = Worksheets("StateTable").Range("C" & i).Interior.color   ' 状態テーブルのWrite列の背景色
            moveText = Worksheets("StateTable").Range("D" & i).Value                        ' 状態テーブルのMove列の値
            transitionText = Worksheets("StateTable").Range("E" & i).Value                  ' 状態テーブルのTransition列の値
            Debug.Print ("writeBackgroundColor:" & writeBackgroundColor)
            Debug.Print ("moveText            :" & moveText)
            Debug.Print ("transitionText      :" & transitionText)

            ' `Tape` シートの A1 セルの下のセルの背景色を　Write列のいう色に塗る
            Worksheets("Tape").Range(currentCell).Interior.color = writeBackgroundColor
            
            Dim horizontal As Long      ' 水平方向
            If moveText = ">" Then      ' Move 列が `>` だったら その右のセルへ
                horizontal = 1
            ElseIf moveText = "<" Then  ' Move 列が `<` だったら その左のセルへ
                horizontal = -1
            End If
            Debug.Print ("horizontal:" & horizontal)
            
            ' Transition 列のいうテキストを入れる
            Dim startFileNumber As Integer
            Dim nextFileAlphabet As String
            startFileNumber = Columns(previousFileAlphabet).Column
            nextFileAlphabet = Split(Cells(1, startFileNumber + horizontal).Address, "$")(1)
            Debug.Print ("startFileNumber :" & startFileNumber)
            Debug.Print ("nextFileAlphabet:" & nextFileAlphabet)
            Worksheets("Tape").Range(nextFileAlphabet & currentRank).Value = transitionText

            Exit For
        End If
    Next i
End Sub
```

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　けっこう　大がかりに　変えることになってしまったぜ」  

📅2023-01-26 thu 21:12  

![kifuwarabe-futsu.png](https://crieit.now.sh/upload_images/beaf94b260ae2602ca8cf7f5bbc769c261daf8686dbda.png)  
「　こんなん　何がどう変わったのか　読者　分からんだろ」  

![202301_excel_26-2114--3rdClock-1.png](https://crieit.now.sh/upload_images/fdfc0cf2fdc2fb0f0adac6ef4e59d0a863d26ebbbdd65.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　１クロック目と　２クロック目で違うところは、　スタート地点の列番号と、行番号だけだったということだぜ」    

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　このように　２つのサブルーチンの差異が　サブルーチンの外に押し出されたものを　**アーギュメント**（Argument；実引数）と呼ぶ」    

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　ふーん」  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　３クロック目は　どう書くの？」  

![202301_excel_26-2121--argument-1.png](https://crieit.now.sh/upload_images/0a6c4a73d2935ca90b25ac2e8473e95563d2709e4c167.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　`A1` とか `B2` というのは、１クロック前に居たセルだぜ。  
だから　前の計算結果を　もらうといい。  
書き直そう」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　VBA でファンクションは　どうやって書いたらいいんだぜ？」  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　ググりゃいいんじゃないの？」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　🔍 `VBA ファンクション` で検索」  

📖 [VBA　Functionプロシージャについて　～関数の解説と使用例～](https://www.bold.ne.jp/engineer-club/vba-function)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　よし　分かったぜ」  

![202301_excel_26-2141--function.png](https://crieit.now.sh/upload_images/b537b697560f1fb0aeeaadc57a17f15b63d2751076122.png)  

```vba
Sub ボタン1_Click()

    Dim resultCell As String
    
    ' 1回目の処理
    resultCell = OnClock("A1")
    
    ' 同様の2回目の処理
    resultCell = OnClock(resultCell)
    
End Sub

Private Function OnClock(previousCell As String) As String
    ' 毎クロック（ｎ回目のクロック）
    Dim previousText As String
    Dim previousBackgroundColor As Long
    Dim currentRank As Long
    Dim currentCell As String
    Dim stateText As String
    Dim readBackgroundColor As Long
    Dim writeBackgroundColor As Long
    Dim moveText As String
    Dim transitionText As String
    Dim i As Long
    
    previousFileAlphabet = Split(Cells(1, Range(previousCell).Column).Address, "$")(1)
    previousRank = Range(previousCell).Row
    currentRank = previousRank + 1
    currentCell = previousFileAlphabet & currentRank
    Debug.Print ("--------")
    Debug.Print ("previousCell        :" & previousCell)
    Debug.Print ("previousFileAlphabet:" & previousFileAlphabet)
    Debug.Print ("previousRank        :" & previousRank)
    Debug.Print ("currentRank         :" & currentRank)
    Debug.Print ("currentCell         :" & currentCell)
        
    ' 開始行の背景色は、次行に引き継ぐ
    If 2 <= previousRank Then
        Dim aBackgroundColor As Long
        Dim bBackgroundColor As Long
        aBackgroundColor = Worksheets("Tape").Range("A" & previousRank).Interior.color
        bBackgroundColor = Worksheets("Tape").Range("B" & previousRank).Interior.color
        Worksheets("Tape").Range("A" & currentRank).Interior.color = aBackgroundColor
        Worksheets("Tape").Range("B" & currentRank).Interior.color = bBackgroundColor
        Debug.Print ("aBackgroundColor:" & aBackgroundColor)
        Debug.Print ("bBackgroundColor:" & bBackgroundColor)
    End If

    previousText = Worksheets("Tape").Range(previousCell).Value                             ' 開始セルの値
    previousBackgroundColor = Worksheets("Tape").Range(previousCell).Interior.color         ' 開始セルの背景色
    Debug.Print ("previousText           :" & previousText)
    Debug.Print ("previousBackgroundColor:" & previousBackgroundColor)

    For i = 2 To 7
        stateText = Worksheets("StateTable").Range("A" & i).Value                           ' 状態テーブルのState値
        readBackgroundColor = Worksheets("StateTable").Range("B" & i).Interior.color        ' 状態テーブルのRead列の背景色
        Debug.Print ("stateText           :" & stateText)
        Debug.Print ("readBackgroundColor :" & readBackgroundColor)
        
        ' 一致するか？
        If previousText = stateText And previousBackgroundColor = readBackgroundColor Then
            writeBackgroundColor = Worksheets("StateTable").Range("C" & i).Interior.color   ' 状態テーブルのWrite列の背景色
            moveText = Worksheets("StateTable").Range("D" & i).Value                        ' 状態テーブルのMove列の値
            transitionText = Worksheets("StateTable").Range("E" & i).Value                  ' 状態テーブルのTransition列の値
            Debug.Print ("writeBackgroundColor:" & writeBackgroundColor)
            Debug.Print ("moveText            :" & moveText)
            Debug.Print ("transitionText      :" & transitionText)

            ' `Tape` シートの A1 セルの下のセルの背景色を　Write列のいう色に塗る
            Worksheets("Tape").Range(currentCell).Interior.color = writeBackgroundColor
            
            Dim horizontal As Long      ' 水平方向
            If moveText = ">" Then      ' Move 列が `>` だったら その右のセルへ
                horizontal = 1
            ElseIf moveText = "<" Then  ' Move 列が `<` だったら その左のセルへ
                horizontal = -1
            End If
            Debug.Print ("horizontal:" & horizontal)
            
            ' Transition 列のいうテキストを入れる
            Dim previousFileNumber As Integer
            Dim nextFileAlphabet As String
            Dim nextCell As String
            previousFileNumber = Columns(previousFileAlphabet).Column
            nextFileAlphabet = Split(Cells(1, previousFileNumber + horizontal).Address, "$")(1)
            nextCell = nextFileAlphabet & currentRank
            Debug.Print ("previousFileNumber :" & previousFileNumber)
            Debug.Print ("nextFileAlphabet   :" & nextFileAlphabet)
            Debug.Print ("nextCell           :" & nextCell)
            Worksheets("Tape").Range(nextCell).Value = transitionText

            ' 関数から抜ける
            OnClock = nextCell
            Exit Function
            
        End If
    Next i
End Function
```

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　さらに　改造してしまったぜ」  

📅2023-01-26 thu 21:43  

![kifuwarabe-futsu.png](https://crieit.now.sh/upload_images/beaf94b260ae2602ca8cf7f5bbc769c261daf8686dbda.png)  
「　こんなん　何がどう変わったのか　読者　分からんだろ」  

![202301_excel_26-2146--main-1.png](https://crieit.now.sh/upload_images/96464638d29440e381b1a1a8848dbbec63d27668d882a.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　このように　前の計算結果を使って　また同様に計算するように作るのは　よくあるテクニックだぜ。  
**漸化式**（ぜんかしき） とか　これだな」  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　さっさと　３クロック目を書きなさいよ」  

![202301_excel_26-2150--3rdClock-1.png](https://crieit.now.sh/upload_images/fb4fbbb839b9dbe36e947cc566b9f0bd63d2773c25cbc.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　３クロック目は、２クロック目と同じコードで　いいはずなんだぜ。  
動かしてみよう」  

![202301_excel_26-2152--error-1.png](https://crieit.now.sh/upload_images/66f641f330add453cbf13187a107f99a63d277b62b870.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　エラーが出たぜ」  

![kifuwarabe-futsu.png](https://crieit.now.sh/upload_images/beaf94b260ae2602ca8cf7f5bbc769c261daf8686dbda.png)  
「　直せ！」  

![202301_excel_26-2153--debug-1.png](https://crieit.now.sh/upload_images/0f024dcbc6526f5a4dcbcdef822d799d63d27800a2056.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　`[デバッグ（D）]` ボタンを押してみようぜ？」  

![202301_excel_26-2155--highlight.png](https://crieit.now.sh/upload_images/e034203acef8d1e6ea39d674be15a19663d278556be61.png)  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　エラーがあった行をハイライト（Highlight）してくれてるのかもしれないけど　説明がないから分かんないわねぇ」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　また　ショートカットを打鍵すれば　ウィンドウが出てくるのかだぜ？」  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　ググりゃいいんじゃないの？」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　🔍 `VBA スタックトレース` で検索」  

📖 [VBAナメてた](https://speakerdeck.com/narazaka/vbanameteta?slide=116)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　行番号とエラーコードしか　分かんないらしいぜ」  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　1994年の環境ねえ」  

📅2023-01-26 thu 22:01  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　🔍 `VBA 1004 Error` で検索」  

📖 [VBA Error 1004 in Excel](https://www.wallstreetmojo.com/vba-1004-error/#:~:text=This%20error%20occurs%20when%20you,name%20with%20the%20excel%20extension)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　無いシート名や、無いアドレスを指定したら　出てくるエラーかも知らん」  

![202301_excel_26-2207--error.png](https://crieit.now.sh/upload_images/3265aec10a1e7becdb64caf901d1eb1763d27b13390ff.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　A列より　左に進もうとして　０列目は無いので　エラーになったんだな」  

![kifuwarabe-futsu.png](https://crieit.now.sh/upload_images/beaf94b260ae2602ca8cf7f5bbc769c261daf8686dbda.png)  
「　A列の左に　列を挿入しろだぜ」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　🔍 `VBA 列の挿入` で検索」  

📖 [VBA 列を挿入する](https://www.tipsfound.com/vba/09005)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　よし　分かったぜ」  

![202301_excel_26-2215--bugfix-1.png](https://crieit.now.sh/upload_images/9c1455ad9824e532e31e7a03b5cf779363d27d2c21916.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　０列にアクセスしようとしたら、  
A列の左に　列挿入して、　列番地も左に　１つずらして　１列にアクセスするように変更しよう」  

![202301_excel_26-2218--3rdClockDone.png](https://crieit.now.sh/upload_images/6e7a3a2dcf4463fa7169bed325e8f84163d27da7bd035.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　３クロック目も　動いたようだな」  

📅2023-01-26 thu 22:18  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　じゃあ　あと　１００クロックぐらい　動かしましょう」  

![202301_excel_26-2226--For-1.png](https://crieit.now.sh/upload_images/fb04bbd9f19fbc722de48dff01546cd663d27fa2132bb.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　`For` 文を使えば　１００行書かずに済むから楽だよな。  
漸化式の形に作っておけば　繰り返しの回数を指定するだけでいいから　手間要らずだぜ」  

![202301_excel_26-2229--Error-1.png](https://crieit.now.sh/upload_images/a1471049d3f876f18e31d667018db7ba63d2805d6d23d.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　エラーが出た」  

![kifuwarabe-futsu.png](https://crieit.now.sh/upload_images/beaf94b260ae2602ca8cf7f5bbc769c261daf8686dbda.png)  
「　🔍 `VBA 1004 Error` で検索しろだぜ」  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　さっきと同じ 1004 番よ」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　今度は何が無かったんだろな？」  

![202301_excel_26-2237--empty-previous-cell-1.png](https://crieit.now.sh/upload_images/e12e7323f7bb4b8c004c359ac49f51f463d282597704d.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　`previousCell` 変数の中身が空文字列だぜ。」  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　`A1` とか、何か番地が入っているはずじゃない」  

![202301_excel_26-2234--reset-button-1.png](https://crieit.now.sh/upload_images/8e8dc3028f75bf93b711ed6b1dc4afeb63d2817a7f5e3.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　`リセット` ボタンを押すと　デバッグは止まるぜ」  

![202301_excel_26-2240--Halt-1.png](https://crieit.now.sh/upload_images/1d653278fa64d4c114346024b8a8d96563d282f95697e.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　多分、 `HALT` （ハルト）が出たら　終了しないといけないんだぜ」  

![202301_excel_26-2246--success-1.png](https://crieit.now.sh/upload_images/841a4e08b5050abfa210923ba8fae31463d28497a867f.png)  

```vba
            If transitionText = "HALT" Then
                ' 関数から抜ける
                OnClock = "SUCCESS"
                Exit Function
            End If
```

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　じゃあ　`HALT` （ハルト）が出たときは　セル番地ではなく　`SUCCESS`　を返すという取り決めにしておこうぜ」  

![202301_excel_26-2252--exitFor-1.png](https://crieit.now.sh/upload_images/c564d0d7d0a67a74807f51a08fe8c7e063d285d9d674a.png)  

```vba
        ' 正常終了時はループから抜ける
        If resultCell = "SUCCESS" Then
            Exit For
        End If
```

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　関数の呼び出し元でも　前判定で　`resultCell` 変数の内容が `SUCCESS` だったら、ループから抜けるようにすれば　コードも短いぜ」  

![202301_excel_26-2255--play.png](https://crieit.now.sh/upload_images/fb5cfeec04b8547813467f8277e03b0363d2866e64c44.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　なんか　動いた跡があるな」  

![202301_excel_26-2257--QUOTE-Wikipedia.png](https://crieit.now.sh/upload_images/fcddc59f9b69d6982a3cfa1957cb583463d286dede709.png)  

📖 [Turing machine](https://en.wikipedia.org/wiki/Turing_machine)  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　👆　Wikipedia のサンプルと違くない？」  

![202301_excel_26-2308--copy-paint-1.png](https://crieit.now.sh/upload_images/28e374907e108dedeadc637e6038f08c63d2896c1f1a6.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　A列と B列しか　背景色をコピーしてないからな」  

![kifuwarabe-futsu.png](https://crieit.now.sh/upload_images/beaf94b260ae2602ca8cf7f5bbc769c261daf8686dbda.png)  
「　テープの長さと同じ幅だけ　コピーしてくれだぜ」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　じゃあ　テープの長さも　数えておかないといけないな」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　VBA で、自由変数の使い方、または　引数の参照渡しのやり方は　どう書くんだろうな？」  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　ググりゃいいんじゃないの？」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　🔍 `VBA 引数の参照渡し` で検索」  

📖 [値渡し（ByVal）、参照渡し（ByRef）について](https://excel-ubara.com/excelvba4/EXCEL218.html)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　よし　分かったぜ」  

![kifuwarabe-futsu.png](https://crieit.now.sh/upload_images/beaf94b260ae2602ca8cf7f5bbc769c261daf8686dbda.png)  
「　ぜったい　そんな記事を　ちらっと読んでも　何も分からないけどな」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　プログラムの組み方は分かってるが　VBA での書き方が分からないだけなんで、ちら見で　充分だぜ」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　あれっ？　`ByRef tapeWidth As Long`　と書いても　参照渡しになって無（ね）。  
調べても分からないことはある。  
自由変数にするか」

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　🔍 `VBA グローバル変数` で検索」  

📖 [変数の適用範囲](http://officetanaka.net/excel/vba/variable/05.htm)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　よし　分かったぜ」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　でも　書いているうちに　`tapeWith`　というスペルミスを見つけて　`tapeWidth` 直したら動くようになった」  

![202301_excel_26-2354--codeA-1.png](https://crieit.now.sh/upload_images/af1d389778a107a54445f1e0442c878963d294a4512f6.png)  
![202301_excel_26-2357--codeB-1.png](https://crieit.now.sh/upload_images/b55aa54001335ce25c06ca45f7c876a263d294fddbcb3.png)  

```vba
Sub ボタン1_Click()

    Dim resultCell As String
    Dim tapeWidth As Long
    tapeWidth = 1   ' "A1" セル１つ分
    
    ' 1回目の処理
    resultCell = OnClock("A1", tapeWidth)

    ' 同様の i 回目の処理
    Dim i As Long
    For i = 2 To 100
    
        ' 正常終了時はループから抜ける
        If resultCell = "SUCCESS" Then
            Exit For
        End If
    
        resultCell = OnClock(resultCell, tapeWidth)
    Next i
    
End Sub

Private Function OnClock(previousCell As String, ByRef tapeWidth) As String
    ' 毎クロック（ｎ回目のクロック）
    Dim previousFileAlphabet As String
    Dim previousRank As Long
    Dim previousText As String
    Dim previousBackgroundColor As Long
    Dim currentRank As Long
    Dim currentCell As String
    Dim stateText As String
    Dim readBackgroundColor As Long
    Dim writeBackgroundColor As Long
    Dim moveText As String
    Dim transitionText As String
    Dim i As Long
    
    previousFileAlphabet = Split(Cells(1, Range(previousCell).Column).Address, "$")(1)
    previousRank = Range(previousCell).Row
    currentRank = previousRank + 1
    currentCell = previousFileAlphabet & currentRank
    Debug.Print ("--------")
    Debug.Print ("previousCell        :" & previousCell)
    Debug.Print ("previousFileAlphabet:" & previousFileAlphabet)
    Debug.Print ("previousRank        :" & previousRank)
    Debug.Print ("currentRank         :" & currentRank)
    Debug.Print ("currentCell         :" & currentCell)
    Debug.Print ("tapeWidth           :" & tapeWidth)
        
    ' 開始行の背景色は、次行に引き継ぐ
    If 2 <= previousRank Then
        For i = 1 To (tapeWidth + 1)
            Dim tempFileAlphabet As String
            Dim tempBackgroundColor As Long
            
            tempFileAlphabet = Split(Cells(1, i).Address, "$")(1)
            tempBackgroundColor = Worksheets("Tape").Range(tempFileAlphabet & previousRank).Interior.color
            Debug.Print ("i                  :" & i)
            Debug.Print ("tempFileAlphabet   :" & tempFileAlphabet)
            Debug.Print ("tempBackgroundColor:" & tempBackgroundColor)
            Worksheets("Tape").Range(tempFileAlphabet & currentRank).Interior.color = tempBackgroundColor
        Next i
    End If

    previousText = Worksheets("Tape").Range(previousCell).Value                             ' 開始セルの値
    previousBackgroundColor = Worksheets("Tape").Range(previousCell).Interior.color         ' 開始セルの背景色
    Debug.Print ("previousText           :" & previousText)
    Debug.Print ("previousBackgroundColor:" & previousBackgroundColor)

    For i = 2 To 7
        stateText = Worksheets("StateTable").Range("A" & i).Value                           ' 状態テーブルのState値
        readBackgroundColor = Worksheets("StateTable").Range("B" & i).Interior.color        ' 状態テーブルのRead列の背景色
        Debug.Print ("stateText           :" & stateText)
        Debug.Print ("readBackgroundColor :" & readBackgroundColor)
        
        ' 一致するか？
        If previousText = stateText And previousBackgroundColor = readBackgroundColor Then
            writeBackgroundColor = Worksheets("StateTable").Range("C" & i).Interior.color   ' 状態テーブルのWrite列の背景色
            moveText = Worksheets("StateTable").Range("D" & i).Value                        ' 状態テーブルのMove列の値
            transitionText = Worksheets("StateTable").Range("E" & i).Value                  ' 状態テーブルのTransition列の値
            Debug.Print ("writeBackgroundColor:" & writeBackgroundColor)
            Debug.Print ("moveText            :" & moveText)
            Debug.Print ("transitionText      :" & transitionText)

            ' `Tape` シートの A1 セルの下のセルの背景色を　Write列のいう色に塗る
            Worksheets("Tape").Range(currentCell).Interior.color = writeBackgroundColor
            
            Dim horizontal As Long      ' 水平方向
            If moveText = ">" Then      ' Move 列が `>` だったら その右のセルへ
                horizontal = 1
            ElseIf moveText = "<" Then  ' Move 列が `<` だったら その左のセルへ
                horizontal = -1
            End If
            Debug.Print ("horizontal:" & horizontal)
            
            ' Transition 列のいうテキストを入れる
            Dim previousFileNumber As Integer
            Dim nextFileAlphabet As String
            Dim nextFileNumber As Integer
            Dim nextCell As String
            previousFileNumber = Columns(previousFileAlphabet).Column
            nextFileNumber = previousFileNumber + horizontal
            Debug.Print ("previousFileNumber :" & previousFileNumber)
            Debug.Print ("nextFileNumber     :" & nextFileNumber)
            
            If nextFileNumber = 0 Then
                Worksheets("Tape").Columns("A").Insert  ' A 列の左に列挿入
                nextFileNumber = nextFileNumber + 1     ' 列番号を 1 つ左へずらす
                tapeWidth = tapeWidth + 1               ' Update
            ElseIf tapeWidth < nextFileNumber Then
                tapeWidth = nextFileNumber              ' Update
            End If
            
            nextFileAlphabet = Split(Cells(1, nextFileNumber).Address, "$")(1)
            nextCell = nextFileAlphabet & currentRank
            Debug.Print ("nextFileAlphabet   :" & nextFileAlphabet)
            Debug.Print ("nextCell           :" & nextCell)
            Worksheets("Tape").Range(nextCell).Value = transitionText

            If transitionText = "HALT" Then
                ' 関数から抜ける
                OnClock = "SUCCESS"
                Exit Function
            End If

            ' 関数から抜ける
            OnClock = nextCell
            Exit Function
            
        End If
    Next i
End Function
```

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　よし　`tapeWidth` も付けたぜ」  

![202301_excel_27-0000--play.png](https://crieit.now.sh/upload_images/91ed806503af3b04a8d5394f6273ba2863d29593358bd.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　チューリング・マシンが　でけたぜ」  

📅2023-01-27 thu 00:01

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　`HALT` は出力しないのが　チューリング・マシンみたいよ？」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　じゃあ　`StateTable`　の　`Transition`　列で　`HALT`　を見つけたところで　止めた方がいいのか」  

![202301_excel_27-0005--moveCode-1.png](https://crieit.now.sh/upload_images/abf090d159ec50ccbce29faabf09177763d297118951f.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　判定部を　上に持っていけば　いいんじゃないかだぜ？」  

![202301_excel_27-0008--copy-1.png](https://crieit.now.sh/upload_images/399be2584be9b9b2bba3b35100bb8eea63d29777bbde4.png)  

![ohkina-hiyoko-futsu2.png](https://crieit.now.sh/upload_images/96fb09724c3ce40ee0861a0fd1da563d61daf8a09d9bc.png)  
「　👆　１行　余分に背景色をコピーしてるわよ？」  

![202301_excel_27-0010--copyB-1.png](https://crieit.now.sh/upload_images/a8099b2405cf0ef7b4787895939535d763d29816eaaea.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　背景色のコピーを、判定部の後ろに持っていけば　いいんじゃないかだぜ？」  

![202301_excel_27-0012--error.png](https://crieit.now.sh/upload_images/12497135a2ffdac5dce45202516e999e63d2987639a62.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　おっと」  

![202301_excel_27-0013--define-1.png](https://crieit.now.sh/upload_images/4eaad7541e0323d5f67ce285389749c163d298cfdb63f.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　変数 `i` は、外側のループの　ループ・カウンターとして使ってたか」  

![202301_excel_27-0016--J-1.png](https://crieit.now.sh/upload_images/6c2e941fdc56048c2fc58049fa4bec4963d2997f8510f.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　じゃあ　変数 `j` を作ればいいや」  

![202301_excel_27-0018--ok.png](https://crieit.now.sh/upload_images/82af7bdc12bd9c36c10effeab11fb3f463d299c572cee.png)  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　👆　これで　チューリング・マシン　は完成だぜ。  
一般化してないのと、不具合が残っているのが　残っているかも知れないが……」  

![kifuwarabe-futsu.png](https://crieit.now.sh/upload_images/beaf94b260ae2602ca8cf7f5bbc769c261daf8686dbda.png)  
「　これは何に使うんだぜ？」  

![ramen-tabero-futsu2.png](https://crieit.now.sh/upload_images/d27ea8dcfad541918d9094b9aed83e7d61daf8532bbbe.png)  
「　小さなメモリを使って、もう少し色々動き回る　**自動人形**（オートマタ；automata）を作ったり、  
もっと大きなものを作るのに使われるぜ  」



# // 書きかけ