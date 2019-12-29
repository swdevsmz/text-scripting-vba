# VBE 開発効率化

## 拡張機能 
- VBE拡張機能 RubberDuck
https://qiita.com/mochimo/items/e9be36619a76e15bc898

## シンタックスハイライト（エディタの色）
- フォント  
https://nelog.jp/myrica

- おすすめ配色  
https://kazusa-pg.com/vba-editor-recommended-settings/


- レジストリに下記エントリーを追加  
    HKEY_CURRENT_USER\Software\Microsoft\VBA\7.1\Common
    ```
    CodeBackColors
    4 0 5 7 6 4 4 4 0 0 0 0 0 0 0 0 

    CodeForeColors
    1 0 0 0 1 9 11 5 0 0 0 0 0 0 0 0 
    ```

### 上記以外の色を使いたい場合は「VBEThemeColorEditor」

- VBEThemeColorEditor  
https://qiita.com/kabylake/items/4434a4793cca09a051e5

- VBE7.DLL の場所
    ```
    C:\Program Files (x86)\Microsoft Office\root\vfs\ProgramFilesCommonX86\Microsoft Shared\VBA\VBA7.1\
    ```

### その他
- 行番号表示アドイン  
    32bit  
    http://mtj-k.no.coocan.jp/software/office_vb6/addlinenumbers_vb6.html  
    64bit  
    http://mtj-k.no.coocan.jp/software/office_vb6/addlinenumbers_vba_x64.html
- コーディングガイドライン  
https://qiita.com/mima_ita/items/8b0eec3b5a81f168822d

## VBE ショートカット
http://span.jp/office2010_manual/excel_vba/reference/excel-vba-shortcut.html

- F8 ステップイン
- Shift + F8 ステップオーバー
- Shift + F2 ：定義の表示  
- Ctrl + Shift + F2 ：定義の表示から元の位置へ戻る
- Ctrl + I キー クイックヒントの表示
- Ctrl + G	イミディエイトウィンドウの表示、移動

## RubberDuckのショートカット
- Ctrl + Shift+E ソースのエクスポート
- Ctrl + Shift+R リネーム
- Ctrl + M       モジュールのインデント整形
- Ctrl + P       プロシージャのインデント整形
- Ctrl + R       コードエクスプローラー
- Ctrl + T       シンボルの検索 