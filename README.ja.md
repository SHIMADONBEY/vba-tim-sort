# vba-tim-sort

VBA で Tim Sort アルゴリズムを実装したライブラリです。

[English](README.md) | 日本語

## 導入方法

1. ファイルのインポート
   - VBA エディタを開き、メニューの「ファイル」→「ファイルのインポート」で次のファイルをプロジェクトに追加します。  
     - [vba-files/VbaTimSort.bas](vba-files/VbaTimSort.bas#L1)  
     - [vba-files/IComparator.cls](vba-files/IComparator.cls#L1)  
   - `IComparator` の実装サンプルは `vba-files/test/comparators/` にあります。必要に応じてプロジェクトに追加してください。

2. 使い方 （配列）

``` vb
Dim arr As Variant
arr = Array(5, 2, 9, 1)

' 自然順（数値・文字列・日付）でソートする場合は comparator に Nothing を渡す
Dim sorted As Variant
sorted = SortArrayInPlace(arr, Nothing)    ' 昇順

' 降順にする場合
Dim sortedDesc As Variant
```

3. 使い方 （Collection）

``` vb
Dim coll As New Collection
coll.Add "b"
coll.Add "a"
Dim sortedColl As Collection
Set sortedColl = SortCollection(coll, Nothing)
```

4. IComparator を実装して使う例

``` vb: MyComparator.cls
Implements IComparator

Private Function IComparator_Compare(ByVal a As Variant, ByVal b As Variant) As Integer
    ' カスタム比較ロジック（例：数値比較）
    If a < b Then
        IComparator_Compare = -1
    ElseIf a > b Then
        IComparator_Compare = 1
    Else
        IComparator_Compare = 0
    End If
End Function
```

``` vb: example.bas
' 使用側
Dim comp As IComparator
Set comp = New MyComparator
Dim sortedWithComp As Variant
sortedWithComp = SortArrayInPlace(arr, comp)
```

5. 補足

- `SortArrayInPlace()` / `SortCollection()` は元の並びを変えず、新しい配列／Collection を返します。
- オブジェクトを比較する場合は、`IComparator` を必ず提供してください。（渡さないとエラーになります。）

## OSS プロジェクトガイドライン

このプロジェクトはオープンソースソフトウェア (OSS) として運用されています。
以下は、利用者とコントリビューターに向けた一般的な方針です。

### ライセンス

本プロジェクトのライセンス条件は [LICENSE](LICENSE) に記載されています。
利用、改変、再配布を行う場合は、同ライセンスに従ってください。

### コントリビュート

コントリビューションを歓迎します。
一般的な流れは以下のとおりです。

1. リポジトリをフォークする。
2. 機能追加または修正用のブランチを作成する。
3. 振る舞いを変更した場合は、テストを追加または更新する。
4. 変更内容を明確にしたプルリクエストを作成する。

変更は小さく焦点を絞り、レビューしやすい形を意識してください。

### Issue と要望

不具合報告や機能要望は Issue で受け付けます。
可能であれば、次の情報を含めてください。

- 期待される動作
- 実際の動作
- 再現手順
- VBA / Excel のバージョン情報

### セキュリティ

[セキュリティポリシー（日本語）](SECURITY.ja.md) をご参照ください。

### サポート範囲

本プロジェクトはベストエフォートでメンテナンスされています。
返信速度やリリース時期は保証されません。

### 謝辞

Issue 報告、ドキュメント改善、コード提供に協力してくださるすべての方に感謝します。
