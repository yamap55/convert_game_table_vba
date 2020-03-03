# convert_game_table_vba
ExcelのVBA（Visual Basic for Applications）を使用し、特定のシート内容を元に新たに別のシートを作成する。
具体的には下記の勝敗表シート、名簿シートから出力シートを作成する。

## データサンプル
### original（勝敗表）
| 日付      | 曜日 | 対戦者① | 結果① | 枚数 | 結果② | 対戦者② |
|-----------|------|---------|-------|------|-------|---------|
| 2020/2/25 | 火   | 鈴木    | ○     | 5    | ×     | 北川    |
| 2020/2/26 | 水   | 高橋    | ×     | 3    | ○     | 伊藤    |
| 2020/2/27 | 木   | 田中    | ○     | 2    | ×     | 鈴木    |
| 2020/2/28 | 金   | 北川    | ×     | 1    | ○     | 鈴木    |
| 2020/2/29 | 土   | 伊藤    | ○     | 1    | ×     | 鈴木    |
| 2020/3/1  | 日   | 鈴木    | ○     | 2    | ×     | 高橋    |
| 2020/3/2  | 月   | 田中    | ○     | 5    | ×     | 北川    |

### meibo（名簿）
| No. | 名前 |
|-----|------|
| 1   | 鈴木 |
| 2   | 高橋 |
| 3   | 田中 |
| 4   | 北川 |
| 5   | 伊藤 |
| 6   | 山下 |

### 結果（出力）
|      | 鈴木 | 高橋 | 田中 | 北川 | 伊藤 | 山下 |
|------|------|------|------|------|------|------|
| 鈴木 | *    | ○    | ×    | ○,○  | ×    | 　   |
| 高橋 | ×    | *    | 　   | 　   | ×    | 　   |
| 田中 | ○    | 　   | *    | ○    | 　   | 　   |
| 北川 | ×,×  | 　   | ×    | *    | 　   | 　   |
| 伊藤 | ○    | ○    | 　   | 　   | *    | 　   |
| 山下 | 　   | 　   | 　   | 　   | 　   | *    |
