# sns-images

SNS投稿用の画像置き場(自動生成)。

## 画像の作り方(事業Bの定期実行が使う)
- `python3 tools/render.py spec.json images/<アカウント>/<投稿ID>.png`
- spec の書き方は `tools/render.py` の先頭のコメントを参照(型: line / memo / check)
- 作った画像は `images/<アカウント>/` に置き、push 後の URL
  `https://raw.githubusercontent.com/however1006/sns-images/main/images/<アカウント>/<投稿ID>.png`
  を投稿キューの U列「画像URL」に入れる
- 作り話のラベルは付けない(11/5 ユーザー判断)。ただし運営者本人の体験やスクショに見せない(よくある場面の再現として描く)
