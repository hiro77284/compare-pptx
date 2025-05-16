# Comparing two PowerPoint pptx files to find similar slides.

This Python tool compares two PowerPoint pptx files, searching similar slides based on two criteria: image similarity and text similarity.

This helps you when you have many subtly different pptx files as a result of repeated fine-tuning of a pptx.

## How to use

### install required packages

`pip install ImageHash numpy scikit-learn comtypes Pillow sentence-transformers`


PowerPointファイルに自動で章・節番号を振るツールです。しかその他にも目次作成、相互参照、配布用資料に載せない情報の自動削除など、必要に迫られてさまざまな機能を実装しました。

### PowerPointに章節番号を振る手間を自動化したい

PowerPointにはMS-Wordのようなアウトライン機能が無いので、各スライドに下記のように章・節番号をつけようとすると手作業になり非常に手間がかかります。研修屋さんの仕事では巨大なpptxを使うことが多く、このようなナンバリングが必須なため悩みの種でした。特に、一度付番してからページの追加・削除・入れ替え等があると番号の振り直しが必要で、単純作業ですがミスも起こりやすく、やってられません。これを自動化したいわけです。
