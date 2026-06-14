!!!!!注意!!!!!
開発環境(Javaインストール環境)でのみ動く

■以下を同階層に配置する
・generateTable(フォルダ)
・SqlGenerator.bat
・SqlGenerator.jar

■フォルダ[generateTable]内にテーブル情報を書いたテキストファイルを配置する
　→ファイル名は[テーブル物理名.txt]にする(拡張子は何でもいい)
　→ファイル内容はA5:SQLのタブ[カラム]をコピペで作成可能

テーブル情報テキストファイル内容例
[mst_user.txt]の場合
==================================================
ユーザマスタ
user_id
ユーザID	user_id	character varying(100)
パスワード	password	character varying(100)
ユーザ名	user_name	character varying(100)
追加ユーザID	insert_user_id	character varying(100)
追加日時	insert_datetime	timestamp without time zone
追加ユーザID	update_user_id	character varying(100)
追加日時	update_datetime	timestamp without time zone
削除フラグ	delete_flg	character varying(1)
==================================================

解説
タブ区切りで項目判別しているので、余計なタブは入れないようにする
空行も入れないようにする

・1行目に[テーブル論理名]を書く
・2行目にPK(物理名)を1つだけ書く
　→WHEREの1行目に "PK IS NOT NULL" を書くため
・3行目にタブ区切りで[項目論理名(\t)項目物理名(\t)項目型]を書く
　→型は先頭文字が[chara/integer/timestamp/date/bytea]または末尾文字が[[]]で判別しているので
　　配列型以外は後ろに余計な文字がついていてもOK
　　→配列型は末尾に[]があればOK

■SqlGenerator.batを実行する

■フォルダ[generateEntity, generateMapper, generateXml]が作成される