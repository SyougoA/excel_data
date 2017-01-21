# coding: utf-8
import openpyxl

file = "office_data.xlsx"
#fileを開く
book = openpyxl.load_workbook(file)
sheet = book.worksheets[0]

data = []
#rowだけど列を意識した動きなイメージ
for row in sheet.rows:
	data.append([
		row[0].value,
		row[5].value
		])

#説明要素の0~4行目まで削除する
#このdataは0, 6列目から2つの要素を持つリストを作っている
del data[0:4]
#上のままだとNoneも含めてしまうからdatasetとして整える。上のコードとダブるのでそこは使い分け。
data_set = data[0:78]

#昇順ソートする
data_set = sorted(data_set, key = lambda x:x[1])
#降順にしたい場合はreverse関数を使用
data_set.reverse()

#data_setで数値10以上のものを表記する
for i , sort in enumerate(data_set):
	if(sort[1]<10): break
	print(i+1, sort[0], int(sort[1]))