# DataMigration

## Excel Data to MySQL
20만개의 식품안전나라에 등록된 식품영양성분 데이터와 그 외의 기업에서 제공하는 영양성분 데이터, 해외 식품영양성분 데이터를 MySQL 데이터베이스에 저장하기 위해 만든 프로그램입니다.

데이터 파일 형식은 .xlsx이어야 합니다.

## Requirements
```
pip install mysql-connector-python openpyxl tqdm
```
*mysql-connector-python
    MySQL 데이터베이스에 접근할 수 있게 해주는 라이브러리입니다. 이 라이브러리를 사용하면, Python 코드에서 MySQL 데이터베이스와 상호작용하면서 데이터를 조회, 삽입, 수정, 삭제할 수 있습니다.

*openpyxl
    Excel xlsx/xlsm/xltx/xltm 파일을 읽고 쓸 수 있게 해주는 라이브러리입니다. 이 라이브러리를 사용하면, Python 프로그램에서 Excel 파일을 생성하고, 기존 파일을 읽고, 데이터를 수정할 수 있습니다.

*tqdm
    빠르고 확장 가능한 진행률 바 라이브러리입니다. "taqaddum" (تقدّم)에서 이름이 유래되었으며, 이는 "진행"을 뜻하는 아랍어입니다. tqdm을 사용하면, for 루프나 다른 반복 작업의 진행 상황을 시각적으로 표시할 수 있습니다. 그래서 사용자는 작업이 얼마나 진행되었는지, 얼마나 남았는지, 예상 완료 시간은 언제인지를 쉽게 확인할 수 있습니다.

## 주요내용
### 1. 코드
```
for row in tqdm(ws.iter_rows(min_row=2, values_only=True), total=ws.max_row - 1):
    category_query = "INSERT IGNORE INTO Category (name) VALUES (%s)"
    category_data = (row[food_idx[1]],)

    brand_query = "INSERT IGNORE Brand (name) VALUES (%s)"
    brand_data = (row[food_idx[2]],)

    cursor.execute(category_query, category_data)
    cursor.execute(brand_query, brand_data)

    cursor.execute("SELECT id FROM Category WHERE name = %s", (row[food_idx[1]],))
    food_data[0] = cursor.fetchone()[0]

    cursor.execute("SELECT id FROM Brand WHERE name = %s", (row[food_idx[2]],))
    food_data[1] = cursor.fetchone()[0]

    div = float(row[food_idx[0]]) / 100

    if row[food_idx[3]] == '-':
        food_data[2] = None
    else: 
        food_data[2] = row[food_idx[3]]

    for index, nut_idx in enumerate(food_idx[4:], start=3):
        if nut_idx is not None:
            food_data[index] = row[nut_idx]
            if is_number(food_data[index]) == False:
                food_data[index] = None
            else:
                food_data[index] = float(row[nut_idx]) / div

    food_query = """INSERT INTO Food (category_id, brand_id, name, energy, protein, fat, carbohydrate, sugar, sodium, 
                cholesterol, saturate_fat, trans_fat)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)"""

    cursor.execute(food_query, food_data)

conn.commit()
cursor.close()
conn.close()
```
### 2. 세부설명
