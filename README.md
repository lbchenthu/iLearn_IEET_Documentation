# iLearn_IEET_Documentation
Generate IEET documents from iLearn Grade Report

1. 安裝相關套件 (Python 3.9)
pip install pandas numpy python-docx openpyxl

2. 
#iLearn匯出成績excel檔 (這個通常不會有問題)，
#Moodle用複製貼上產生excel檔 ，

#Excel格式參考範例:
IEET_score_list_108520_3.xlsx (注意紅字部分)

3. 編輯適當欄位，執行
python IEET_CG_SCRIPT.py IEET_score_list_108520_3.xlsx

輸出資料夾:
IEET_OUTPUT_IEET_score_list_108520_3/

4. 手動填寫其他欄位

