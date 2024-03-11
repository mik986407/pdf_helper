簡介:

pdf_helper為一款由python製作的離線小軟體，含有pdf轉word檔、pdf及圖檔的文字辨識、pdf分割及合併等功能。

注意:
一、py檔為pdf_helper各功能運作邏輯。

二、ui檔為qt designer軟體產出的介面設計檔案。

詳細情形請參閱官網資訊:https://doc.qt.io/qt-6/qtdesigner-manual.html

三、OCR的相關資料源使用開源檔(僅下載辨識中英文的資料源)，因此不放在github。

如需使用OCR功能，請參照以下步驟

1.於當前環境安裝pytesseract
。pip install pytesseract

2.額外安裝圖片辨識文字功能的tesseract程式

。Windows:https://github.com/UB-Mannheim/tesseract/wiki

。MAC:輸入brew install tesseract

3.因預設只含英文，需額外下載語系包，參考連結:https://github.com/tesseract-ocr/tessdata_best

step1:下載特定語言包

。例:繁中(chi_tra.traineddata)、簡中(chi_sim.traineddata)、日文(jpn.traineddata)

step2:將語言包放至Tesseract-OCR的tessdata資料夾

。例:C:\Program Files (x86)\Tesseract-OCR\tessdata
