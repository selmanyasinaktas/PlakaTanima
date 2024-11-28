import cv2
import os
import pytesseract
from PIL import Image
import numpy as np
import xlsxwriter

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
harcascade = "haarcascade_russian_plate_number.xml"

cap = cv2.VideoCapture(0)

if not cap.isOpened():
    print("Kamera açılamadı")
    exit()

cap.set(3, 640)  
cap.set(4, 480)  

min_area = 500
count = 0
row = 0
column = 0

workbook = xlsxwriter.Workbook('Okunan_Plakalar.xlsx')
worksheet = workbook.add_worksheet()

plate_cascade = cv2.CascadeClassifier(harcascade)

if not os.path.exists("plates"):
    os.makedirs("plates")

while True:
    success, img = cap.read()
    if not success:
        print("Görüntü alınamadı")
        break

    img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    plates = plate_cascade.detectMultiScale(img_gray, 1.1, 4)

    img_roi = None
    for (x, y, w, h) in plates:
        area = w * h
        if area > min_area:
            cv2.rectangle(img, (x, y), (x + w, y + h), (0, 255, 0), 2)
            cv2.putText(img, "Plaka", (x, y - 5), cv2.FONT_HERSHEY_COMPLEX_SMALL, 1, (255, 0, 255), 2)
            img_roi = img[y: y + h, x: x + w]
            cv2.imshow("ROI", img_roi)

    cv2.imshow("Sonuç", img)

    key = cv2.waitKey(1) & 0xFF
    if key == ord('s'):
        if img_roi is not None:
            cv2.imwrite("plates/scanned_img_" + str(count) + ".jpg", img_roi)
            cv2.rectangle(img, (0, 200), (640, 300), (0, 255, 0), cv2.FILLED)
            cv2.putText(img, "Plaka Kaydedildi", (150, 265), cv2.FONT_HERSHEY_COMPLEX_SMALL, 2, (0, 0, 255), 2)
            cv2.imshow("Sonuç", img)
            cv2.waitKey(500)
            
            image_path = "plates/scanned_img_" + str(count) + ".jpg"
            config_turkish = r'-l tur --oem 3 --psm 9 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPRSTUVYZ'
            plate_img = cv2.imread(image_path)
            plate_img_resized = cv2.resize(plate_img, (600, 200))
            plate_gray = cv2.cvtColor(plate_img_resized, cv2.COLOR_BGR2GRAY)
            plate_blurred = cv2.GaussianBlur(plate_gray, (5, 5), 0)
            plate_text = pytesseract.image_to_string(plate_blurred, config=config_turkish)
            print("Algılanan Plaka Metni:", plate_text)
            
            worksheet.write(row, column, plate_text)
            count += 1
            row += 1
        else:
            print("Plaka bulunamadı, kaydedilecek bir şey yok")
    elif key == 27:  
        workbook.close()
        print("ESC tuşuna basıldı, çıkış yapılıyor.")
        break

cap.release()
cv2.destroyAllWindows()
