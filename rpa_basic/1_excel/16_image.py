#pip3 install Pillow해야 기능 사용 가능
from openpyxl import Workbook
from openpyxl.drawing.image import Image
wb = Workbook()
ws = wb.active

img = Image("img.png") # 같은 폴더 내에 있는 이미지

# C3 위치에 img.png 파일의 이미지를 삽입
ws.add_image(img, "C3")
wb.save("sample_imgae.xlsx")