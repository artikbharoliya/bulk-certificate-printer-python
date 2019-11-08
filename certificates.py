from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw
import cv2
import img2pdf,xlrd

x_cord= 0
y_cord= 0
texts = []
loc = ("names.xlsx")
#Click event 

#function for opening excel sheet----------------

def excel_read(location):
	global texts

	
	wb = xlrd.open_workbook(loc)
	sheet = wb.sheet_by_index(0)
	
	for i in range(1,sheet.nrows):
		texts.append(sheet.cell_value(i, 0))

	return texts




def click_and_crop(event, x, y, flags, param):
	#references to the global variables
	global x_cord,y_cord

	
	if event == cv2.EVENT_LBUTTONDOWN:
		x_cord,y_cord =x, y        


#function for the opening the image and getting cords------------------
def certi_open(path):

	image = cv2.imread(path)

	screen_res = 1280, 720
	scale_width = screen_res[0] / image.shape[1]
	scale_height = screen_res[1] / image.shape[0]
	scale = min(scale_width, scale_height)

	window_width = int(image.shape[1] * scale)
	window_height = int(image.shape[0] * scale)

	cv2.namedWindow('Resized Window', cv2.WINDOW_NORMAL)

	cv2.resizeWindow('Resized Window', window_width, window_height)

	cv2.imshow('Resized Window', image)
	cv2.setMouseCallback("Resized Window", click_and_crop)


	#opening the image for coordinates------------------------------------

	while True:


		key = cv2.waitKey(27)

		if key == 27 or (x_cord and y_cord):
			break

	cv2.destroyAllWindows()

	print_it(x_cord, y_cord)


	#Printing on the certificate------------------------------------
def print_it(x_cord, y_cord):
	global loc
	texts = excel_read(loc)
	selectFont = ImageFont.truetype("Calibri.ttf", size = 60)
	for text in texts:
		img = Image.open("template.jpg")

		draw = ImageDraw.Draw(img)

		draw.text( (x_cord,y_cord), text, (255,255,255), font=selectFont)
		certi_name = "certificatess/"+"certi_"+text+".pdf"
		img = img.convert('RGB')
		img.save(certi_name, "PDF")
	print("Successsssss")


path = "Computer Diploma Certificate.jpg"
certi_open(path)
