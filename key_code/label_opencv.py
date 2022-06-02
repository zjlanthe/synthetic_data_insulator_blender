import cv2
import os
import time
import xlwt  # pip install xlwt  需安装

Fi0 = "/home/lanthe/DL/dataset/Synthetic_dataset/sysdata-001_jpg/Annotations/"
img_folder = "/home/lanthe/DL/dataset/Synthetic_dataset/sysdata-001_jpg/Image_no_Annotations"
# excel_folder = "/home/lanthe/DL/dataset/Synthetic_dataset/sysdata-001_jpg/"

# savepath = excel_folder + 'label.xls'
img_file = [os.path.join(img_folder, f) for f in os.listdir(img_folder) if f.endswith(".png")]

print("------------------")
print(Fi0)
print(img_folder)
# print(excel_folder)
# print(savepath)
print(img_file)
print("------------------")
num = 0

# book = xlwt.Workbook(encoding='utf-8', style_compression=0)
# sheet = book.add_sheet('01', cell_overwrite_ok=True)

for folder in img_file:
    num += 1
    start = time.time()
    image = cv2.imread("/home/lanthe/DL/dataset/Synthetic_dataset/sysdata-001_jpg/Image_no_Annotations/" + str(num) + '.png')
    jpgi = cv2.imread("/home/lanthe/DL/dataset/Synthetic_dataset/sysdata-001_jpg/JPEGImages/" + str(num) + '.jpg')

    W = image.shape[1]
    H = image.shape[0]
    # print(image.shape)  
    # print(image.shape[0]) # h
    # print(image.shape[1]) # w

    # cv2.imshow("winname", image)

    # 对图像进行阈值分割，阈值设定为160，得到二值化灰度图
    image1 = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    ret, image1 = cv2.threshold(image1, 195, 255, cv2.THRESH_BINARY)
    image1 = cv2.bitwise_not(image1)
    #cv2.imshow('grayscale', image1)
    
    # 获取二值图像中绝缘子的轮廓及坐标
    image00, contours, hierarchy = cv2.findContours(image1, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    for cidx, cnt in enumerate(contours):
        (x, y, w, h) = cv2.boundingRect(cnt)
        print('RECT: x={}, y={}, w={}, h={}'.format(x, y, w, h))
    cv2.rectangle(jpgi, (x, y), (x+w, y+h), (0, 0, 255), 2)
    # cv2.imshow('end', jpgi)
    
    # cv2.imwrite(os.path.join(img_folder, '2.' + str(num) + '.jpg'), jpgi)
    # cv2.waitKey(1000)
    # end = time.time()
    # sheet.write(num, 1, (end-start))
    # book.save(savepath)

    # Finally calculating and storing the results of the data
#for i in range(len(idxs)):
    file = open("/home/lanthe/DL/dataset/Synthetic_dataset/sysdata-001_jpg/Annotations/" +
                str(num)+".xml", "w")
    s = '''<annotation>
    <folder>'''+ str(folder) +'''</folder>
    <filename>''' + str(num)+'''</filename>
    <path>''' + str(Fi0)+'''/'''+str(num)+''''.jpg</path>
    <source>
        <database>Unknown</database>
    </source>
    <size>
        <width>'''+str(W)+'''</width>
        <height>'''+str(H)+'''</height>
        <depth>3</depth>
    </size>'''


    s += '''
    <object>
        <name>ceramicinsulators</name>
        <pose>Unspecified</pose>
        <truncated>0</truncated>
        <difficult>0</difficult>
        <bndbox>
            <xmin>'''+str(x)+'''</xmin>
            <ymin>'''+str(y)+'''</ymin>
            <xmax>'''+str(x+w)+'''</xmax>
            <ymax>'''+str(y+h)+'''</ymax>
        </bndbox>
    </object>'''

    s+='''
    </annotation>'''
    file.write(s)
    file.close()

