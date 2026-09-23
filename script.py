
import cv2
import numpy as np

# 1.读取原图
img = cv2.imread(r"F:\seele\8D955976215C036F5273844FFB2C87AD.png")
gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
cv2.imshow("original", img)

# 2.图像滤波
blur = cv2.blur(img, (5, 5))          #均值滤波9
gauss = cv2.GaussianBlur(img,(5,5),0) #高斯滤波
median = cv2.medianBlur(img, 5)      #中值滤波
cv2.imshow("gauss_filter", gauss)

# 3.对数变换、伽马(指数)变换
log_img = np.uint8(25 * np.log(1 + gray))
gamma_val = 1.2
gamma_img = np.uint8(255 * (gray / 255) ** gamma_val)
cv2.imshow("log_transform", log_img)
cv2.imshow("gamma_transform", gamma_img)

#4.二值化 + Canny边缘检测
ret, binary = cv2.threshold(gray, 127, 255, cv2.THRESH_BINARY)
edge = cv2.Canny(gray, 50, 150)
cv2.imshow("binary", binary)
cv2.imshow("canny_edge", edge)

#5.图像分割 轮廓绘制
contours, hierarchy = cv2.findContours(binary, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
cv2.drawContours(img, contours, -1, (0,0,255), 2)
cv2.imshow("segment_contour", img)


cv2.waitKey(0)         #按任意键关闭全部窗口
cv2.destroyAllWindows()