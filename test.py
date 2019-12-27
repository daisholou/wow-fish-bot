import dlib
import PIL
import cv2

vidcap = cv2.VideoCapture('testcase.mov')
success,image = vidcap.read()
detector = dlib.simple_object_detector("detector.svm")
win = dlib.image_window()
count = 0
while success:
    dets = detector(image, 1)
    print("Number of faces detected: {}".format(len(dets)))
    win.clear_overlay()
    win.set_image(image)
    if len(dets) > 0:
        for i, d in enumerate(dets):
            print("Detection {}: Left: {} Top: {} Right: {} Bottom: {}".format(
                i, d.left(), d.top(), d.right(), d.bottom()))
        win.add_overlay(dets)
    success,image = vidcap.read()
    print('Read a new frame: {}'.format(count), success)
    count += 1

