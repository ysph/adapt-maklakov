import cv2
import numpy as np
                
def findCnts(img):
    ref,thresh_value = cv2.threshold(img,230,255,cv2.THRESH_BINARY_INV)
    kernel = np.ones((1,1),np.uint8)
    dilated_value = cv2.dilate(thresh_value,kernel,iterations=1)
    contours,hierarchy = cv2.findContours(dilated_value,cv2.RETR_TREE,cv2.CHAIN_APPROX_SIMPLE)
    return contours
    
def FindData(frames):
    result = {}
    for i, frame in enumerate(frames):
        frRGB = cv2.cvtColor(frame, cv2.COLOR_GRAY2BGR)
        
        cnts = findCnts(frame)
        prev = -1
        for c in cnts:
            x,y,w,h = cv2.boundingRect(c)
            if w > prev:
                prev = w
                x1,y1,w1,h1 = x,y,w,h
        if h1/w1 < 0.48:
            result[i+1] = '-'
        else:
            result[i+1] = '+'
    return result
    
def FindTables(path,rows,cols):
    # opening images this way cause library cannot read cyrillic
    stream = open(path,'rb')
    bytes = bytearray(stream.read())
    numarr = np.asarray(bytes)
    im1 = cv2.imdecode(numarr,cv2.IMREAD_GRAYSCALE)
    im = cv2.imdecode(numarr,cv2.IMREAD_UNCHANGED)

    if im1.shape[0] > im1.shape[1]:
        im1 = np.ascontiguousarray(np.rot90(im1,k=1))
        im = np.ascontiguousarray(np.rot90(im,k=1))
    
    contours = findCnts(im1)
    rects = list()
    tables = list()
    for c in contours:
        x,y,w,h = cv2.boundingRect(c)
        if w > im1.shape[1] / 1.2:
            tables.append((x,y,w,h))
        else:
            rects.append((x,y,w,h))
    tables_cells = {}
    put_image = {}
    for i,table in enumerate(tables,1):
        tables_cells[i] = list()
        j = 1
        for r in rects:
            x0,y0,w0,h0 = table
            x,y,w,h = r
            if x>=x0 and x+w <= x0+w0 and y>=y0 and y+h <= y0+h0:
                width = w0 / cols
                if w < (width * 2) and w > (width / 1.5):
                    tables_cells[i].append((x,y,w,h))
                    j += 1
        cells = []
        sorted_cells = sorted(tables_cells[i],key=lambda ff: ff[1])

        for j, cell in enumerate(sorted_cells, 1):
            x,y,w,h = cell
            if j <= (cols * rows):
                cells.append(cell)
                lx,ly,lw,lh = cell
        res_cells = []
        k = 1
        for step in range(0,cols*rows+1,cols):
            for x,y,w,h in sorted(cells[step-cols:step],key=lambda ff: ff[0]):
                if k > 165:
                    break
                res_cells.append(im1[y+5:y+h-5,x+5:x+w-5])
                k += 1
        put_image[i] = {'image':im[sorted_cells[0][1]:ly+(lh*3),x0:x0+w0],'cells':res_cells}
    
    return put_image
