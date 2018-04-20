import win32gui
import win32api
import win32con
import win32ui
from PIL import Image, ImageStat
from time import sleep
import requests
import json


def get_screenshot():    
    hwnd = win32gui.GetDesktopWindow()
    
    # get complete virtual screen including all monitors
    SM_XVIRTUALSCREEN = 76
    SM_YVIRTUALSCREEN = 77
    SM_CXVIRTUALSCREEN = 78
    SM_CYVIRTUALSCREEN = 79
    w = vscreenwidth = win32api.GetSystemMetrics(SM_CXVIRTUALSCREEN)
    h = vscreenheigth = win32api.GetSystemMetrics(SM_CYVIRTUALSCREEN)
    l = vscreenx = win32api.GetSystemMetrics(SM_XVIRTUALSCREEN)
    t = vscreeny = win32api.GetSystemMetrics(SM_YVIRTUALSCREEN)
#    r = l + w
#    b = t + h
    
    hwndDC = win32gui.GetWindowDC(hwnd)
    mfcDC  = win32ui.CreateDCFromHandle(hwndDC)
    saveDC = mfcDC.CreateCompatibleDC()
     
    saveBitMap = win32ui.CreateBitmap()
    saveBitMap.CreateCompatibleBitmap(mfcDC, w,h)
    saveDC.SelectObject(saveBitMap)
    saveDC.BitBlt((0, 0), (w, h),  mfcDC,  (l, t),  win32con.SRCCOPY)
    
    bmpinfo = saveBitMap.GetInfo()
    bmpstr = saveBitMap.GetBitmapBits(True)
    im = Image.frombuffer(
        'RGB',
        (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
        bmpstr, 'raw', 'BGRX', 0, 1)
    
    # Free Resources
    saveDC.DeleteDC()
    mfcDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwndDC)
    win32gui.DeleteObject(saveBitMap.GetHandle())
    
    return im
    
    
def get_color_means(im):
    crop_height = 250 
    im_width = im.size[0]
    im_height = im.size[1]
    crops = []
    color_means=[]
    for i in range(0,6):
        crops.append ( im.crop((i*im_width/6, im_height-crop_height , (i+1)*im_width/6, im_height)) )
        #crops[i].save('crop' + str(i) + '.jpg')
        current_median =  ImageStat.Stat(crops[i]).mean
        bgr = int(current_median[2]); #red 0 - blue 2
        bgr = (bgr << 8) + int(current_median[1]); #green 1
        bgr = (bgr << 8) + int(current_median[0]); #blue 2 - red 0
        color_means.append(bgr)
    
    return color_means
 


def get_session():
    body = """{
        "title": "RazerLight",
        "description": "This is like ambilight, I built this for razer chroma v2",
        "author": {
            "name": "Gonzalo Piotti",
            "contact": "gonzalopiotti@hotmail.com"
        },
        "device_supported": [
            "keyboard"],
        "category": "application"
    }"""
    data = json.loads(body)
    
    r_init = requests.post('http://localhost:54235/razer/chromasdk', json = data)
    response = r_init.json()
    
    return  response['sessionid']
  
def create_effect(sessionid, color_means):
    keys=build_keys(color_means)
    
    create_url = 'http://localhost:' + str(sessionid) +'/chromasdk/keyboard'
    body = """{
        "effect":"CHROMA_CUSTOM",
         "param":  """ + keys + """
        
    }"""
    data = json.loads(body)
    r_create = requests.post(create_url, json = data)
    response = r_create.json()
 
    return  response['id']

def activate_effect(sessionid, effect_id):
    activate_url = 'http://localhost:' + str(sessionid) +'/chromasdk/effect'
    body = '{    "id" : "'+ str(effect_id) + '"}'
    data = json.loads(body)
    requests.put(activate_url, json = data)
    #delete after activation
    requests.delete(activate_url, json = data)
    
    
    
def update_keyboard(sessionid):
    im = get_screenshot()
    
    color_means = get_color_means(im)
    effectid = create_effect(sessionid, color_means)
    activate_effect(sessionid, effectid)

def build_keys(color_means):
    keys = []    
    
    for i in range(0,6):
        temp = []
        for i in range (0,22):
            
            if i <=3 :
                temp.append(color_means[0])
            elif i <=7:
                temp.append(color_means[1])
            elif i<=11:
                temp.append(color_means[2])
            elif i<=15:
                temp.append(color_means[3])
            elif i<=19:
                temp.append(color_means[4])
            elif i<=21:
                temp.append(color_means[5])
        keys.append(temp)
       
    return json.dumps(keys)


i=0
sessionid = get_session()
print('Running...(hit CTRL+C to stop)')
while True:

    update_keyboard(sessionid)
    #print('updated %s times' % i)
    sleep(0.01)
    i+=1



