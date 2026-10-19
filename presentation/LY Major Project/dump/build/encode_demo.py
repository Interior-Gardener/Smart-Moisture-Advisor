from pathlib import Path
import json, math, bisect
import cv2
import numpy as np
from PIL import Image,ImageDraw,ImageFont

B=Path(__file__).resolve().parent
D=json.loads((B/'demo_capture.json').read_text())
OUT=B.parent.parent/'SMA_Major_Project_Demo.mp4'
W,H=1920,1080;FPS=6
def ft(n,bold=False):return ImageFont.truetype('C:/Windows/Fonts/arialbd.ttf' if bold else 'C:/Windows/Fonts/arial.ttf',n)
def wrapped(d,txt,font,width):
 lines=[];cur=''
 for word in txt.split():
  n=(cur+' '+word).strip()
  if d.textlength(n,font=font)>width and cur:lines.append(cur);cur=word
  else:cur=n
 if cur:lines.append(cur)
 return lines
def center(d,txt,y,size,bold=False,color='white'):
 f=ft(size,bold);d.text(((W-d.textlength(txt,font=f))/2,y),txt,font=f,fill=color)
def titlecard(ending=False):
 im=Image.new('RGB',(W,H),'#10372a');d=ImageDraw.Draw(im)
 d.rectangle((140,180,W-140,190),fill='#61d397')
 center(d,'Smart Moisture Advisor',265,78,True)
 center(d,'Irrigation decision support for sugarcane',390,42)
 center(d,'A demonstration of the running application' if not ending else 'Field validation is the next step',490,37,color='#b9e2cf')
 center(d,'Kartik Verma and Kushal Soni',675,36,True)
 center(d,'Project Guide: Dr. Nilesh Yadav',735,32)
 center(d,'Scenario estimates use agronomy-generated model labels',885,28,color='#b9e2cf')
 return cv2.cvtColor(np.array(im),cv2.COLOR_RGB2BGR)
writer=cv2.VideoWriter(str(OUT),cv2.VideoWriter_fourcc(*'mp4v'),FPS,(W,H))
assert writer.isOpened(),'MP4 video encoder unavailable'
for _ in range(8*FPS):writer.write(titlecard())
frames=D['frames'];times=[f['t']/1000 for f in frames];duration=D['durationMs']/1000
cache_idx=None;cached=None
for i in range(math.ceil(duration*FPS)):
 t=i/FPS;idx=max(0,bisect.bisect_right(times,t)-1);entry=frames[idx]
 if idx!=cache_idx:
  im=Image.new('RGB',(W,H),'#0c221b');d=ImageDraw.Draw(im)
  center(d,entry['title'],12,32,True)
  screen=Image.open(B/'demo_frames'/entry['file']).convert('RGB');im.paste(screen,(240,62))
  ls=wrapped(d,entry['caption'],ft(27),1780)
  for n,line in enumerate(ls):center(d,line,974+n*33,27)
  cached=cv2.cvtColor(np.array(im),cv2.COLOR_RGB2BGR);cache_idx=idx
 writer.write(cached)
for _ in range(8*FPS):writer.write(titlecard(True))
writer.release()
cap=cv2.VideoCapture(str(OUT));count=int(cap.get(cv2.CAP_PROP_FRAME_COUNT));fps=cap.get(cv2.CAP_PROP_FPS);ok=True;samples=[]
for sec in [0,8,25,45,65,85,105,125,145]:
 cap.set(cv2.CAP_PROP_POS_MSEC,sec*1000);good,frame=cap.read();ok=ok and good
 if good:
  p=B/f'video_check_{sec:03}.png';cv2.imwrite(str(p),frame);samples.append(str(p))
cap.release()
v={'path':str(OUT),'seconds':count/fps,'frames':count,'fps':fps,'dimensions':[W,H],'codec':'MPEG-4 Part 2 / mp4v','audio':'Captions only','capture':'Real headless Chrome interactions with the unchanged local application','sample_decode_checks_passed':ok,'browser_errors':D['errors'],'sample_images':samples}
(B/'video_validation.json').write_text(json.dumps(v,indent=2));print(json.dumps(v,indent=2))
