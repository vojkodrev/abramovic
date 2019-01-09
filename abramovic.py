
import string

import os

import win32gui
import win32api
import win32con

from time import sleep
from time import time

import numpy as np
import cv2

import hashlib

from random import randrange, choice, random

from pyautogui import \
  typewrite, click, moveTo, press, mouseDown, mouseUp,\
  keyDown, keyUp, moveRel, position, screenshot

from threading import Thread, Timer

class MoveMouseRandomlyThread(Thread):
  def __init__(self, deviation):
    Thread.__init__(self)
    self.daemon = True
    self.terminated = False
    self.deviation = deviation
    self.startingPosition = position()

  def run(self):
    while not self.terminated:
      moveTo(self.startingPosition[0] + randrange(-self.deviation, self.deviation), \
        self.startingPosition[1] + randrange(-self.deviation, self.deviation),\
        0.2)

  def stop(self):
    self.terminated = True

def focusWindow(hwnd, x, y):
  win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
  win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE + win32con.SWP_NOSIZE)  
  win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE + win32con.SWP_NOSIZE)  
  win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_SHOWWINDOW + win32con.SWP_NOMOVE + win32con.SWP_NOSIZE)

  click(x = x + 150, y = y + 10)

def exitIfMercenaryIsNotPresent(x, y, w, h):
  pt = matchWindowAndTemplate("images/mercenary.png", x, y, w, h, 0.7)
  if pt == None:
    print("no mercenary") 
    os._exit(1)    

def exitIfDead(x, y, w, h):
  print("detecting death")

  images = [
    'images/dead.png',
    'images/dead2.png',
    'images/dead3.png',
    'images/dead4.png',
    'images/dead5.png',
    'images/dead6.png',
  ]

  screenshot = takeScreenshot(x, y, w, h)

  for i in images:
    pt = matchWindowAndTemplate(i, x, y, w, h, 0.7, screenshot)
    if pt != None:
      print("death detected", i) 
      click()     
      os._exit(1)
      break

def gotoPortal(x, y, w, h):
  moveTo(x + 308 + randrange(-15, 0), y + 438 + randrange(0, 15), 0.2)
  mouseDown()
  t = MoveMouseRandomlyThread(40)
  t.start()
  sleep(5.5) #sleep(4.7)
  t.stop()  

  moveTo(x + 300 + randrange(-10, 0), y + 311 + randrange(0, 20), 0.2)
  t = MoveMouseRandomlyThread(40)
  t.start()
  sleep(0.5)
  mouseUp()
  t.stop()  
  sleep(0.2)

def takeScreenshot(x, y, w, h):
  return np.array(screenshot(region=(x, y, w, h)))

def matchWindowAndTemplate(templateFilename, x, y, w, h, treshold = 0.8, screenshot = None):
  if screenshot is None:
    screenshot = takeScreenshot(x, y, w, h)
  
  # screenshot = np.array(screenshot)

  template = cv2.imread(templateFilename, 0)
  image = cv2.cvtColor(screenshot, cv2.COLOR_RGB2GRAY)
  res = cv2.matchTemplate(image, template, cv2.TM_CCOEFF_NORMED)
  loc = np.where(res >= treshold)
  loc = zip(*loc[::-1])
  result = next(loc, None)

  if result == None:
    filename = 'errors/%s-%s.png' % (os.path.splitext(os.path.basename(templateFilename))[0], currentTimeMillis())
    cv2.imwrite(filename, cv2.cvtColor(screenshot, cv2.COLOR_RGB2BGR))

  return result

def matchWindowAndTemplateWithRetries(templateFilename, x, y, w, h, treshold = 0.8, retries = 3):
  pt = matchWindowAndTemplate(templateFilename, x, y, w, h, treshold)
  
  if pt == None and retries > 0:
    print("retrying", templateFilename)
    sleep(0.5)
    return matchWindowAndTemplateWithRetries(templateFilename, x, y, w, h, treshold, retries - 1)

  return pt    

def clickOnPortal(x, y, w, h, retries = 1):
  pt = matchWindowAndTemplate('images/portal.png', x, y, w, h, treshold = 0.6)

  if pt != None:
    moveTo(x + pt[0] + 55 + randrange(-10, 10), y + pt[1] + 80 + randrange(-10, 10), 0.1)
    click()
    sleep(1.5)
    return True
  elif retries > 0:  
    print("portal was not found trying unstuck light")
    pt = matchWindowAndTemplate('images/unstuck light.png', x, y, w, h, treshold = 0.6)
    if pt != None:
      moveTo(x + pt[0] + 57 + randrange(-10, 10), y + pt[1] + 83 + randrange(-10, 10), 0.2)
      mouseDown()
      t = MoveMouseRandomlyThread(20)
      t.start()
      sleep(0.5)
      t.stop()
      mouseUp()
      sleep(0.8)
      return clickOnPortal(x, y, w, h, retries - 1)
    else:
      return False

def teleportToDoor(x, y, w, h):
  press('w')

  startX = x + 590 + randrange(-10, 50)
  startY = y + 124 + randrange(-10, 30)
  moveTo(startX, startY)

  for i in range(3):
    click(button = "right")
    moveTo(startX + randrange(-20, 20), startY + randrange(-20, 20))
    sleep(0.4)

  door = matchWindowAndMultipleTemplatesAndMoveMouse([
    ('images/door2.png', 0.5, 75 + randrange(-10, 10), 110 + randrange(-10, 10)),
    ('images/door light.png', 0.5, 80 + randrange(-10, 10), 116 + randrange(-10, 10)),
    ('images/door4.png', 0.5, 49 + randrange(-10, 10), 37 + randrange(-10, 10)),
    ('images/door5.png', 0.5, 108 + randrange(-10, 10), 35 + randrange(-10, 10)),
  ], x, y, w, h, 0, 0)

  click(button = "right")  

  sleep(0.6)

  return True

def switchWeapon():
  press('s')
  sleep(0.2 + randrange(200, 700) / 1000)

def currentTimeMillis():
  return int(round(time() * 1000))

def clickMouseForXMillis(millis):
  start = currentTimeMillis()

  while (currentTimeMillis() - start < millis):
    click(button = "right")
    moveRel(randrange(-10, 10), randrange(-10, 10))
    sleep(0.2)   

def getRandomText():
  r = range(0, randrange(5, 10))
  return "".join([choice(string.ascii_lowercase) for i in r])

# def pickUpPotIfNoneExist(x, y, w, h, type):
#   if getPotPositionInBelt(x, y, w, h, type) is None:
#     if type == "blue":
#       pickUpPot('images/mana pot white.png', x, y, w, h)
#     elif type == "red":
#       pickUpPot('images/health pot.png', x, y, w, h)
#   else:
#     print("not picking up", type, "because it is present in belt")      

def countNonesInArray(arr):
  count = 0
  for i in arr:
    if i is None:
      count += 1
  return count

def pickupPots(x, y, w, h):
  belt = getBeltContent(x, y)

  bluePots = 0
  redPots = 0
  # lastEmptyColumnWentTo = "red"

  for column in belt:
    if column[0] is None:
      bluePots += 4
      # if lastEmptyColumnWentTo == "blue":
      #   redPots += 4
      #   lastEmptyColumnWentTo = "red"
      # elif lastEmptyColumnWentTo == "red":
      #   bluePots += 4
      #   lastEmptyColumnWentTo = "blue"
    
    elif column[0] == "red":
      redPots += countNonesInArray(column)
    elif column[0] == "blue":
      bluePots += countNonesInArray(column)

  print("will pick up", bluePots, "blue pots and", redPots, "red pots")

  for _ in range(bluePots):
    if not pickUpPot('images/mana pot white.png', x, y, w, h):
      break
  
  for _ in range(redPots):
    if not pickUpPot('images/health pot.png', x, y, w, h):
      break

def inRangeBlue(hsv):
  return cv2.inRange(hsv, (115, 200, 20), (125, 255, 255))

def inRangeRed(hsv):
  maskRed1 = cv2.inRange(hsv, (175, 100, 20), (180, 255, 255))
  maskRed2 = cv2.inRange(hsv, (0, 100, 20), (5, 255, 255))
  return cv2.bitwise_or(maskRed1, maskRed2)

def getPotPositionInBelt(x, y, w, h, type):
  screenshot = takeScreenshot(x + 424, y + 590, 125, 33)
  # screenshot = np.array(screenshot)
  hsv = cv2.cvtColor(screenshot, cv2.COLOR_RGB2HSV)

  if type == "blue":
    mask = inRangeBlue(hsv)
  elif type == "red":
    mask = inRangeRed(hsv)

  # cv2.imwrite("t.png", mask)

  points = cv2.findNonZero(mask)

  if points is not None:
    return points[0][0][0]
  else:
    return None

def drinkPot(x, y, w, h, type):

  position = getPotPositionInBelt(x, y, w, h, type)

  if position is not None:
    numberToPress = int(position / 33) + 1
    print("drinking", type, "pot at", numberToPress)
    press(str(numberToPress))
  else:
    print("no", type, "pot found")

def killPindle(x, y, w, h):
  pt = matchWindowAndTemplateWithRetries('images/house middle2.png', x, y, w, h, 0.7)

  if pt == None:
    return False

  moveTo(x + pt[0] + 188 + randrange(-10, 10), y + pt[1] + 116 + randrange(-10, 10))

  press('a')

  clickMouseForXMillis(1300 + randrange(-100, 500))

  # drinkPot(x, y, w, h, "blue")

  # clickMouseForXMillis(1100 + randrange(-100, 500))

  # switchWeapon()

  press('q')

  clickMouseForXMillis(2000 + randrange(-500, 2500))    

  return True

def teleportToLoot():
  press('w')

  moveRel(randrange(-5, 5), randrange(-5, 5))

  click(button = "right")

def castShield():
  press('e')

  click(button = "right") 

  sleep(0.5)      

def createGame(x, y):
  moveTo(x + 589, y + 484)

  click()   

  t = getRandomText()
  print(t)
  typewrite(t)

  press("\t")	

  t = getRandomText()
  print(t)
  typewrite(t)

  press("\t")	

  typewrite("GS 6" if random() < 0.5 else "GS 9")

  moveTo(x + 706, y + 398)

  click() 

  moveTo(x + 671, y + 444)

  click() 

  moveTo(x + randrange(200, 600), y + randrange(100, 500))

  sleep(4.5)   

def exitGame(x, y, start):
  keyUp('alt')
  
  # switchWeapon()
  
  press('esc')
  moveTo(x + 272, y + 287)
  click()      

  sleep(5 + randrange(0, 3000) / 1000)

  print("total run", (currentTimeMillis() - start) / 1000, "s")

def pickUpPot(image, x, y, w, h):
  pt = matchWindowAndTemplate(image, x, y, w, h, 0.5)
  
  if pt != None:        
    print("moving to pot", image)
    press("d")
    moveTo(x + pt[0] + 20, y + pt[1] + 8)
    
    if random() < 0.3:
      click()
    else:
      click(button = "right")

    sleep(0.50 + randrange(-20, 20) / 1000)

    return True
  else:
    print("no pot", image) 
    return False

def inRangeGreen(hsv):
  return cv2.inRange(hsv, (45, 100, 20), (75, 255, 255)) 

class DrinkPotWhenLowOnResources(Thread):
  def __init__(self, x, y, w, h, type, rx, ry, rw, rh):
    Thread.__init__(self)
    self.daemon = True
    self.x = x
    self.y = y
    self.w = w
    self.h = h
    self.rx = rx
    self.ry = ry
    self.rw = rw
    self.rh = rh
    self.type = type
    self.terminated = False

  def stop(self):
    self.terminated = True    

  def run(self):
    print("drink pot when low on resources started", \
      self.x, self.y, self.w, self.h, \
      self.type, self.rx, self.ry, self.rw, self.rh)

    while not self.terminated:
      screenshot = takeScreenshot(self.x + self.rx, self.y + self.ry, self.rw, self.rh)
      hsv = cv2.cvtColor(screenshot, cv2.COLOR_RGB2HSV)

      if self.type == "red":
        maskRed = inRangeRed(hsv)
        maskGreen = inRangeGreen(hsv)
        mask = cv2.bitwise_or(maskRed, maskGreen)
        # cv2.imwrite("red.png", mask)
      elif self.type == "blue":
        mask = inRangeBlue(hsv)
        # cv2.imwrite("blue.png", mask)

      points = cv2.findNonZero(mask)

      if points is not None:
        startY = points[0][0][1]
        if startY > 15 and not self.terminated:
          print("resource", self.type, "too low drinking pot")  
          drinkPot(self.x, self.y, self.w, self.h, self.type)
          sleep(6)

      sleep(0.2)

def teleportToTownOrExit(x, y, w, h, start):
  keyUp("alt")

  pt = matchWindowAndTemplateWithRetries('images/light.png', x, y, w, h, 0.8)
  if pt != None:
    print("tp position found")
    moveTo(x + pt[0] + 28, y + pt[1] + 122, 0.2)
    click()
    sleep(0.7)
  else:
    print("tp position not found")
    exitGame(x, y, start)
    return 

  press('t')
  click(button = "right")

  sleep(2)

  tp = matchWindowAndMultipleTemplatesAndMoveMouse([
    ('images/tp.png', 0.6, 40, 60),
    ('images/tp2.png', 0.6, 20, 47),
    ('images/tp3.png', 0.6, 67, 62),
  ], x, y, w, h, 0.2)

  if tp:
    click()
  else:
    exitGame(x, y, start)

def matchWindowAndMultipleTemplatesAndMoveMouse(templates, x, y, w, h, interval = 0, retries = 3):
  for template in templates:
    pt = matchWindowAndTemplateWithRetries(template[0], x, y, w, h, template[1], retries)
    if pt != None:
      moveTo(x + pt[0] + template[2], y + pt[1] + template[3], interval)
      return True
    else:
      print(template[0], "failed")      
  return False

def pickUpItems(x, y, w, h, start):
  for i in range(10):
    sleep(0.50 + randrange(-20, 20) / 1000)

    pt = matchWindowAndTemplate('images/pickup.png', x, y, w, h, 0.5)
    
    if pt == None:
      print("no more items")
      break

    print("found an item!")
    moveTo(x + pt[0] + 20, y + pt[1] + 8)
    click()        

    if i == 9:
      print("inventory is full, tping")
      teleportToTownOrExit(x, y, w, h, start)
      return False

  # teleportToTownOrExit(x, y, w, h, start)
  # return False    

  return True

def closeTeamviewerDialog(x, y, w, h):
  pt = matchWindowAndTemplate("images/teamviewer.png", x, y, w, h)
  if pt != None:
    print("detected teamviewer")
    moveTo(x + pt[0] + 426, y + pt[1] + 125, 0.2)
    click()
    sleep(0.5)
  else:
    print("no teamviewer")  

def setBeltPots(mask, belt, type):
  points = cv2.findNonZero(mask)

  if points is not None:
    for row in range(4):
      for column in range(4):
        startX = row * 33
        stopX = (row + 1) * 33
        startY = (3 - column) * 33
        stopY = (4 - column) * 33
        # print("column", column, "row", row, "y", startY, "-", stopY, "x", startX, "-", stopX)
        for p in points:
          px = p[0][0]
          py = p[0][1]
          if startX <= px and px < stopX and \
            startY <= py and py < stopY:
            belt[row][column] = type
            break  

def getBeltContent(x, y):
  press("~")
  sleep(0.1)

  belt = [[None, None, None, None], [None, None, None, None], [None, None, None, None], [None, None, None, None]]
  
  bx = x + 422
  by = y + 492
  bw = 127
  bh = 133

  screenshot = takeScreenshot(bx, by, bw, bh)
  hsv = cv2.cvtColor(screenshot, cv2.COLOR_RGB2HSV)
  maskBlue = inRangeBlue(hsv)
  maskRed = inRangeRed(hsv)
  # cv2.imwrite("t.png", mask)

  setBeltPots(maskBlue, belt, "blue")
  setBeltPots(maskRed, belt, "red")

  print("belt content", belt)

  press("~")  

  return belt            

def onWindowsFound(hwnd, extra):
  rect = win32gui.GetWindowRect(hwnd)
  x = rect[0]
  y = rect[1]
  w = rect[2] - x
  h = rect[3] - y
  windowText = win32gui.GetWindowText(hwnd)

  if (windowText == "Diablo II"):
    focusWindow(hwnd, x, y)

    programStart = currentTimeMillis()
    runTime = randrange(45, 90) * 60 * 1000
    # runTime = 15 * 60 * 1000
    print("run time will be", runTime)

    while currentTimeMillis() - programStart < runTime:
      start = currentTimeMillis()

      createGame(x, y)

      dhpt = DrinkPotWhenLowOnResources(x, y, w, h, "red", 30, 535, 85, 85)
      dhpt.start()

      dmpt = DrinkPotWhenLowOnResources(x, y, w, h, "blue", 691, 535, 85, 85)
      dmpt.start()

      closeTeamviewerDialog(x, y, w, h)

      press('b')

      castShield()

      press('b')

      dt = Timer(1, exitIfDead, (x, y, w, h))
      dt.setDaemon(True)
      dt.start()
      
      gotoPortal(x, y, w, h)

      closeTeamviewerDialog(x, y, w, h)

      # exitIfMercenaryIsNotPresent(x, y, w, h)

      if not clickOnPortal(x, y, w, h):
        print("something went wrong while clicking on portal, exiting early")
        sleep(0.5)
        dhpt.stop()
        dmpt.stop()
        exitGame(x, y, start)
        continue
      
      if not teleportToDoor(x, y, w, h):
        print("something went wrong while teleporting to door, exiting early")
        dhpt.stop()
        dmpt.stop()
        sleep(0.5)
        exitGame(x, y, start)
        continue

      if not killPindle(x, y, w, h):
        print("something went wrong while killing Pindle, exiting early")
        dhpt.stop()
        dmpt.stop()
        sleep(0.5)
        exitGame(x, y, start)
        continue        

      teleportToLoot()

      keyDown('alt')

      sleep(2 + randrange(-500, 6000) / 1000)

      pickupPots(x, y, w, h)

      if not pickUpItems(x, y, w, h, start):
        print("stopping because inventory is full")
        dhpt.stop()
        dmpt.stop()
        break
      
      keyUp('alt')

      dhpt.stop()
      dmpt.stop()

      exitGame(x, y, start)

if __name__ == '__main__':
  win32gui.EnumWindows(onWindowsFound, None)
