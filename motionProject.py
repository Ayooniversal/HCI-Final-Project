import sys
import thread
import time
import win32api
import Leap
import math
import os
import Tkinter

from Tkinter import *

from Leap import KeyTapGesture, SwipeGesture, ScreenTapGesture

# Assign key codes for Windows to variables
VK_ESCAPE = 0x1B
VK_MEDIA_PLAY_PAUSE = 0xB3
VK_MEDIA_NEXT_TRACK = 0xB0
VK_MEDIA_PREV_TRACK = 0xB1
VK_VOLUME_MUTE = 0xAD
VK_VOLUME_DOWN = 0xAE
VK_VOLUME_UP = 0xAF

#Assign virtual keys to variables
code = win32api.MapVirtualKey(VK_ESCAPE, 0)
code1 = win32api.MapVirtualKey(VK_MEDIA_PLAY_PAUSE, 0)
code2 = win32api.MapVirtualKey(VK_MEDIA_NEXT_TRACK, 0)
code3 = win32api.MapVirtualKey(VK_MEDIA_PREV_TRACK, 0)
code4 = win32api.MapVirtualKey(VK_VOLUME_MUTE, 0)
code5 = win32api.MapVirtualKey(VK_VOLUME_DOWN, 0)
code6 = win32api.MapVirtualKey(VK_VOLUME_UP, 0)


class LeapEventListener(Leap.Listener):


    # Printed on startup of execution
    def on_init(self, controller):
        print "Initialized"
        self.prevTime = time.time()
        self.minTime = 1
        
    # Upon connection, enable desired gestures to run in the background
    def on_connect(self, controller):
        print "Connected"

        controller.enable_gesture(Leap.Gesture.TYPE_SWIPE);
        controller.enable_gesture(Leap.Gesture.TYPE_KEY_TAP);
        controller.enable_gesture(Leap.Gesture.TYPE_SCREEN_TAP);
        controller.set_policy(Leap.Controller.POLICY_BACKGROUND_FRAMES);
        controller.config.save();

    def on_disconnect(self, controller):
        print "Motion Sensor Disconnected"

    def on_exit(self, controller):
        print "Exited"

    # Handle motions while looping through frames
    def on_frame(self, controller):

        frame = controller.frame()

        #hand = frame.hands[0]

        # Open Spotify/Close
        if len(frame.hands) == 2 and (time.time() - self.prevTime) > self.minTime:
            self.prevTime = time.time()
            if(math.fabs(frame.hands[0].grab_strength - frame.hands[1].grab_strength) > 0.95):
                print "Close Spotify"
                os.system("TASKKILL /F /IM Spotify.exe")
            else:
                print "Open Spotify"
                os.system("start Spotify.lnk")

        for gesture in frame.gestures():

            if gesture.type == Leap.Gesture.TYPE_SWIPE:
                swipe = SwipeGesture(gesture)
                direction = swipe.direction

                if (direction.x > 0 and math.fabs(direction.x) > math.fabs(direction.y) and (time.time() - self.prevTime) > self.minTime):
                    win32api.keybd_event(VK_MEDIA_NEXT_TRACK, code2)
                    self.prevTime = time.time()
                    print "Next Track, and the speed is %f" % (swipe.speed)

                #elif gesture.state is Leap.Gesture.STATE_START and direction.x < 0 and math.fabs(direction.x) > math.fabs(direction.y):
                elif (direction.x < 0 and math.fabs(direction.x) > math.fabs(direction.y) and (time.time() - self.prevTime) > self.minTime):
                    win32api.keybd_event(VK_MEDIA_PREV_TRACK, code3)
                    self.prevTime = time.time()
                    print "Previous Track"

                #elif gesture.state is Leap.Gesture.STATE_START and direction.y < 0 and math.fabs(direction.y) > math.fabs(direction.x):
                elif (direction.y < 0 and math.fabs(direction.y) > math.fabs(direction.x)):
                    win32api.keybd_event(VK_VOLUME_DOWN, code5)
                    print "Volume Down"

                #elif gesture.state is Leap.Gesture.STATE_START and direction.y > 0 and math.fabs(direction.y) > math.fabs(direction.x):
                elif (direction.y > 0 and math.fabs(direction.y) > math.fabs(direction.x)):
                    win32api.keybd_event(VK_VOLUME_UP, code6)
                    print "Volume Up"

 
            elif gesture.type == Leap.Gesture.TYPE_KEY_TAP and (time.time() - self.prevTime) > self.minTime:
                tap = KeyTapGesture(gesture)
                win32api.keybd_event(VK_MEDIA_PLAY_PAUSE, code1)
                self.prevTime = time.time()
                print "Play/Pause Track"

            # Screen tap to mute/unmute
            elif gesture.type == Leap.Gesture.TYPE_SCREEN_TAP and (time.time() - self.prevTime) > self.minTime:
                tapping = ScreenTapGesture(gesture)
                win32api.keybd_event(VK_VOLUME_MUTE, code4)
                self.prevTime = time.time()
                print "Mute/Unmute"


def main():


    listener = LeapEventListener()

    controller = Leap.Controller()

    controller.add_listener(listener)


    def swipeClick():
        novi = Toplevel()
        novi.title("Next Song Gesture: Swipe Right")
        canvas = Canvas(novi, width=500, height=500)
        canvas.pack(expand=YES, fill = BOTH)
        swipe = PhotoImage(file = "swipe.gif")
        canvas.create_image(50, 100, image = swipe, anchor = NW)
        canvas.swipe = swipe
        
    def keyClick():
        novi = Toplevel()
        novi.title("Play/Pause Gesture: Key Press")
        canvas = Canvas(novi, width=500, height=500)
        canvas.pack(expand=YES, fill = BOTH)
        swipe = PhotoImage(file = "keyTap.gif")
        canvas.create_image(50, 100, image = swipe, anchor = NW)
        canvas.swipe = swipe
        
    def screenClick():
        novi = Toplevel()
        novi.title("Mute/Unmute: Screen Tap")
        canvas = Canvas(novi, width=500, height=500)
        canvas.pack(expand=YES, fill = BOTH)
        swipe = PhotoImage(file = "screentap.gif")
        canvas.create_image(50, 100, image = swipe, anchor = NW)
        canvas.swipe = swipe
        
    def swipeClickPrev():
        novi = Toplevel()
        novi.title("Previous Song Gesture: Swipe Left")
        canvas = Canvas(novi, width=500, height=500)
        canvas.pack(expand=YES, fill = BOTH)
        swipe = PhotoImage(file = "swipe.gif")
        canvas.create_image(50, 100, image = swipe, anchor = NW)
        canvas.swipe = swipe

    window = Tk()
    window.geometry("500x500")
    window.title("M.I.G.I.")
    spc = Label(window, text = " ")
    lbl2 = Label(window, text = "M.I.G.I.: Music Interaction and Gesture Identification", font = ("Arial Bold", 10), anchor = "w")
    desc = Label(window, text = "MIGI uses Leap motion gesture identification and the Windows API to ", anchor = "w")
    desc2 = Label(window, text = "give users an easy, intuitive way to interact with their media player.", anchor = "w") 
    lbl1 = Label(window, text = "Gesture Interaction Options", font = ("Arial Bold", 10), anchor = "w")
    nextSong = Button(window, text = "Next Song", command = swipeClick, anchor = "w")
    previous = Button(window, text = "Previous Song", command = swipeClickPrev, anchor = "w")
    play = Button(window, text = "Play/Pause", command = keyClick, anchor = "w")
    mute = Button(window, text = "Mute/Unmute", command = screenClick, anchor = "w")


    lbl2.grid(column=0, row=0)
    desc.grid(column=0, row=1)
    desc2.grid(column=0, row=2)
    spc.grid(column = 0, row = 3)
    spc.grid(column = 0, row = 4)
    spc.grid(column = 0, row = 5)

    lbl1.grid(column=0, row=5)
    nextSong.grid(column = 0, row = 6)
    #swipePic.grid(column = 2, row = 1)
    previous.grid(column = 0, row = 7)
    #swipePic.grid(column = 2, row= 2)
    play.grid(column = 0, row = 8)
    #keyTap.grid(column = 2, row = 3)
    mute.grid(column = 0, row = 9)
    #screenTap.grid(column = 2, row = 4)
    window.mainloop()

    print "Press Enter to quit"

    try:
        sys.stdin.readline()
    except KeyboardInterrupt:
        pass
    finally:
        controller.remove_listener(listener)

if __name__ == "__main__":
    main()
