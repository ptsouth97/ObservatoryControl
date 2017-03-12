import time
import win32com.client

ERROR = True
NOERROR = False

#######################
#   Class: cCamera    #
#######################

class cCamera:
    def __init__(self):
        print("Connecting to Maxim DL...")
        self.__CAMERA = win32com.client.Dispatch("MaxIm.CCDCamera")
        self.__CAMERA.DisableAutoShutdown = True
        try:
            self.__CAMERA.LinkEnabled = True
        except:
            print("... cannot connect to camera")
            print("--> Is camera hardware attached?")
            print("--> Is some other application already using camera hardware?")
            raise EnvironmentError('Halting program')
        if not self.__CAMERA.LinkEnabled:
            print("... camera link DID NOT TURN ON; CANNOT CONTINUE")
            raise EnvironmentError('Halting program')

    def exposeLight(self,length,filterSlot):
        print("Exposing light frame..")
        self.__CAMERA.Expose(length,1,filterSlot)
        while not self.__CAMERA.ImageReady:
            time.sleep(1)
        print("Light frame exposure and download complete!")

    def setFullFrame(self):
        self.__CAMERA.SetFullFrame()
        print("Camera set to full-frame mode")

    def setBinning(self,binmode):
        tup = (1,2,3)
        if binmode in tup:
            self.__CAMERA.BinX = binmode
            self.__CAMERA.BinY = binmode
            print("Camera binning set to {}x{}".format(binmode,binmode))
            return NOERROR
        else:
            print("ERROR: Invalid binning specified")
            return ERROR

##
##  END OF 'cCamera' Class
##

###########################
#    cCAMERA UNIT TEST    #
###########################

if __name__ == "__main__":

    # Create an instance of the cCamera class
    testCamera = cCamera()

    # Setup MaxIm DL to take a full frame image
    testCamera.setFullFrame()
    # Setup binning for 2x2
    if not testCamera.setBinning(2):
        # Expose filter slot 2 (Blue) for 12.5 seconds
        testCamera.exposeLight(12.5,2)
    else:
        print("Image not take due to previous error")