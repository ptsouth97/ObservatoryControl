import time
import os
import win32com.client

LIGHT_PATH = r"C:\Users\Blake\Documents\astro_images\\"
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

    def generateFilename(self,path,baseName):
        # path is the path to where the file will be saved
        baseName.replace(':', '_')      # colons become underscores
        baseName.replace(' ', '_')      # blanks become underscores
        baseName.replace('\\', '_')     # backslash becomes underscore
        # make sure the base filename has an '_' at the end
        if not baseName.endswith("_"):
            baseName = baseName + "_"
        # add 1 to use next available number
        seqMax = self.getSequenceNumber(path,baseName)
        seqNext = seqMax + 1
        filename = "%s%s%05d.fit" % (path,baseName,seqNext)
        return filename

    def getSequenceNumber(self,path,baseName):
        # get a list of files in the image directory
        col = os.listdir(path)
        # Loop over these filenames and see if any match the basename
        retValue = 0
        for name in col:
            front = name[0:-9]
            back = name[-9:]
            if front == baseName:
                # baseName match found, now get sequence number for this file
                seqString = name[-9:-4] # get last 5 chars of name (seq number)
                try:
                    seqInt = int(seqString)
                    if seqInt > retValue:
                        retValue = seqInt       # store greatest sequence number
                except:
                    pass
        return retValue

    def exposeLight(self,length,filterSlot,name):
        print("Exposing light frame..")
        self.__CAMERA.Expose(length,1,filterSlot)
        while not self.__CAMERA.ImageReady:
            time.sleep(1)
        print("Light frame exposure and download complete!")
        # save image
        filename = self.generateFilename(LIGHT_PATH,name)
        print("Saving light image -> %s" % filename)
        self.__CAMERA.SaveImage(filename)

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
        for i in range(4):
            # Expose filter slot 0 (Red) for 15 seconds
            testCamera.exposeLight(15,0,'m51_R_2x2')
    else:
        print("Image not take due to previous error")