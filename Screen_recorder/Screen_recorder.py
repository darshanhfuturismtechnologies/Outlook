import subprocess
import time

class ScreenRecorder:
    def __init__(self, filename="record.mp4",fps=15):
        self.filename = filename                                 #The output video file name
        self.fps = fps                                           #Frames per second for recording
        self.process =None                                       #Will hold the running FFmpeg process

    def start(self):
        ffmpeg_path = r"C:\Users\swapnalik\Downloads\FFmpeg\ffmpeg-master-latest-win64-gpl-shared\bin\ffmpeg.exe"
        #ffmpeg command to record the desktop
        cmd = [
            ffmpeg_path,
            "-y",                                               #Overwrite output file if it already exists
            "-f", "gdigrab",                                    #Grab the screen on Windows
            "-framerate", str(self.fps),                        #Recording speed
            "-i", "desktop",                                    #Capture the entire desktop
            "-pix_fmt","yuv420p",                               #Ensures compatibility with most players
            self.filename                                       #Output file name
        ]

        #Launching FFmpeg
        #ubprocess.Popen starts FFmpeg as a background process.
        #stdout and stderr are piped so that you can later read FFmpeg’s logs or errors.
        #stdin is piped so you can send input to FFmpeg (like the 'q' command to stop recording).
        self.process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, stdin=subprocess.PIPE)
        time.sleep(1)  #Wait for initialize ffmpeg properly.


    def stop(self):
        # Only try to stop if the process exists
        if self.process:
            try:
                #Send 'q' to ffmpeg via stdin to stop recording gracefullyThis allows ffmpeg to finalize the MP4 file properly
                self.process.communicate(input=b'q',timeout=10)
            except subprocess.TimeoutExpired:
                #If ffmpeg does not stop in 10 seconds, force kill it
                self.process.kill()
                #Reset the process to None so it can be started again
            self.process = None

