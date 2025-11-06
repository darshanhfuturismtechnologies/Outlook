
import subprocess
import time

class ScreenRecorder:
    # Initialize the recorder with a file name and frames per second
    def __init__(self, filename="record.mp4",fps=15):
        self.filename = filename                                 #The output video file name
        self.fps = fps                                           #Frames per second for recording
        self.process = None

    def start(self):
        ffmpeg_path = r"C:\Users\swapnalik\Downloads\FFmpeg\ffmpeg-master-latest-win64-gpl-shared\bin\ffmpeg.exe"
        #ffmpeg command to record the desktop
        cmd = [
            ffmpeg_path,
            "-y",                                               #Overwrite output file if it already exists
            "-f", "gdigrab",                                    #Grab the screen on Windows
            "-framerate", str(self.fps),                        #Set recording FPS
            "-i", "desktop",                                    #Input source:the entire desktop
            "-pix_fmt", "yuv420p",                              #Set pixel format to a widely compatible format
            self.filename                                       #Output file name
        ]

        # Start the ffmpeg process
        # stdout and stderr are piped so we can see errors if needed
        # stdin is piped so we can send 'q' to stop recording gracefully
        self.process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, stdin=subprocess.PIPE)
        time.sleep(1)  #Wait for initialize ffmpeg properly.

    def stop(self):
        # Only try to stop if the process exists
        if self.process:
            try:
                #Send 'q' to ffmpeg via stdin to stop recording gracefully
                #This allows ffmpeg to finalize the MP4 file properly
                self.process.communicate(input=b'q', timeout=10)
            except subprocess.TimeoutExpired:
                #If ffmpeg does not stop in 10 seconds, force kill it
                self.process.kill()
                #Reset the process to None so it can be started again
            self.process = None

