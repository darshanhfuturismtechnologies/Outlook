# import os
# import subprocess
# import time
#
#
# class ScreenRecorder:
#     def __init__(self, filename, project_folder, system_folder, fps=15):
#         self.process_system = None
#         self.filename = filename                                    #The output video file name
#         self.project_folder = project_folder                        #The PyCharm project folder
#         self.system_folder = system_folder                          #system folder
#         self.fps = fps
#         self.process = None                                        # Will hold the running FFmpeg process
#
#
#
#         # Ensure both folders exist, create them if not
#         if not os.path.exists(self.project_folder):
#             os.makedirs(self.project_folder)
#         if not os.path.exists(self.system_folder):
#             os.makedirs(self.system_folder)
#
#
#
#     def start(self):
#         ffmpeg_path = r"C:\Users\swapnalik\Downloads\FFmpeg\ffmpeg-master-latest-win64-gpl-shared\bin\ffmpeg.exe"
#
#         #Path to save the video in the PyCharm project folder
#         project_filepath = os.path.join(self.project_folder, self.filename)
#         #Path to save the video in the system folder
#         system_filepath = os.path.join(self.system_folder, self.filename)
#
#         # ffmpeg command to record the desktop
#         cmd = [
#             ffmpeg_path,
#             "-y",                                                         #Overwrite output file if it already exists
#             "-f", "gdigrab",                                              #Grab the screen on Windows
#             "-framerate", str(self.fps),                                  #Recording speed
#             "-i", "desktop",                                              #Capture the entire desktop
#             "-pix_fmt", "yuv420p",                                        #Ensures compatibility with most players
#             project_filepath                                              #Output file name (project folder)
#         ]
#
#         # Launching FFmpeg
#         # ubprocess.Popen starts FFmpeg as a background process.
#         # stdout and stderr are piped so that you can later read FFmpeg’s logs or errors.
#         # stdin is piped so you can send input to FFmpeg (like the 'q' command to stop recording).
#         self.process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, stdin=subprocess.PIPE)
#         time.sleep(1)
#
#         # Also save the recording to the system folder
#         cmd[10] = system_filepath                                     #Change the path to the system folder
#         self.process_system = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,stdin=subprocess.PIPE)
#         time.sleep(1)
#
#
#
#     def stop(self):
#         if self.process:
#             try:
#                 # Send 'q' to FFmpeg via stdin to stop recording gracefully
#                 self.process.communicate(input=b'q', timeout=10)
#                 self.process_system.communicate(input=b'q', timeout=10)
#             except subprocess.TimeoutExpired:
#                 # If FFmpeg does not stop in 10 seconds, force kill it
#                 self.process.kill()
#                 self.process_system.kill()
#             finally:
#                 self.process = None
#                 self.process_system =None





import os
import subprocess
import shutil
import ctypes

def get_primary_screen_size():
    user32 = ctypes.windll.user32                              #Access Windows system functions.
    user32.SetProcessDPIAware()                                #Ensures correct screen resolution is returned
    width = user32.GetSystemMetrics(0)                         #Returns screen width in pixels.
    height = user32.GetSystemMetrics(1)                        #Returns screen height in pixels.
    return width, height                                       #Returns a tuple (width, height) for FFmpeg to know the screen size.

class ScreenRecorder:
    def __init__(self, filename, project_folder, system_folder, fps=15):
        self.filename = filename                                           #Output file name
        self.project_folder = project_folder                               #Temporory saved folder on project
        self.system_folder = system_folder                                 #backup folder for failed tests.
        self.fps = fps                                                     #Frames per second-smoothness of recording
        self.process = None                                                #placeholder for the FFmpeg process.

        os.makedirs(self.project_folder, exist_ok=True)                     #Creates the folders if they don’t exist.
        os.makedirs(self.system_folder, exist_ok=True)                      #exist_ok=True prevents errors if the folder already exists.


        #Full path to FFmpeg executable.
        self.ffmpeg_path = r"C:\Users\swapnalik\Downloads\FFmpeg\ffmpeg-master-latest-win64-gpl-shared\bin\ffmpeg.exe"
        #temporary recording in MKV format (less likely to corrupt if FFmpeg is stopped abruptly).
        self.temp_file = os.path.join(self.project_folder, f"{self.filename}.mkv")
        #Join the final MP4 path in the project folder.
        self.final_file_project = os.path.join(self.project_folder, self.filename)
        #final MP4 path in the system folder.
        self.final_file_system = os.path.join(self.system_folder, self.filename)

    def start(self):
        #Get the current screen size
        width, height = get_primary_screen_size()
        #ffmpeg command to record the desktop
        cmd = [
            self.ffmpeg_path,
            "-y",                                               #overwrite existing file.
            "-f", "gdigrab",                                    #grabs the Windows desktop.
            "-framerate", str(self.fps),                        #FPS for the video.
            "-video_size", f"{width}x{height}",                 #width x height of the screen.
            "-i", "desktop",                                    #input is the full desktop.
            "-pix_fmt", "yuv420p",                              #ensures compatibility with most players.
            self.temp_file                                      #Output file
        ]

        # Launching FFmpeg
        # ubprocess.Popen starts FFmpeg as a background process.
        # stdout and stderr are piped so that you can later read FFmpeg’s logs or errors.
        # stdin is piped so you can send input to FFmpeg (like the 'q' command to stop recording gracefully).
        self.process = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, stdin=subprocess.PIPE)
        print(f"[INFO] Recording started: {self.temp_file}")


    #Sends q to FFmpeg via stdin:-tells FFmpeg to stop and finish writing the file properly.
    #timeout=30 → waits up to 30 seconds for FFmpeg to close.
    #If FFmpeg doesn’t stop in 30s then kill() it forcefully.
    #Sets self.process to None to clean up.

    def stop(self, save=False):
        if self.process:
            try:
                # Send 'q' to stop recording gracefully
                self.process.communicate(input=b'q', timeout=30)
            except subprocess.TimeoutExpired:
                self.process.kill()
            self.process = None

        #Save or discard recording
        #If save=True (test failed), then:
        #Convert MKV:-MP4 without re-encoding (-c copy).
        #Copy final MP4 to system folder.
        #Delete temp MKV file.
        #Print success message.
        #Or If save=False (test passed):
        #Delete temp .mkv.
        #No MP4 is saved.

        if save and os.path.exists(self.temp_file):
            # Convert MKV filr to MP4
            convert_cmd = [
                self.ffmpeg_path,
                "-y",
                "-i", self.temp_file,
                "-c", "copy",
                self.final_file_project
            ]
            subprocess.run(convert_cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            shutil.copy2(self.final_file_project, self.final_file_system)
            os.remove(self.temp_file)
            print(f"[INFO] Recording saved: {self.final_file_project}")

            #If test passed, simply delete the temp MKV file.
            #Ensures no unnecessary videos are kept.

        elif os.path.exists(self.temp_file):
            os.remove(self.temp_file)
            print(f"[INFO] Recording discarded (test passed)")



















































# import os
# import subprocess
# import time
#
# class ScreenRecorder:
#     def __init__(self, filename, project_folder, system_folder, fps=15, ffmpeg_path=None):
#         self.ffmpeg_path =r"C:\Users\swapnalik\Downloads\FFmpeg\ffmpeg-master-latest-win64-gpl-shared\bin\ffmpeg.exe"
#         self.filename = filename
#         self.project_folder = project_folder
#         self.system_folder = system_folder
#         self.fps = fps
#
#         # Ensure folders exist
#         os.makedirs(self.project_folder, exist_ok=True)
#         os.makedirs(self.system_folder, exist_ok=True)
#
#         self.project_filepath = os.path.join(self.project_folder, self.filename)
#         self.system_filepath  = os.path.join(self.system_folder,  self.filename)
#         self.process = None
#         self.process_system = None
#
#     def start(self):
#         # Build command for project folder
#         cmd1 = [
#             self.ffmpeg_path,
#             "-y",
#             "-f", "gdigrab",                       # capture desktop on Windows :contentReference[oaicite:2]{index=2}
#             "-framerate", str(self.fps),
#             "-i", "desktop",
#             "-pix_fmt", "yuv420p",
#             self.project_filepath
#         ]
#         # Launch first process
#         self.process = subprocess.Popen(cmd1, stdout=subprocess.PIPE, stderr=subprocess.PIPE, stdin=subprocess.PIPE)
#
#         # Build second command for system folder (same input, different output)
#         cmd2 = cmd1.copy()
#         cmd2[-1] = self.system_filepath
#         self.process_system = subprocess.Popen(cmd2, stdout=subprocess.PIPE, stderr=subprocess.PIPE, stdin=subprocess.PIPE)
#
#         # Give a little time to stabilize
#         time.sleep(1)
#
#     def stop(self):
#         # Stop first
#         if self.process:
#             try:
#                 self.process.communicate(input=b"q", timeout=10)
#             except subprocess.TimeoutExpired:
#                 self.process.kill()
#             finally:
#                 self.process = None
#
#         # Stop second
#         if self.process_system:
#             try:
#                 self.process_system.communicate(input=b"q", timeout=10)
#             except subprocess.TimeoutExpired:
#                 self.process_system.kill()
#             finally:
#                 self.process_system = None
#
#
