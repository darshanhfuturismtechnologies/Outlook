
import logging

# Create a custom logger/logger instance
logger = logging.getLogger()
#Set the level to Debug so it will take all levels upto debug like error,info,warning,critical
logger.setLevel(logging.DEBUG)

#Create handlers that writes log messages to the specified file,mode=a means new log are added at the end of file
# file_handler = logging.FileHandler('Test_log.log',mode='a')
file_handler = logging.FileHandler(r"C:\Users\swapnalik\PycharmProjects\Outlook\Logger\Test_log.log",mode='a')
#Console handler that writes the log msgs on console/terminal
console_handler = logging.StreamHandler()

# Set logging level for handlers
file_handler.setLevel(logging.DEBUG)
console_handler.setLevel(logging.DEBUG)

#Formatter object specify the format of the log msg
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
#Add formatters to handlers
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

# Add handlers to the logger
logger.addHandler(file_handler)
logger.addHandler(console_handler)


















