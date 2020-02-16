#Author Dakota Lambert of 523 Tech
#Github: https://github.com/Dako7a

import wmi #accessing windows
import socket  #getting host name for the calling PC
import subprocess #pinging utility
import datetime
import smtplib #EMAIL SMTP
import ssl #EMAIL SSL
import urllib.request # gets proxy information
import os #used for printing the directory list
from email.mime.text import MIMEText  #email formatting (HTML)
from email.mime.multipart import MIMEMultipart #email formatting (HTML)

sender_email   = "email@523tech.com" #where we want the message sent from
receiver_email = "email@523tech.com" #where we want to see our message
pwd = "Password"
smtp_server = "smtp.ipage.com" #email SMTP server
port = 465 #SSL
log_file = open("log.txt", "w+")

address = '192.168.10.'  # our networking is within the .10 subnet which only requires the last octet to be altered

#packages
#pip install wmi
#pip install pywin32


def collectDriveInformation():
    """
    :params type -- NONE
    :ret         -- return a dictionary containing tuples of strings
    :desc        -- This function converts the output from the WMI API and stores it into a container for further processes.
    """

    collection_of_drives = wmi.WMI() #work with an API wrapper
    storage = {}

    for drive in collection_of_drives.Win32_LogicalDisk():
        try:  # in some instances, a drive may not have an overallsize.
            freespace = float(drive.FreeSpace) * (9.31 * 10 ** -10)
            drivespace= float(drive.Size) * (9.31 * 10 ** -10)

        except TypeError:  # The float() std function will throw an exception if a conversion is made.The float() std function will throw an exception if a conversion is made.
            continue

        storage[drive.Caption] = (freespace, drivespace, getDriveType(drive.DriveType))  #store the values in a container, rather than an object.

    return storage


def getDriveType(type):
    """
    :params type(int) -- An integer which is used for determining what kind of drive was found.
    :ret              -- a string converting an integer to something a human could understand.
    :desc             -- This function converts the output from the WMI API for a type of drive to output more information regarding the type of drive.
    """

    switcher = {
      1: "No root directory.",
      2: "Removable drive.",
      3: "Local hard disk.",
      4: "Network disk.",
      5: "Compact disk.",
      6: "RAM disk.",
    }
    return switcher.get(type, "Drive type could not be determined.")


def printDrives(thisPC):
    """
        :params dict -- a dictionary containing specific information for each drive located on the PC
            "dict[letter_of_drive] = (remaining_storage, total_storage, drive_type)"
        :ret list    -- A LIST of strings containing drive information
        :desc        -- prints the contents of the output from the WMI API
    """
    full_message = [] #container for all of our messages to easily be iterated

    log_file.write("Getting the drive information:"+'\n')
    print("Getting the drive information:")
    for drives in thisPC: #convert everything into a string to prevent any issues with
        message = "Drive Letter: %s <br>" % (str(drives))
        try: #incase there is an issue with returning one of the values, we catch the exception
            rem_storage, tot_storage, drive_type = thisPC[drives]
            ratio = round((rem_storage/tot_storage)*100, 2) #round our answer to just 2 decimal places
            message += ("Space available: %s%% <br>" % str(ratio))
            message += ("Remaining storage: %s GBs <br>" % (str(round(rem_storage, 2)))) #returns the current available space on the drive
            message += ("Total storage: %s GBs <br>" % (str(round(tot_storage, 2))))  # returns the total size of the drive
            message += ("Drive type: %s <br>" % (str(drive_type)))
        except ValueError:
            message = ("Drive value error <br>")

        full_message.append(message)
    log_file.write("Finished"+'\n')
    print("Finished")
    return full_message


def list_files_in_curdir(path):
    """
            :params path (string) -- a defined (absolute) path for the address
            :ret         -- A list of directories to be printed and emailed. (We could use yield here instead of return if we wanted to
            :desc        -- prints the contents of the provided directory at the shallowest depth
    """
    dirs = [] #container we wish to return at the end of the function

    log_file.write("Getting the folders in the current directory:"+'\n')
    print("Getting the folders in the current directory:")
    # files = os.listdir(path)
    if os.path.isdir(path):
        for directoryName in (next(os.walk(path))[1]):  #prints only the directories within the provided path
            dirs.append(directoryName)
    else:
        dirs.append(("Directory:", path, "does not exist."))

    log_file.write("Finished" +'\n')
    print("Finished")
    return dirs


def pingCameras(startingOctet, endingOctet):
    """
        :params startingOctet (int) -- the starting octet in the range for the static IP
        :params endingOctet   (int) -- the end octet in the range for the static IP
        :ret (intger)--  The number of cameras online
        :ret  (dict) -- A container holding the ipaddresses and the status of that IP address.
        :desc        -- queries the status of cameras to verify if they are up or down
    """
    cameraCount = 0
    cameraDict = {} #holds a dict for online and offline accounts mapped to an ip address

    log_file.write("Pinging cameras:"+'\n')
    print("Pinging cameras:")
    for i in range(startingOctet,endingOctet):  #our static addresses
        cur_address = address + str(i)

        pingparms = "ping -n 1 " + cur_address
        try: 
            output = subprocess.Popen(["ping","-n","1", cur_address],stdout = subprocess.PIPE, shell=True).communicate()[0] #our output should be something similar to what we see in the cmd when we ping.

            if "Destination host unreachable" in output.decode('utf-8'): #offline
                cameraDict[cur_address] = "Offline" #if we fail, we do not want to increment our cameras.
            elif "Request timed out" in output.decode('utf-8'): #offline
                cameraDict[cur_address] = "Offline" #if we fail, we do not want to increment our cameras.
            elif "Please check the name and try again" in output.decode('utf-8'): #offline   Out of range for our IP address
                cameraDict[cur_address] = "Offline" #if we fail, we do not want to increment our cameras.
            else: #not offline!
                cameraCount += 1 #we want to make sure we record the ratio for online cameras
                cameraDict[cur_address] = "Online"

        except  subprocess.CalledProcessError as err:
            cameraDict[cur_address] = "There was a problem [Debug issue]" #if we fail, we do not want to increment our cameras.
            log_file.write("error with pinging:" + err+ '\n')
    
    log_file.write("Finished"+'\n')
    print("Finished")
    return cameraCount, cameraDict


def pingNVR(nvrAddress):
    """
        :params nvrAddress (int) -- The address in which the NVR can be located
        :ret   (dict)  -- A container which stores the ip address of the NVR and the power status
        :desc        -- prints the status of cameras up or down
    """

    nvrDict = {}  # holds a dict for online and offline accounts mapped to an ip address
    
    log_file.write("Pinging NVR:"+'\n')
    print("Pinging NVR:")

    cur_address = address + str(nvrAddress)
    pingparms = "ping -n 1 " + cur_address
   
    try:  # subprocess will throw an error if the ping fails, this is because the ping packets returns a 0 value upon failure (I believe)
        output = subprocess.Popen(["ping","-n","1", cur_address],stdout = subprocess.PIPE, shell=True).communicate()[0] #our output should be something similar to what we see in the cmd when we ping.
        if "Destination host unreachable" in output.decode('utf-8'): #offline
                nvrDict[cur_address] = "Unreachable" 
        elif "Request timed out" in output.decode('utf-8'): #offline
                nvrDict[cur_address] = "Unreachable" 
        elif "Please check the name and try again" in output.decode('utf-8'): #offline Out of range for our IP address
                nvrDict[cur_address] = "Unreachable" #offline
        else: #not offline!
                nvrDict[cur_address] = "Reachable"
    except  subprocess.CalledProcessError as err:
        cameraDict[cur_address] = "There was a problem [Debug issue]" #if we fail, we do not want to increment our cameras.

    log_file.write("Finished"+'\n')
    print("Finished")
    return nvrDict

def formatEmail(subject, driveInfo, dirStructure, cameraInfo, nvrInfo, camCount, expectedCount, proxyList):
    """
            :params subject (string)    -- the host name of the hosting computer
            :params driveInfo (list)    -- a container which holds strings containing information regarding the drives
            :params dirStructure (list) -- a container which holds strings containing information regarding the directory structure
            :params cameraInfo (dict)   -- a container which holds strings containing information regarding the cameras
            :params nvrInfo (dict)      -- a container which holds strings containing information regarding the NVR
            :params camCount (int)      -- the number of online cameras
            :params proxyList           -- a list of the proxies which are enabled on the PC
            :ret     -- NONE
            :desc    -- sends email to a designated server containing information about cameras, NVR, drives and more. TODO: Fix later better comment
    """
    time_now = datetime.datetime.now()
    month = time_now.strftime("%B")
    day = time_now.strftime("%d")
    year = time_now.strftime("%Y")
    hour = time_now.strftime("%I")
    minute = time_now.strftime("%M")
    ampm = time_now.strftime("%p")

    # we draft an email which will be defaulted if the smtp server rejects our styled up page
    text = """Machine (Location):\n%s\nTime and Date:\n%s %s, %s %s:%s %s\nDrive Information:""" % (str(subject), str(month), str(day), str(year), str(hour), str(minute), str(ampm))

    for drive in driveInfo:  # iterate through the drive information
        text += str(drive.replace("<br>", '\n'))
        text += "\n"

    text += """Directory Structure:"""

    for folder in dirStructure:  # print out the list of drives
        text += str(folder)
        text += '\n'

    text += """Camera Status:"""
    for camera in cameraStatus:
        text += str(camera) + " - " + str(cameraStatus[camera])
        text += '\n'

    text += """NVR Status:"""
    for nvr in nvrStatus:
        text += (str(nvr) + " - " + str(nvrStatus[nvr])).lstrip()
        text += '\n'

    for proxy in proxyList:
        text += (str(proxyList[proxy]))
        text += "\n"

    # we draft a second email with styling as our prefered message.
    html = """\
        <html>
          <body>
            <p><b> Machine (Location):</b><br>
               %s<br> <br>
               <b>Time and Date:</b><br>
               %s %s, %s %s:%s %s <br> <br>
               <b>Drive Information:</b> <br>
        """ % (str(subject), str(month), str(day), str(year), str(hour), str(minute), str(ampm))

    for drive in driveInfo:  # print out the list of drives
        html += str(drive)
        html += "<br>"

    html += """\
                <b> Directory Structure: </b> <br>
                """

    for folder in dirStructure:  # print out the list of drives
        html += str(folder)
        html += "<br>"

    html += "<br>"

    html += """\
                <b> Camera Status: </b> <br>
                There are %s of %s camera(s) online. <br>
                """ % (camCount, expectedCount)

    for camera in cameraStatus:
        html += str(camera) + " - " + str(cameraStatus[camera])
        html += "<br>"

    html += "<br>"

    html += """\
                <b> NVR Status: </b> <br>       
                """

    for nvr in nvrStatus:
        html += (str(nvr) + " - " + str(nvrStatus[nvr]))
        html += "<br>"

    html += "<br>"

    html += """\
                <b> Proxy Status: </b> <br>       
                """

    for proxy in proxyList:
        html += (str(proxyList[proxy]))
        html += "<br>"

    html += "<br>"

    html += """\
            </p>
          </body>
        </html>
        """

    return text, html


def emailService(subject, text, html):
    """
        :params subject (string)-- the host name of the hosting computer (used in the subject of the email to idenfity the sender)
        :params text (string)   -- an unstyled string for an email message
        :params html (string)   -- a styled string for an email message
        :ret     -- NONE
        :desc    -- sends email to a designated server containing information about cameras, NVR, drives and more. TODO: Fix later better comment
    """
    #configure to,from,subject,text fields for the sent message.
    message = MIMEMultipart("alternative") # "alternative" is important because this is rendered by the browser before displaying.
    message["Subject"] = "[" + subject + "] " + "Daily Monitoring Results " + str(datetime.date.today()) #displayed
    message["From"] = subject + "<" + sender_email + ">"
    message["To"] = "Admin" + "<" + receiver_email + ">"

    # Turn these into plain/html MIMEText objects
    plain = MIMEText(text, "plain")
    html = MIMEText(html, "html")

    # The email client will try to render the last part first
    message.attach(plain)
    message.attach(html)
    log_file.write("Sending email:"+'\n')
    print("Sending email:")
    try: #if there is an issue, we would rather just catch it and prevent onscreen errors.
        context = ssl.create_default_context() #creates the SSL context
        with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
            server.login(sender_email, pwd)
            server.sendmail(sender_email, receiver_email, message.as_string())
            server.quit()

        log_file.write("Message successfully sent")
        print("Message successfully sent")
    except smtplib.SMTPException as err: #prevents the program from shutting down poorly if there was an error
        log_file.write("Error: Unable to send message")
        log_file.write(str(err))
        log_file.write("\n")
        print("Error: Unable to send message")
        print(err)
        print('\n')
    except Exception as err:
        log_file.write("Error: Unable to send message")
        log_file.write(str(err))
        log_file.write("\n")
        print("Error: Unable to send message")
        print(err)
        print('\n')

#executed upon launch
if __name__ == "__main__":   #main function to be called upon start
    expectedCount = 16 #number of cameras in the store

    hostname = socket.gethostname() #saves the name of the computer called. This should have been set by the administrator when the camera system was installed (NOT DONE FOR LAKE WORTH)
    storage = collectDriveInformation() #stores the information for all drives inserted or connected to the PC
    driveInfo = printDrives(storage) #prints information for all drives
    dirStructure = list_files_in_curdir(r"F:\DVR Backup\192.168.10.20") #prints the directories within this path
    #dirStructure = list_files_in_curdir(f"F:\DVR Backup\192.168.10.20\{curdate}") #prints the number of channels recording within our FTP server
    proxyList = urllib.request.getproxies()
    camCount, cameraStatus = pingCameras(21, 21+expectedCount) #pings cameras
    nvrStatus = pingNVR(20) #ping nvr

    text, html = formatEmail(hostname, driveInfo, dirStructure, cameraStatus, nvrStatus, camCount, expectedCount, proxyList) #format our email
    emailService(hostname, text, html) #send an email containing all of the above returned values
    log_file.close() #close our logfile
