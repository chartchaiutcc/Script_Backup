
import pandas as pd
from netmiko import ConnectHandler
import os
import time
import datetime
from paramiko.ssh_exception import SSHException
from netmiko.exceptions import NetMikoTimeoutException
from netmiko.exceptions import AuthenticationException
from netmiko.exceptions import SSHException
from getpass import getpass
import openpyxl
from openpyxl import load_workbook

data = pd.read_excel(r'data\02-hosts.xlsx')

ip = list(data.ip)
ios = list(data.device_type)
username = list(data.username)
password = list(data.password)

### Gat data from Excel file "02-hosts.xlsx"
container = {}
x = 0
while x < len(ip):
    container [ip[x]] = {
        'device_type':ios[x],
        'ip':ip[x],
        'username':username[x],
        'password':password[x]
    }
    x=x+1

with open(r'data\backup.txt') as d:
    lines1 = d.read().splitlines()


path = "output/"+ '{0:%Y-%m-%d %H-%M-%S}'.format(datetime.datetime.now())

if not os.path.exists(path):
    os.makedirs(path)

workbook = openpyxl.Workbook()
# Select the default sheet (usually named 'Sheet')
sheet = workbook.active
data = [
    ["IP","STATUS","REMARK"]
]
for row in data:
    sheet.append(row)

textstatus = path +'/' +'TerminalData_log.txt'
for k in container.values():
    
    name = k['ip']
    total = len(ip)
    number = ip.index(name) + 1
    terminal01 = ('{0:%Y-%m-%d %H:%M:%S.%f}'.format(datetime.datetime.now()) +' | '+'In progress : ' + str(number) + "/" + str(total) +' | connecting : '+name )
    logterminal = open(textstatus, "a")
    logterminal.write(terminal01)
    logterminal.write('\n')
    print (terminal01 )

### Failured Detection
    try :
        net_connect = ConnectHandler(**k)
    except (AuthenticationException):
        terminalAuthen = ('{0:%Y-%m-%d %H:%M:%S.%f}'.format(datetime.datetime.now()) +' | '+'FAILED      : ' + str(number) + "/" + str(total) +' | connecting : '+name + " | STATUS : Authentication failure " )
        logterminal = open(textstatus, "a")
        logterminal.write(terminalAuthen)
        logterminal.write('\n')
        print (terminalAuthen)
        data = [
        [name,"FAIL","Authentication failure"]
        ]
        for row in data:
            sheet.append(row)
        continue
    except (EOFError):
        terminalTimeout = ('{0:%Y-%m-%d %H:%M:%S.%f}'.format(datetime.datetime.now())+' | '+'FAILED      : ' + str(number) + "/" + str(total) +' | connecting : '+name + " | STATUS : Session Time out " )
        logterminal = open(textstatus, "a")
        logterminal.write(terminalTimeout)
        logterminal.write('\n')
        print (terminalTimeout)
        data = [
        [name,"FAIL","Session Time out"]
        ]
        for row in data:
            sheet.append(row)
        continue
    except (SSHException):
        terminalSSH = ('{0:%Y-%m-%d %H:%M:%S.%f}'.format(datetime.datetime.now()) +' | '+'FAILED      : ' + str(number) + "/" + str(total) +' | connecting : '+name + " | STATUS : SSH Issue. TCP connection to device failed " )
        logterminal = open(textstatus, "a")
        logterminal.write(terminalSSH)
        logterminal.write('\n')
        print (terminalSSH)
        data = [
        [name,"FAIL","SSH Issue. TCP connection to device failed"]
        ]
        for row in data:
            sheet.append(row)
        continue
    except Exception as unknown_error:
        terminalError = ('{0:%Y-%m-%d %H:%M:%S.%f}'.format(datetime.datetime.now()) +' | '+'FAILED      : ' + str(number) + "/" + str(total) +' | connecting : '+name + " | STATUS : Unknown_ERROR " )
        logterminal = open(textstatus, "a")
        logterminal.write(terminalError)
        logterminal.write('\n')
        print (terminalError)
        print ('Some other error: ' + str(unknown_error))
        data = [
        [name,"FAIL","Unknown_ERROR",+ str(unknown_error)]
        ]
        for row in data:
            sheet.append(row)
        continue

    data = [
    [name,"Successed"]
    ]
    for row in data:
        sheet.append(row)
    
    hostname = net_connect.send_command('show run | i hostname')
    hostname.split(" ")
    hostname,device = hostname.split(" ")
   
    ### Create text file
    filename = path +'/' + device + "-" + name +'.txt'
    
    for m in lines1 :
        output = net_connect.send_command(m)
        try :
            log_file = open(filename, "a")   # in append mode
        except OSError as e :
            lines02 = filename.splitlines()
            nw = lines02[0] + lines02[1]
            log_file = open(nw, "a")
            log_file.write('#####'+ m +'#####')
            log_file.write("\n")
            log_file.write(output )
            log_file.write("\n")
            log_file.write("\n")
            continue
        log_file.write('#####'+ m +'#####')
        log_file.write("\n")
        log_file.write(output )
        log_file.write("\n")
        log_file.write("\n")
    terminalFull = ('{0:%Y-%m-%d %H:%M:%S.%f}'.format(datetime.datetime.now()) +' | '+'Successed   : ' + str(number) + "/" + str(total) +' | connected  : '+name + " | STATUS : " + device + " is Complete." )
    logterminal = open(textstatus, "a")
    logterminal.write(terminalFull)
    logterminal.write('\n')
    print (terminalFull)
    
# Finally close the connection
net_connect.disconnect()
### Save Status to Excel file 
workbook.save(path +'/'+"log_Status.xlsx")

       
       
