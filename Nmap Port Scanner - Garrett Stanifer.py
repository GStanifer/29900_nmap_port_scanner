# Garrett Stanifer
# CIT 29900
# Nmap Port Scanner
#  !!! Run as admin for os_detection to work !!!
################################################################################################################################################

#!/usr/bin/env python3

import nmap3
import sqlite3
import xlwt

nmap = nmap3.Nmap()

data = sqlite3.connect('database.db')

class Hosts:
    def __init__(self, hostName, IP, hostType) -> None:
        self.hostName = hostName
        self.IP = IP
        self.hostType = hostType
  
    def insertDatabase(self):
        try:
            data.execute(''' insert into Hosts(name, ip, scan_data, os, accuracy, os_type) values (?,?,?,?,?,?)''', (host.hostName, host.IP, str(results[first]['ports']), str(os_results[first]['osmatch'][0]['name']), str(os_results[first]['osmatch'][0]['osclass']['accuracy']), str(os_results[first]['osmatch'][0]['osclass']['type'])) )
        except IndexError:
            data.execute(''' insert into Hosts(name, ip, scan_data, os, accuracy, os_type) values (?,?,?,?,?,?)''', (host.hostName, host.IP, str(results[first]['ports']), "N/A", "N/A", "N/A") )        
            
#Creates SQLtable
data.execute(''' create table if not exists Hosts(
    id INTEGER primary key autoincrement,
    name text not null,
    ip varchar(15),
    scan_data text not null,
    os text not null,
    accuracy text not null,
    os_type text not null
 );''')

#Creates objects and adds some temporary value list
google = Hosts('google.com', '', 'internet')
facebook = Hosts('facebook.com', '', 'internet')
gstanifer = Hosts('gstanifer.com', '', 'internet')

hostnames = {google, facebook, gstanifer}

print("\nSTARTING NMAP PORT SCANNER...\n")

#Goes through and scans each hostname with nmap
for host in hostnames:
    #Top ports scan
    results= nmap.scan_top_ports(host.hostName)
    #OS detection scan
    os_results = nmap.nmap_os_detection(host.hostName)
   

    #Takes the host ip returned and sets it to the hostip in objects
    first = next(iter(results))
    host.IP = first

    #Takes hostname returned and stores as hostname value in objects
    host.hostName = results[first]['hostname'][0]['name']

    #Inserts data into database
    Hosts.insertDatabase(host)

    ############################
    #Prints what is happening  
    print(f"Scan on {host.hostName} at {first} complete")               
    print("\n")                
    ############################

    workbook = xlwt.Workbook()

    #Styling for the document
    style_top = xlwt.easyxf('font: bold 1; align: horiz center')
    
    #Writes headings for the data
    output = workbook.add_sheet('NMAP Output')
    output.write(0,0, "Garrett Stanifer")
    output.write(2,0, "ID", style_top )
    output.write(2,1, "Name", style_top )
    output.write(2,2, "IP", style_top )
    output.write(2,3, "Scan Data", style_top )
    output.write(2,4, "OS", style_top )
    output.write(2,5, "Accuracy", style_top )
    output.write(2,6, "OS Type", style_top)
    query = data.execute(''' select * from Hosts''')

    i=3
    #Queries the database and puts output into the excel file
    for row in query:
        output.write(i,0, row[0] )
        output.write(i,1, row[1] )
        output.write(i,2, row[2] )
        output.write(i,3, row[3] )
        output.write(i,4, row[4] )
        output.write(i,5, row[5] )
        output.write(i,6, row[6] )
        i+=1

    #Saves excel changes
    workbook.save('nmapOutput.xls')

print("All hosts successfully scanned!\n")
    
#Commits the data into database
data.commit()

#Closes connection with database
data.close()