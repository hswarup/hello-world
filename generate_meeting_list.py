# “working” POC process to - schedule/run/communicate and track the sessions for all  employees - with the following parameters:
# 3 in each session
# has to have both representations from 2 countries
# no same people meet in future meetings
# Utilise all the Roles  (a list will be supplied – use a sample list in your POC)
# Sends an email out with a Teams link to each coffee catchup
# Microsoft Teams/Excel/Outlook compatible 

import random
import string
from pprint import pprint
from datetime import datetime,timedelta
import xlsxwriter
import time

#workbook=xlsxwriter.Workbook("/Users/swaruph/OneDrive - gmail/BOET/Portfolio Analytics/Coffee_Catchup_POC/v1_Random_small_Virtual_Coffee_Invite_list.xlsx")
#workbook=xlsxwriter.Workbook("/Users/swaruph/OneDrive - gmail/BOET/Portfolio Analytics/Coffee_Catchup_POC/PortfolioAnalytics_Virtual_Coffee_Invite_list.xlsx")
workbook=xlsxwriter.Workbook("/Users/swaruph/Desktop/CMDA_Virtual_Coffee_Invite_list.xlsx")
worksheet1=workbook.add_worksheet()

#Logging
debug=False    #True or False
phase_starttime = time.time()   

email_list_mel = []
email_list_gcc = []
email_list_all = []
meeting_combo = []
startdate=datetime(2024, 5,8,13,30,0)         #Start day of meeting(Week 0)"
# end_date="2024-04-10"                         #End day of meeting(Week 0)
# start_time="T14:00:00"                        #Start time of meeting
# end_time="T14:30:00"                          #End time of meeting
timezone="AUS Eastern Standard Time"          #Timezone of sender
calendar_id = "Calendar"                      #Calendar ID of sender
message = "Let us get to know each other!"    #Body of meeting invite 
subject = "CD&A Coffee Catchup"     #Subject of meeting invite
for i in range(1, 11):        ## Bump up the value to generate more email addresses
    email = ''.join(random.choices(string.ascii_lowercase + string.digits, k=8)) + '@gmail.com'
    if i % 2 == 0:
        region = 'GCC'
        email_list_gcc.append({'email': 'location1_'+email, 'region': region})
    else:
        region = 'MEL'
        email_list_mel.append({'email': 'location2_'+email, 'region': region})

email_list_all = email_list_gcc+email_list_mel

print ("--------------------------------------------------")
print ("{} people in email_list_mel".format(len(email_list_mel)))
print (email_list_mel) if debug==True else None
print ("--------------------------------------------------")
print ("{} people in email_list_gcc".format(len(email_list_gcc)))
print (email_list_gcc) if debug==True else None
print ("--------------------------------------------------")
print ("{} people in email_list_all".format(len(email_list_all)))
print (email_list_all) if debug==True else None
phase_endtime = time.time()   
print ("Phase 1: Data Preparation took {} seconds".format(phase_endtime-phase_starttime))
phase_starttime=phase_endtime
#print (random.choice(email_list_mel))
#print (random.choice(email_list_gcc))
#meeting_combo.append({'email1':random.choice(email_list_mel), 'email2':random.choice(email_list_gcc), 'email3':random.choice(email_list_all)})
#print (meeting_combo)
loop_counter=0
while loop_counter < 3:      ## Bump up the value to get maximum combinations
    for i in range(len(email_list_mel)):
        for j in range(len(email_list_gcc)):
            current_combo=()
            email3 = random.choice(email_list_all)['email']
            #print ("{} {} {}".format(email_list_mel[i]['email'],email_list_gcc[j]['email'],email3))
            #for k in range(len(email_list_all)):
                #meeting_combo.append({'email1':email_list_mel[i]['email'],'email2':email_list_gcc[j]['email'],'email3':email_list_all[k]['email']})
            if email3!=email_list_mel[i]['email'] and email3!=email_list_gcc[j]['email']:
                current_combo = (email_list_mel[i]['email'],email_list_gcc[j]['email'],email3)
            if current_combo not in meeting_combo and len(current_combo) > 1:
                meeting_combo.append(current_combo)
    loop_counter+=1
print ("--------------------------------------------------")
print ("{} items in meeting_combo".format(len(meeting_combo)))
#pprint (meeting_combo)
print ("--------------------------------------------------")
phase_endtime = time.time()   
print ("Getting initial meeting combinations took {} seconds".format(phase_endtime-phase_starttime))
phase_starttime=phase_endtime

exit

### Identify duplicate combinations
dup_meeting_combo = []
for l in range(len(meeting_combo)):
    for m in range(len(meeting_combo)):
        if l != m :
            emails_match_count=0
            if meeting_combo[l] in dup_meeting_combo:
                #print ("{} already removed".format(meeting_combo[l]))
                break
            #print ("checking for {} and {}".format(meeting_combo[l],meeting_combo[m]))
            if meeting_combo[l][0] == meeting_combo[m][0] or meeting_combo[l][0] == meeting_combo[m][1] or meeting_combo[l][0] == meeting_combo[m][2]:
                #print ("{}=={}".format(meeting_combo[l][0],meeting_combo[m][0]))
                #print ("{}=={}".format(meeting_combo[l][0],meeting_combo[m][0]))
                #print ("{}=={}".format(meeting_combo[l][0],meeting_combo[m][0]))
                #print ("One email matches")
                emails_match_count+=1
            if meeting_combo[l][1] == meeting_combo[m][0] or meeting_combo[l][1] == meeting_combo[m][1] or meeting_combo[l][1] == meeting_combo[m][2]:
                #print ("Two emails matches")
                emails_match_count+=1
            if meeting_combo[l][2] == meeting_combo[m][0] or meeting_combo[l][2] == meeting_combo[m][1] or meeting_combo[l][2] == meeting_combo[m][2]:
                #print ("ALL emails matches")
                emails_match_count+=1
            #print ("emails_match_count: {}".format(emails_match_count))
            if emails_match_count > 1 :
                #print ("{} to be removed".format(meeting_combo[m]))
                dup_meeting_combo.append(meeting_combo[m])
print ("--------------------------------------------------")
print ("{} items in dup_meeting_combo".format(len(dup_meeting_combo)))
#pprint (dup_meeting_combo)
phase_endtime = time.time()   
print ("Identifying duplicate combinations took {} seconds".format(phase_endtime-phase_starttime))
phase_starttime=phase_endtime


### Eliminate duplicate combinations
unique_meeting_combo  = list(set(meeting_combo)-set(dup_meeting_combo))
print ("--------------------------------------------------")
print ("{} items in unique_meeting_combo".format(len(unique_meeting_combo)))
#pprint (unique_meeting_combo)

phase_endtime = time.time()   
print ("Eliminating duplicate combinations took {} seconds".format(phase_endtime-phase_starttime))
phase_starttime=phase_endtime

### Allocate meetings to weeks
unique_meetings_with_weeks={}
person_occurences={}
person_schedule={}
#person_schedule_list=[]
person_schedule_list2=[]
for p in range(len(unique_meeting_combo)):
    #print ("--------------------------------------------------")
    #print ('Processing <<<<<<<<<{}>>>>>>>>>>>>>>'.format(unique_meeting_combo[p]))
    #print ("--------------------------------------------------")
    if p==0:  #First meeting
        #print ("---------------First meeting---------------")
        person_schedule[unique_meeting_combo[p][0]]={'email':unique_meeting_combo[p][0],'week'+str(p):unique_meeting_combo[p],'max_assigned_week':p}
        person_schedule[unique_meeting_combo[p][1]]={'email':unique_meeting_combo[p][1],'week'+str(p):unique_meeting_combo[p],'max_assigned_week':p}
        person_schedule[unique_meeting_combo[p][2]]={'email':unique_meeting_combo[p][2],'week'+str(p):unique_meeting_combo[p],'max_assigned_week':p}
        #print (person_schedule)
        recalculated_enddate=(startdate+timedelta(minutes=30))
        recalculated_startdate_trnsfrm=startdate.strftime("%Y-%m-%dT%H:%M:%S")
        recalculated_enddate_trnsfrm=recalculated_enddate.strftime("%Y-%m-%dT%H:%M:%S")

        # person_schedule_list.append({'meeting':unique_meeting_combo[p],
        #                              'subject':subject,
        #                              'message':message,
        #                              'calendar':calendar_id,
        #                              'timezone':timezone,
        #                              'Attendee1':unique_meeting_combo[p][0],
        #                              'Attendee2':unique_meeting_combo[p][1],
        #                              'Attendee3':unique_meeting_combo[p][2],
        #                              'week':0,
        #                              'starttime':recalculated_startdate_trnsfrm,
        #                              'endtime':recalculated_enddate_trnsfrm})
        person_schedule_list2.append([','.join(unique_meeting_combo[p]),
                                     subject,
                                     message,
                                     calendar_id,
                                     timezone,
                                     unique_meeting_combo[p][0],
                                     unique_meeting_combo[p][1],
                                     unique_meeting_combo[p][2],
                                     0,
                                     recalculated_startdate_trnsfrm,
                                     recalculated_enddate_trnsfrm])
    else:     #Subsequent meetings
        #print ("---------------Subsequent meetings------------------")
        if person_schedule.get(unique_meeting_combo[p][0]) is None:  #First person in the meeting
            #print ("Max assigned week for {} is None".format(unique_meeting_combo[p][0]))
            person1_max=-1
        else:
            #print ('Max assigned week for {}...{}'.format(person_schedule.get(unique_meeting_combo[p][0]).get('email'),person_schedule.get(unique_meeting_combo[p][0]).get('max_assigned_week')))
            person1_max=person_schedule.get(unique_meeting_combo[p][0]).get('max_assigned_week')
        if person_schedule.get(unique_meeting_combo[p][1]) is None:  #Second person in the meeting
            #print ("Max assigned week for {} is None".format(unique_meeting_combo[p][1]))
            person2_max=-1
        else:
            #print ('Max assigned week for {}...{}'.format(person_schedule.get(unique_meeting_combo[p][1]).get('email'),person_schedule.get(unique_meeting_combo[p][1]).get('max_assigned_week')))
            person2_max=person_schedule.get(unique_meeting_combo[p][1]).get('max_assigned_week')
        if person_schedule.get(unique_meeting_combo[p][2]) is None:  #Third person in the meeting
            #print ("Max assigned week for {} is None".format(unique_meeting_combo[p][2]))
            person3_max=-1
        else:
            #print ('Max assigned week for {}...{}'.format(person_schedule.get(unique_meeting_combo[p][2]).get('email'),person_schedule.get(unique_meeting_combo[p][2]).get('max_assigned_week')))
            person3_max=person_schedule.get(unique_meeting_combo[p][2]).get('max_assigned_week')
        max_week_so_far = max(person1_max,person2_max,person3_max)
        #print('Max. week scheduled so far is {}...Thus this meeting can be scheduled for week {}'.format(max_week_so_far,max_week_so_far+1))
        #print ("--------------------------------------------------")
        #print ("--------------------------------------------------")
        #print ("--------------------------------------------------")
        person_schedule[unique_meeting_combo[p][0]]={'email':unique_meeting_combo[p][0],'week'+str(max_week_so_far+1):unique_meeting_combo[p],'max_assigned_week':max_week_so_far+1}
        person_schedule[unique_meeting_combo[p][1]]={'email':unique_meeting_combo[p][1],'week'+str(max_week_so_far+1):unique_meeting_combo[p],'max_assigned_week':max_week_so_far+1}
        person_schedule[unique_meeting_combo[p][2]]={'email':unique_meeting_combo[p][2],'week'+str(max_week_so_far+1):unique_meeting_combo[p],'max_assigned_week':max_week_so_far+1}
        recalculated_startdate=(startdate+timedelta(days=7*(max_week_so_far+1)))
        recalculated_enddate=(recalculated_startdate+timedelta(minutes=30))
        recalculated_startdate_trnsfrm=recalculated_startdate.strftime("%Y-%m-%dT%H:%M:%S")
        recalculated_enddate_trnsfrm=recalculated_enddate.strftime("%Y-%m-%dT%H:%M:%S")
        # person_schedule_list.append({'meeting':unique_meeting_combo[p],
        #                              'subject':subject,
        #                              'message':message,
        #                              'calendar':calendar_id,
        #                              'timezone':timezone,
        #                              'Attendee1':unique_meeting_combo[p][0],
        #                              'Attendee2':unique_meeting_combo[p][1],
        #                              'Attendee3':unique_meeting_combo[p][2],
        #                              'week':max_week_so_far+1,
        #                              'starttime':recalculated_startdate_trnsfrm,
        #                              'endtime':recalculated_enddate_trnsfrm})

        person_schedule_list2.append([','.join(unique_meeting_combo[p]),
                                     subject,
                                     message,
                                     calendar_id,
                                     timezone,
                                     unique_meeting_combo[p][0],
                                     unique_meeting_combo[p][1],
                                     unique_meeting_combo[p][2],
                                     max_week_so_far+1,
                                     recalculated_startdate_trnsfrm,
                                     recalculated_enddate_trnsfrm])

print ("BEFORE FIX--------------------------------------------------") if debug==True else None
print ("{} meetings in person_schedule_list".format(len(person_schedule_list2))) if debug==True else None
#pprint (person_schedule_list2)
for person_schedule_list in person_schedule_list2:
    print ('Week '+str(person_schedule_list[8])+'--->',person_schedule_list[0],) if debug==True else None

phase_endtime = time.time()   
print ("Allocating meeting to weeks took {} seconds".format(phase_endtime-phase_starttime))
phase_starttime=phase_endtime

# Prepare a list of people scheduled for each week
print ("--------------------------------------------------")
print ("List of people scheduled for each week")
print ("--------------------------------------------------")
person_week_list={}
for person_schedule in person_schedule_list2:
  #if person_schedule[8]==0:
  #print (person_schedule[0])
  person_week_list['Week'+str(person_schedule[8])]=person_week_list.get('Week'+str(person_schedule[8]),',')+person_schedule[0]+','
pprint (person_week_list) if debug==True else None

phase_endtime = time.time()   
print ("Identifying people scheduled for each week took {} seconds".format(phase_endtime-phase_starttime))
phase_starttime=phase_endtime

print ("--------------------------------------------------")
print ("ENHANCING MEETING SCHEDULE")
print ("--------------------------------------------------")

for reschedule_attempts in range(1800):     #Bump up the value to optimize further
    print ("--------------------------------------------------") if debug==True else None
    print ('Updated meeting list...Reschedule Attempt#{}'.format(reschedule_attempts)) if debug==True else None
    print ("--------------------------------------------------") if debug==True else None
    for person_schedule_list in person_schedule_list2:
        print ('Week '+str(person_schedule_list[8])+'--->',person_schedule_list[0],) if debug==True else None
    print ("--------------------------------------------------") if debug==True else None
    for person_schedule_list in person_schedule_list2:
        #print ('Week '+str(person_schedule_list[8])+'--->',person_schedule_list[0],)
        restart_after_reassign=False
        for z in range(len(person_schedule_list2)):
            person_schedule=person_schedule_list2[z]
            print ('Week '+str(person_schedule[8])+'--->',person_schedule[5],person_schedule[6],person_schedule[7],)  if debug==True else None
            for week in range(0,max_week_so_far+1):
                print ('Checking if any of these folks have a meeting in week {} which has attendees {}'.format(week,person_week_list['Week'+str(week)])) if debug==True else None
                if week >= person_schedule[8] or (person_schedule[5] in person_week_list['Week'+str(week)] or person_schedule[6] in person_week_list['Week'+str(week)] or person_schedule[7] in person_week_list['Week'+str(week)]):
                    continue    # Cant reassign , Move on
                else:
                    print ("Reassigning this meeting from week {} to week {}".format(person_schedule[8],week)) if debug==True else None
                    print ("Remove {} from {}".format(person_schedule[0],person_week_list['Week'+str(person_schedule[8])])) if debug==True else None
                    person_week_list['Week'+str(person_schedule[8])] = person_week_list['Week'+str(person_schedule[8])].replace(person_schedule[0],'')       # Remove the persons from the later week  
                    person_week_list['Week'+str(week)] = person_week_list.get('Week'+str(week),',')+person_schedule[0]+','                                   # Add the persons in the earlier week
                    print('REVISED : List of people scheduled for each week') if debug==True else None
                    pprint (person_week_list) if debug==True else None
                    person_schedule_list2[z][8]=week    #Reassign to an earlier week
                    restart_after_reassign=True
                    break
            if restart_after_reassign==True:
                break
        if restart_after_reassign==True:
            break

# # Print the revised meeting schedule
print ("--------------------------------------------------")
print ("AFTER FIX - FINAL SCHEDULE")
print ("--------------------------------------------------")

print ("{} meetings in person_schedule_list".format(len(person_schedule_list2)))
#pprint (person_schedule_list2)
for person_schedule_list in person_schedule_list2:
    print ('Week '+str(person_schedule_list[8])+'--->',person_schedule_list[0],)

print ("--------------------------------------------------")
print ("FINAL - List of people scheduled for each week")
print ("--------------------------------------------------")
pprint (person_week_list)

phase_endtime = time.time()   
print ("Enhancing meeting schedule took {} seconds".format(phase_endtime-phase_starttime))
phase_starttime=phase_endtime

caption = "Commercial Data Virtual Coffee Catchup Teams Meeting Invite List."

# Set the columns widths.
worksheet1.set_column("B:L", 22)

# Write the caption.
worksheet1.write("B1", caption)

# Add a table to the worksheet.
header_columns= [
    {"header": "Meeting"},
    {"header": "Subject"},
    {"header": "Message"},
    {"header": "Calendar Id"},
    {"header": "Time Zone"},
    {"header": "Attendee1"},
    {"header": "Attendee2"},
    {"header": "Attendee3"},
    {"header": "MeetingWeek"},
    {"header": "Start Time"},
    {"header": "End Time"}]

worksheet1.add_table("B3:L2000", {"data": person_schedule_list2,"columns":header_columns})

# for i in range(len(person_schedule_list)):
#     #print (person_schedule_list[i]['subject'])
#     worksheet1.write_row("C"+str(i+4), person_schedule_list[i]['subject'])
#     worksheet1.write_row("D"+str(i+4), person_schedule_list[i]['message'])
#     worksheet1.write_row("E"+str(i+4), person_schedule_list[i]['meeting'])
workbook.close()

phase_endtime = time.time()   
print ("Writing to Excel took {} seconds".format(phase_endtime-phase_starttime))
